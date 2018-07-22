# Author: 林潇
# Contact: 940950943@qq.com
# Time: 2018-1-24
# UPDATE：2018-4-11
'''   定时抓取交易邮件，并根据交易信息执行动作 '''

import jqdata
import sys
import poplib
from email import parser
from email.parser import Parser
import string
import traceback
import TransactionMail
import re
import datetime
import time

#有效模型名
global models
models=['上证50ETF十五再优化5','沪深300ETF十五再优化5','沪深300ETF十五再优化6','创业板十五再优化5','证券ETF十五分钟5','券商ETF十五分钟5','创业板50十五再优化5',
        '深红利十五优化5','红利ETF十五优化5','深100ETF十五再优化5','深100ETF十五优化6','中证500ETF十五优化+期现5','中小板十五再优化5',
        '中小板十五再优化6','军工ETF十五优化5','创业ETF十五再优化5','上海国企十五优化5','中证银行十五优化5','上证180ETF十五优化5',
        '中证1000ETF十五优化5','红利基金十五优化5','小康ETF十五优化5','国债ETF十五五加十优化5','环保ETF十五优化5','消费行业十五优化5',
        '金融ETF十五优化5','消费ETF十五优化5','广发医药十五优化5','创业板E十五优化5','一带一路十五优化5','资源ETF十五优化5','医药ETF十五优化5',
        '非银ETF十五优化5','人工智能十五优化5',
        '银华锐进十五优化5','白酒B十五优化5','建信50B十五优化5',
        '国企十五改H股ETF5','国企十五改恒生ETF5','中概互联ETF十五分钟5','中国互联ETF十五分钟5','香港中小十五优化5','恒指ETF十五优化5',
        '纳指ETF十五优化5','标普500十五优化5','MSCIA股十五优化5','MSCIA股十五优化转A50','A50ETF十五优化5','AH50十五优化5','AH50精选十五再优化5','香港大盘十五再优化5']


#运行间隔时间
global run_interval
run_interval=15*60

#可接受邮件发送人
global senderList
senderList=[
'18826500406@163.com',]

#邮箱用户名
global mail_user_name
mail_user_name='940950943@qq.com'

#邮箱pop登陆授权码
global mail_password
mail_password='yilcvfhfjkowbcejsfsf'

#记录重复登陆的次数
global login_count
login_count=0

#重复登陆尝试次数
global max_login_count
max_login_count=3

#缓存已读邮件列表
global hashSet
hashSet=set()

# 初始化函数，设定基准等等
def initialize(context):
    g.security='510500.XSHG'
    # 设定510500.XSHG作为基准
    set_benchmark('510500.XSHG')
    # 开启动态复权模式(真实价格)
    set_option('use_real_price', True)
    # 过滤掉order系列API产生的比error级别低的log
    # log.set_level('order', 'error')
    global run_interval
    time_list=[]
    
    
    __now=datetime.datetime.now()
    g.trade_time_dict={
    'start_time_of_am':__now.replace(hour=9,minute=30,second=0,microsecond=0),
    'end_time_of_am':__now.replace(hour=11,minute=30,second=0,microsecond=0),
    'start_time_of_pm':__now.replace(hour=13,minute=0,second=0,microsecond=0),
    'end_time_of_pm':__now.replace(hour=15,minute=0,second=0,microsecond=0)
    }
    trade_time_dict=g.trade_time_dict
    #print(sys.version)
    interval_time=datetime.timedelta(seconds=run_interval)

    tmptime=trade_time_dict['start_time_of_am']
    tmptime=tmptime.replace(minute=tmptime.minute+1) 
    #在抓取时间点延迟一分钟后连续抓取两次，
    #效果会这样：如果交易邮件是9:45开始往邮箱里发，抓取时间分别是9:46和9:47
    while(compare_time(tmptime,trade_time_dict['end_time_of_pm'])<0):
        if ((compare_time(tmptime,trade_time_dict['start_time_of_am'])>=0 and compare_time(tmptime,trade_time_dict['end_time_of_am'])<0) or (compare_time(tmptime,trade_time_dict['start_time_of_pm'])>=0 and compare_time(tmptime,trade_time_dict['end_time_of_pm'])<0)):
            time_list.append(tmptime)
            #延迟两分钟抓取
            time_list.append(tmptime.replace(minute=tmptime.minute+1))
        tmptime=tmptime+interval_time
    #收盘提前提前两分钟
    time_list.append(trade_time_dict['end_time_of_am'].replace(hour=11,minute=28))
    time_list.append(trade_time_dict['end_time_of_am'].replace(hour=11,minute=29))
    time_list.append(trade_time_dict['end_time_of_pm'].replace(hour=14,minute=58))
    time_list.append(trade_time_dict['end_time_of_pm'].replace(hour=14,minute=59))
    time_list.append(trade_time_dict['start_time_of_am'])
    time_list.append(trade_time_dict['start_time_of_pm'])
    
    g.time_list=time_list
    print([repr(t.hour)+':'+repr(t.minute) for t in g.time_list])
    ### 股票相关设定 ###
    # 股票类每笔交易时的手续费是：买入时佣金万分之三，卖出时佣金万分之三加千分之一印花税, 每笔交易佣金最低扣5块钱
    set_order_cost(OrderCost(close_tax=0.001, open_commission=0.0003, close_commission=0.0003, min_commission=5), type='stock')
    
    ## 运行函数（reference_security为运行时间的参考标的；传入的标的只做种类区分，因此传入'000300.XSHG'或'510300.XSHG'是一样的）
      # 开盘前运行
    #run_daily(before_market_open, time='before_open', reference_security='000300.XSHG') 
      # 开盘时或每分钟开始时运行
    run_daily(market_open, time='every_bar', reference_security='000300.XSHG')
      # 收盘前2分钟
    #run_daily(after_market_close, time='close-1m', reference_security='000300.XSHG')
    #get_new_mail()
    
## 开盘时运行函数
def market_open(context):
    #global time_list
    #simple_backtest 
    #print('run type is:%s\ndatetime.now:%s\ncontext.current_date:%s\npaused?%s' % (context.run_params.type,datetime.datetime.now(),context.current_dt,get_current_data()[g.security].paused))
    #if context.run_params.type == 'sim_trade':
    now=datetime.datetime.now()
    hour=now.hour
    minute=now.minute
    for tmp in g.time_list:
        #print(hour == tmp.hour and minute == tmp.minute)
        if hour == tmp.hour and minute == tmp.minute:
            if (hour == 11 and minute == 28) or (hour == 14 and minute == 58):
                time.sleep(30)
            #初始化已读邮件列表
            update_hashvalues()
            get_new_mail(context)
            
#通过比较时间的时和分判定大小
def compare_time(date1,date2):
    
    if date1.hour == date2.hour and date1.minute == date2.minute and date1.day == date2.day:
        return 0
    if (date1.hour > date2.hour) or (date1.hour == date2.hour and date1.minute > date2.minute):
        return 1
    if (date1.hour < date2.hour) or (date1.hour == date2.hour and date1.minute < date2.minute):
        return -1
        
def is_right_time(tmptime):
    
    if ((compare_time(tmptime,g.trade_time_dict['start_time_of_am'])>=0 and compare_time(tmptime,g.trade_time_dict['end_time_of_am'])<0) 
    or (compare_time(tmptime,g.trade_time_dict['start_time_of_pm'])>=0 and compare_time(tmptime,g.trade_time_dict['end_time_of_pm'])<0)):
        return True
    return False
    
#获取新邮件
def get_new_mail(context):
    #try:
        #登陆邮箱
    has_new_mail_count=0
    complete_count=0
    messages=login_mail()
    # Concat message pieces:
    #处理中文乱码步骤一：
    newmssg=[]
    for m in messages:
        gbcharset=None
        for t in m[1]:
            charset=re.search(r'charset="*(gb([0-9]+|k))"*',t,re.I)
            if charset:
                #print(charset.group(1))
                gbcharset=charset.group(1).strip()
                break
        if gbcharset:
            print('gbcharset:%s' % (gbcharset))
            try:
                newmssg.append('\n'.join(m[1]).decode(gbcharset))
            except (UnicodeDecodeError),e:
                newmssg.append('\n'.join(m[1]))
                print(e)
    messages=newmssg
    #Parse message into an email object:
    messages =[parser.Parser().parsestr(mssg) for mssg in messages]
    #print("get new mail!")
    for mail in messages:
       print('init transaction mail...')
       tm=TransactionMail.TransactionMail(mail)
       print('checking mail...')
       if check_mail(tm):
            if not check_mail_readed(tm):
                if not hasattr(g,'trade_time_dict'):
                    __now=datetime.datetime.now()
                    g.trade_time_dict={
                    'start_time_of_am':__now.replace(hour=9,minute=30,second=0,microsecond=0),
                    'end_time_of_am':__now.replace(hour=11,minute=30,second=0,microsecond=0),
                    'start_time_of_pm':__now.replace(hour=13,minute=0,second=0,microsecond=0),
                    'end_time_of_pm':__now.replace(hour=15,minute=0,second=0,microsecond=0)
                    }
                
                trade_time_dict=g.trade_time_dict        
                #限制交易时段
                tmptime=datetime.datetime.now()
                if is_right_time(tmptime):
    
                    print('has new mail:%s[%s]'%(tm.get_subject(),tm.get_date()))
                    has_new_mail_count+=1
                    ti=tm.get_trading_info()
                    if ti:
                        trade_mark,action,dnum,price,code=ti
                        if action in ['买','卖']:
                            orderResult = None
                            if trade_mark[0] == '5':
                                trade_mark = trade_mark+'.XSHG'
                            if trade_mark[0] == '1':
                                trade_mark = trade_mark+'.XSHE'
                            if action == '买':
                                orderResult = order(trade_mark,dnum)
                            else:
                                orderResult = order(trade_mark,-dnum)
                                
                            if orderResult and orderResult.status == OrderStatus.held:
                                #保存邮件
                                if(save_mail(tm)):
                                    log_mail(tm)
                                    complete_count+=1
                                    print('交易成功！:%s'%orderResult)
                                else:
                                    print('保存邮件出现错误，订单取消。')
                                    cancel_order(orderResult)
                            else:
                                print('交易失败！')
                        else:
                           print('买卖动作信息有误！')
                    else:
                        print('交易信息有误！')
                else:
                    print('请在交易时间提交交易.')
    if has_new_mail_count>0:
        print('新交易邮件:(%d)/交易成功：(%d)' % (has_new_mail_count, complete_count))
    else:
        print('未读取到新的交易邮件。')
    print('运行结束：%s' % datetime.datetime.now().strftime('%Y年%m月%d日 %H:%M:%S'))
    #except RuntimeError:
        #print('运行时错误.')
    #except Exception:
        #print('程序错误。')

#login mail
def login_mail(host='pop.qq.com',user=mail_user_name,password=mail_password):
    try:
        global login_count
        global max_login_count
        print('正在连接QQ邮箱...')
        login_count+=1
        # host='pop.163.com'
        # username='18826500406@163.com'
        # print('正在连接163邮箱...')
        # password='linxiao163'
        pop_conn = poplib.POP3_SSL(host)
        print('邮箱登陆成功。当前时间：%s' % datetime.datetime.now().strftime('%Y年%m月%d日 %H:%M:%S'))
        #pop_conn.set_debuglevel(2)
        #print pop_conn.getwelcome()
        pop_conn.user(user);
        pop_conn.pass_(password);	
        #Get messages from server: 
        messages = [pop_conn.retr(i) for i in range(1, len(pop_conn.list()[1]) + 1)]
        pop_conn.quit()
        login_count=0
        return messages
    except Exception:
        print('登陆失败!(%d/%d)' % (login_count,max_login_count))
        if login_count<max_login_count:
            time.sleep(30)
            return login_mail()
        else:
            raise RuntimeError('%s暂时无法登陆' % host)
#保存邮件的hashvalue到hashvalues.txt文件
#tm:自定义邮件 邮件对象
def save_mail(tm):
    try:
        global hashSet
        hashvalue=tm.get_hashvalue()
        #print('hashvalue(save mail):'+hashvalue)
        hashSet.add(hashvalue)
        write_file('hashvalues.txt','\n'.join(hashSet)+'\n',append=False)
        return True
    except:
        return False
#从hashvalues.txt文件更新hashSet
def update_hashvalues():
   global hashSet
   #print('update hashvalues...')
   file=read_file('hashvalues.txt')
   hashSet.clear()
   hashSet=set(file.strip().split('\n'))

#检查邮件是否已读
#tm:自定义邮件 邮件对象
#return bool
def check_mail_readed(tm):
    global hashSet
    m_hashvalue=tm.get_hashvalue()
    return m_hashvalue in hashSet

#过滤邮件
#方式：验证发送者信息
#tm:自定义邮件 邮件对象
def check_mail(tm):
    global senderList
    global models
    modelName=tm.get_model()   
    sender=tm.get_sender()
    sender_date_text=tm.get_date()
    #print('model:%s\tsender:%s\ttradingInfo:%s' % (modelName,sender,tm.has_tradingInfo()))
    if (modelName is None) or (sender is None) or (sender_date_text is None):
        return False
    
    sender_date=get_sender_date_no_tzinfo(sender_date_text)
    now=datetime.datetime.now()
    #把接受邮件的发送时间限制在昨天
    if (now - sender_date).days >1 :
        return False
    # print('sender:%s'%sender)
    if (sender in senderList) and (modelName in models) and (tm.has_tradingInfo()):
        return True
    return False

def get_sender_date_no_tzinfo(sender_date_text):
    sender_date = None
    try:
        ref_date='Wed, 24 Jan 2018 10:04:19'
        if len(sender_date_text) > len(ref_date):
            sender_date_text=sender_date_text[0:len(ref_date)].strip()
        print('send date check:%s' % sender_date_text)
        sender_date=datetime.datetime.strptime(sender_date_text,'%a, %d %b %Y %H:%M:%S')
    except Exception as e:
        print('时间转换出现异常:')
        print(e)
    return sender_date
    
#log 
def log_mail(tm):
    now_text=datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    sender_date=get_sender_date_no_tzinfo(tm.get_date())
    log_info='{}\t接收时间：{}\t发送时间：{}\t{}\t{}\n'.format(tm.get_hashvalue(), now_text, sender_date, tm.get_model(), tm.get_sender())
    # log_info=tm.get_hashvalue()+'\t'+tm.get_date()+'\t'+tm.get_model()+'\t'+tm.get_sender()+'\n'
    write_file('log_mail.txt', log_info, append=True)

