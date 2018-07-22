#-*- coding: utf-8 -*-
# Author: 林潇
# Contact: 940950943@qq.com

'''   自定义邮件类  '''

import poplib
from email.header import Header
from email.utils import parseaddr
from email.header import decode_header
import email
import string
import traceback
import datetime
import time
import hashlib
from hashlib import md5
import re

#自定义邮件类   
class TransactionMail:
    def __init__(self,message):
        self.__mail=message
        self.__has_tradingInfo=False
        self.__senderName=None 
        self.__sender=None        
        self.__date=None
        self.__subject=None
        self.__model=None
        self.__module=None
        self.__contract=None
        self.__tradingInfo=None
        self.__hashvalue=None
        self.__parse_mail()
    
    #分析邮件信息
    def __parse_mail(self):
        mail=self.__mail
        #print('parse email...')
        #print('keys:%s' % mail.keys())
        #print(mail)
        #发件人信息
        sender=mail.get('From','')
        if sender:
            senderName,self.__sender=parseaddr(sender)
            self.__senderName=self.__decode_str(senderName)
        else:
            # print(self.__mail)
            pass
        #print('sender:%s' % self.get_sender())
        #日期        
        date=mail.get('Date','')
        if date:
            self.__date=date
        else:
            print('未读取到邮件发送时间！')
        #主题
        subject=mail.get('Subject','')
        if subject:
            self.__subject=self.__decode_str(subject)
            
        #print('send date:%s\nsubject:%s' %(self.__date,self.__subject))
        
        #分析交易信息
        for part in mail.walk():
                    content_type = part.get_content_type()
                    if content_type=='text/plain':#or content_type=='text/html':
                        #print(mail)
                        tranEncoding=part.get('Content-Transfer-Encoding','')
                        charset = self.__guess_charset(part)
                        if charset=='gb2312':
                            content = part.get_payload(decode=False)
                        else:
                            content = part.get_payload(decode=True)
                            content = content.decode(charset).encode('utf-8')
                            #content=self.decode_base64(content,charset)
                        #print('content text:%s' % content)
                        #print(part)
                        if content:
                                 
                            #content=content.encode('utf-8').decode('utf-8')
                            #print(content)
                            mz=re.search(r'模组名: (.+)',content)
                            mx=re.search(r'模型名: (.+)',content,re.MULTILINE)
                            hy=re.search(r'合约名: (.+)[ ]*周期',content,re.MULTILINE)
                            
                            self.__module=mz.groups(1)[0].strip() if mz else None
                            #print('模组名:%s' % self.__module)
                            
                            self.__model=mx.groups(1)[0].strip() if mx else None
                            #print('模型名:%s' % self.__model)
                            
                            self.__contract=hy.groups(1)[0].strip() if hy else None
                            #print('合约名:%s' % self.__contract)
                            
                            cj=re.search(r'成交\((.+)\)',content)
                            self.__has_tradingInfo=True if cj else False
                            self.__tradingInfo=cj.groups(1)[0].strip() if cj else None
                                
                            #if self.__has_tradingInfo:
                                #print('成交信息:%s' % self.__tradingInfo)
                        # print('分析完成.')
                        break#仅分析纯文本
                    else:
                        # print('Attachment: %s' % (content_type))
                        pass
    
    #获取发送时间
    def get_date(self):
        return self.__date
        
    #获取邮件标题
    def get_subject(self):
        return self.__subject
       
    #获取发送人地址
    def get_sender(self):
        return self.__sender
        
    #获取模型名
    def get_model(self):
        return self.__model
    
    #是否有合格的交易信息
    #return bool
    def has_tradingInfo(self):
        return self.__has_tradingInfo
        
    #交易信息
    #return None or [action,dnum,price,code]  
    def get_trading_info(self):
        if not self.has_tradingInfo():
            return None
        else:                         
            tradingInfo=self.__tradingInfo.split(',')
            trade_mark=tradingInfo[0]
            price=float(tradingInfo[1]) 
            action=tradingInfo[2]
            dnum=int(tradingInfo[3])
            code=tradingInfo[4].split(':')[1].strip()
            print('交易信息：\n交易标:%s\n价格：%s\n数量：%s\n动作：%s\n委托编号：%s'%(trade_mark , price , dnum , action , code))
            if action=='买':
                 print('买入%d手'%(dnum/100))
            elif action=='卖':
                 print('卖出%d手'%(dnum/100))
            else:
                 print('数据出现异常：买卖动作')
            return [trade_mark,action,dnum,price,code]
           
    #获取邮件hash编码 用于标记邮件
    #return (sender+date+model)的MD5值
    #sender 发送人信息
    #date 邮件发送时间
    #model 模型名
    def get_hashvalue(self):
        if self.__hashvalue:
            return self.__hashvalue 
        
        model=self.get_model()
        date=self.get_date()
        sender=self.get_sender()
        if date and sender:
            if model:
                hashtext=sender+date+model
            else:
                hashtext=sender+date
            self.__hashvalue = hashlib.md5(hashtext.encode('utf-8')).hexdigest()
            return self.__hashvalue
        else:
            return None
        
    def __decode_str(self,s):
        value, charset = decode_header(s)[0]
        #print('charset:%s'%(charset))
        if charset:
            value = value.decode(charset)
            pass
        return value

    def __guess_charset(self,msg):
        charset = msg.get_charset()
        
        if charset is None:
            print(charset is None)
            content_type = msg.get('Content-Type', '').lower()
            pos = content_type.find('charset=')
            if pos >= 0:
                charset = content_type[pos + 8:].strip()
                charset = charset.replace('"','')
                
        #print('guess_charset:%s'%charset)
        return charset

 