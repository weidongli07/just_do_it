# -*- coding: utf-8 -*-
"""
Created on Tue Aug 16 13:52:37 2022

@author: PC
"""
# sms.py
#!/usr/local/bin/python
#-*- coding:utf-8 -*-
# Author: jacky
# Time: 14-2-22 下午11:48
# Desc: 短信http接口的python代码调用示例
import http.client
import urllib
import json
import requests
#参数的配置 请登录zz.253.com获取以下API信息 ↓↓↓↓↓↓↓
#创蓝接口域名
host = "smssh1.253.com"
#创蓝API账号
account  = "N699184_N9953671"
#创蓝API密码
password = "5vl2dIa8qqce79"
#端口号
port = 80
#版本号
version = "v1.1"
#余额查询的URL
balance_get_uri = "/msg/balance/json"
#普通短信发送的URL
sms_send_uri = "/msg/send/json"
params = {'account': account, 'password' : password}
headers = {"Content-type": "application/json"}
params=json.dumps(params)
r = requests.post('http://smssh1.253.com/msg/balance/json', json=params,headers=headers)
r.json()


def get_user_balance():
    """
    取账户余额
    """
    
    params = {'account': account, 'password' : password}
    
    params=json.dumps(params)
    
    
    
    headers = {"Content-type": "application/json"}
    
    conn = http.client.HTTPConnection(host, port=port)
    
    conn.request('POST', balance_get_uri, params, headers)
    
    response = conn.getresponse()
    
    response_str = response.read()
    
    conn.close()
    
    return response_str
def send_sms(text, phone):
    """
    能用接口发短信
    """
    
    
    
    params = {'account': account, 'password' : password, 'msg': urllib.request.quote(text), 'phone':phone, 'report' : 'false'}
    
    params=json.dumps(params)
    
    
    
    headers = {"Content-type": "application/json"}
    
    conn = http.client.HTTPConnection(host, port=port, timeout=30)
    
    conn.request("POST", sms_send_uri, params, headers)
    
    response = conn.getresponse()
    
    response_str = response.read()
    
    conn.close()

    return response_str

if __name__ == '__main__':
    phone = "13750508107"
    #设置您要发送的内容：其中“【】”中括号为运营商签名符号，多签名内容前置添加提交
    text = '''【回收宝】尊敬的用户，目前尚未收到您的回收机器，手机贬值速度快，为防您的设备跌价，越早寄出越值钱！24小时内寄出且最终完成回收可额外获赠20元话费！话费3个工作日充值给您，邮费到付～邮寄地址：(东莞市东坑镇科技路280号恒钜科技园C栋101室
    0755-33941079
    回收宝收）如有疑问可联系400-080-9966'''
    
    
    
    #查账户余额
    
    print(get_user_balance())
    
    
    
    # 调用智能匹配模版接口发短信
    
    print(send_sms(text, phone))