# -*- coding: utf-8 -*-
"""
Created on Wed Apr 26 16:23:09 2023

@author: PC
"""
import requests
import json
import datetime
import time
class Sobot:
    
    def __init__(self,apiKey,apiSecret):
        self.apiKey = apiKey
        self.apiSecret = apiSecret
        
    def getToken(self):
        self.url = 'https://gw.soboten.com/api/6.0.0/tokens?apiKey={0}&apiSecret={1}'.format(self.apiKey,self.apiSecret)
        self.rep = requests.get(self.url)
        self.json_result = self.rep.json()
        self.accessToken = self.json_result.get('content').get('accessToken') #返回的token
        self.create_time = self.json_result.get('content').get('createTime') #创建token时间，格式为毫秒时间戳
        self.expiresIn = self.json_result.get('content').get('expiresIn') #token有效时间
        self.headers = { "Content-Type": "application/json;charset=utf-8"
            , 'Authorization': f'Bearer '+ self.accessToken }
        return self.accessToken,self.create_time,self.expiresIn,self.headers
    
    def addnum(self,companyId,taskId,data):
        self.url_add_number = 'https://gw.soboten.com/api/6.0.0/companies/{0}/tasks/{1}/task-contacts'.format(companyId,taskId)
        self.rep_add_number = requests.post(self.url_add_number,headers = self.headers,data = data)
        self.json_add_number = self.rep_add_number.json()
        print("数据导入结果：",self.json_add_number)

    def taskresult(self,companyId,taskId,data):
       self.url_taskresult = 'https://gw.soboten.com/api/6.0.0/companies/{0}/tasks/{1}/dialing-results'.format(companyId,taskId)
       self.rep_taskresult =  requests.post(self.url_taskresult,headers = self.headers,data = data)
       self.json_taskresult = self.rep_taskresult.json()
       return self.json_taskresult





