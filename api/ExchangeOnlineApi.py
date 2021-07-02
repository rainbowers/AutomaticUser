# -*- coding:utf-8 -*-
__author__ = 'Rainbower'

import json
import base64
import logging
import requests
from config.settings import Params


# 日志设置
LOG_FORMAT = "%(asctime)s  %(levelname)s  %(filename)s  %(lineno)d  %(message)s"
logging.basicConfig(filename='../logs/logs.log', level=logging.INFO, format=LOG_FORMAT)


class ExchangeOnline():
    def __init__(self):
        self.url = Params['ExchangeOnline']['url']
        self.user_pwd = "%s:%s" %(Params['ExchangeOnline']['user'],Params['ExchangeOnline']['password'])
        self.auth = str(base64.b64encode(self.user_pwd.encode("utf-8")),"utf-8")
        self.headers = {
            'Authorization': 'Basic %s' % self.auth
        }
        self.payload = {}
    def AddAssignLicense(self,upn):
        '''
            添加邮箱授权
        :param upn: 如：1601016@onesmart.org
        :return:
        '''

        addassignlicense_url = self.url+upn+"/assignLicense"

        try:
            response = requests.request("POST",addassignlicense_url,headers=self.headers,data=self.payload)
            response = json.loads(response.text.encode('utf8'))
            if response['statusCode'] == 200:
                logging.info("用户邮箱:"+upn + "授权成功")
            else:
                logging.error("用户邮箱:" + upn + "授权失败")
        except Exception as e:
            logging.error(e)

    def DelAssignLicense(self,upn):
        '''
            删除邮箱授权
        :param upn:
        :return:
        '''
        delassignlicense_url = self.url+upn+"/disableLicense"
        print(delassignlicense_url)
        try:
            response = requests.request("POST", delassignlicense_url, headers=self.headers,data=self.payload)
            response = json.loads(response.text.encode('utf8'))
            print(type(response),response)
            print(response['statusCode'])
            if response['statusCode'] == 200:
                logging.info("用户邮箱:"+upn + "取消授权成功")
            else:
                logging.error("用户邮箱:" + upn + "取消授权失败")
        except Exception as e:
            logging.error(e)

if __name__ ==  '__main__':
    exobj=ExchangeOnline()
    exobj.AddAssignLicense('1601016@onesmart.org')

