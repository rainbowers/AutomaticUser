# -*- coding:utf-8 -*-
__author__ = 'Rainbower'

import json
import requests
from config import settings


class EHR():
    def __init__(self,funcid):
        self.funcid = funcid
        self.headers = {"Content-Type": "application/json; charset=UTF-8"}


    def get_session(self):
        url = "http://"+settings.Params['EHRConfig']['host'] + settings.Params['EHRConfig']['url']['get_session']
        parameter = {"acc": settings.Params['EHRConfig']['user'], "pwd": settings.Params['EHRConfig']['password']}
        response = requests.post(url, data=json.dumps(parameter), headers=self.headers).json()
        token = response['Result']
        return token


    def get_data(self):
        url = "http://"+settings.Params['EHRConfig']['host'] + settings.Params['EHRConfig']['url']['get_data']
        parameter = {"funcId": self.funcid, "paras":{"WorkCity":1},"dataFormat":"json","dataPart":"DS","accessToken": self.get_session()}
        response = requests.post(url,data=json.dumps(parameter),headers=self.headers).json()
        # data = response['Result']
        data = response
        return data





if __name__ == "__main__":
    # import sys
    from plugins.timetools import hours
    obj = EHR(5)
    data=obj.get_data()

    #增量用户接口测试
    # for i in data['Result']['empadd']['Row']:
    #     if hours(i['xClosedTime']) < 12:
    #         if "离职" in i['xtype']:
    #     # if i['Badge']=='0718446':
    #             print(i)

    #     if i['Badge'] == '9500926':
    #         print(i)
    #     if i['xname'] == '贺驰':
    #         print(i)

    # 全量用户接口测试
    for i in data['Result']['emp']['Row']:
        if i['Badge'] == '0223220':
            print(i)

    # from plugins.timetools import hours
    # 增加组织架构
    # for i in data['Result']['organadd']['Row']:
    #     if '失效' in i['oprtype'] and hours(i['ClosedTime'])<7200:
    #         print(i)
        # day= hours(i['ClosedTime'])
        # if day<48:
        #     print(i)