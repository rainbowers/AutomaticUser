# -*- coding:utf-8 -*-
__author__ = 'Rainbower'
import os

BaseDir='\\'.join(os.path.abspath(os.path.dirname(__file__)).split('\\')[:-1])

Params = {
    'Adconfig':{
        "host" : '10.0.102.5',
        "user" : 'andy',
        "password" : '123456',
        "domain" : 'onesmart.local'
    },
    'Aadconfig':{
        "host" : '10.0.102.7',
        "user" : 'andy',
        "password" : '123456',
        "domain" : 'onesmart.local'
    },
    'ExchangeConfig':{
        "host" : '10.0.102.9',
        "user" : 'andy@onesmart.local',
        "password" : '123456',
    },
    'ExchangeOnline': {
        'user': 'exchangeonline-api',
        'password': '123456',
        'url': 'http://exchangeonline-api.onesmart.org/api/users/'
    },
    'ACGConfig':{
        "host" : '10.0.102.7',
        "user" : 'onesmart.local\andy',
        "password" : '123456',
    },
    'EHRConfig':{
        "host": '10.0.0.251',
        "user": 'inter_ehr',
        "password": '8ZAn9n6AkNyGz21Or/r0rg==',
        'url':{
            'get_session':'/kyapi/kayangwebapi/data/startsession',
            'get_data':'/kyapi/KayangWebApi/Data/GetData',
        }
    },
    'MailDomain':{
        'enable_mail':{
            '368': 'onesmart.org',      #A群管理
            '502': 'onesmart.org',      #加盟事业部
            '621': 'onesmart.org',      #家学宝
            '102': 'onesmart.org',      #精锐.个性化
            '334': 'onesmart.org',      #精锐.优毕慧
            '617': 'onesmart.org',      #精锐在线
            '627': 'onesmart.org',      #精锐在线|精锐.个性化
            '679': 'onesmart.org',      #精锐在线|精锐.班组课
            '666': 'onesmart.org',      #精锐在线|精锐.班课
            '268': 'onesmart.org',      #精锐在线国际教育
            '652': 'onesmart.org',      #精锐在线少儿
            '148': 'onesmart.org',      #在线少儿
            '101': 'onesmart.org',      #集团
            '698': 'onesmart.org',      #壮源高中
            '378': 'onesmart.org',      #B群管理
            '103': 'onesmart.org',      #至慧少儿数学
            '401': 'ftkenglish.com',    #小小地球
        },
        'not_enable_mail':{
            '339': 'onesmart.local',    #沪东外国语学校
            '599': 'onesmart.local',    #蒂比事业部


        }
    },
    'xtype': {
        1: '入职',
        2: '离职',
        3: '人事变更',
        4: '复职',
        5: '属性变更',
    },

    'MysqlConfig':{
        'port':3406,
        'user':'ldap_user',
        'password':'123456',
        'host':'192.168.100.139',
        'database':'jr_ldap_exchange',
    },

    'ScriptPath': '%s\\scripts'%BaseDir,
    'ExcelPath': '%s\\excel'%BaseDir,
    # 'ExcelFileList':['maillist1.xls','maillist2.xls','maillist3.xls'],
    'ExcelFileList':['地球校区邮件账号.xlsx'],
    'HeadOffice2161':'aq03.list@onesmart.org.xls',

}