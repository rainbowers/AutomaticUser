# -*- coding:utf-8 -*-
__author__ = 'Rainbower'

import sys
import os
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)

import time
from multiprocessing import Pool
from config.settings import Params
from plugins.timetools import hours
from plugins.NameSerialize import NameConvert
from api import ExchangeApi as exchange, UserApi, Adapi
from lib.MailUserDistribution import MailUserDistribution
from lib.core import HandleArgvs

def Induction(**kwargs):
    obj = Adapi.Ad_Opertions()
    exec_obj=HandleArgvs()
    mail_obj = MailUserDistribution()
    user = kwargs['Badge']
    name = kwargs['xname']
    alias_name = NameConvert(name)
    rank = kwargs['JOBGRADENEW']
    result = mail_obj.select_maildb()
    org_id = str(kwargs['SHIYEBU_new_ID'])
    try:
        mail_domain = Params['MailDomain']['enable_mail'][org_id]
    except:
        mail_domain = Params['MailDomain']['not_enable_mail'][org_id]

    enable_mail_org_id = []
    not_enable_mail_org_id = []
    for k, v in Params['MailDomain']['enable_mail'].items():
        enable_mail_org_id.append(k)
    for k, v in Params['MailDomain']['not_enable_mail'].items():
        not_enable_mail_org_id.append(k)

    if org_id in enable_mail_org_id:
        upn = alias_name + '@' + Params['MailDomain']['enable_mail'][org_id]
    new_user_dn, ou_dn = exec_obj.OrganizationalStructure(user, kwargs)
    user_obj = obj.Get_ObjectDN(type='user', sAMAccountName=user)
    if org_id in enable_mail_org_id:
        if not user_obj:
            # print(kwargs)
            if "外籍" in kwargs['JOB_NEW']:
                alias_name = kwargs['Badge']
            obj.Add_User(ou_dn, kwargs['xname'], kwargs['Badge'], kwargs['tel'],
                              kwargs['JOB_NEW'], kwargs['JOBGRADENEW'], domain=mail_domain,
                              alias=alias_name)
            time.sleep(1)
            # 把账号加入到通讯组
            obj.AddUserInGroups_PS(user, 'all.list')
            obj.AddUserInGroups_PS(user, org_id)
            obj.AddUserInGroups_PS(user, kwargs['depid_new_id'])
            time.sleep(1)

            mail_user_dic = exec_obj.inject_dbdata(kwargs, result, ou_dn, org_id, upn, mail_domain,enable_mail_org_id)
            # print('入职用户信息：',mail_user_dic)

            # 启用邮箱
            exchange.EnableMailboxhigh(user, alias=alias_name, database=result[0])

            # 设置Exchange读取属性
            obj.AclExchangeReadField(kwargs['Badge'])

            # 入库
            mail_obj.insert_update_maildb(**mail_user_dic)


if  __name__ == '__main__':
    Induction_list = []
    start=time.time()
    user_obj = UserApi.EHR(6)
    user_data = user_obj.get_data()
    for user_line in user_data['Result']['empadd']['Row']:
        if hours(user_line['xClosedTime']) < 2:
            if "入职" in user_line['xtype']:
                Induction_list.append(user_line)
    print(len(Induction_list))
    ProcessPool = Pool(10)
    for i in Induction_list:
        ProcessPool.apply_async(Induction, kwds=i)
    ProcessPool.close()
    ProcessPool.join()
    end = time.time()
    print("总共用时{}秒".format((end - start)))