# -*- coding:utf-8 -*-
__author__ = 'Rainbower'

import sys
import os
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)

import time
import logging
from multiprocessing import Pool
from lib.core import HandleArgvs
from config.settings import Params
from plugins.timetools import hours
from plugins.NameSerialize import NameConvert
from api import ExchangeApi as exchange, UserApi, Adapi
from lib.MailUserDistribution import MailUserDistribution

def ResumePost(**kwargs):
    '''AD账号默认不删除、复职用户需启用AD账号；职级、职位、部门'''
    obj = Adapi.Ad_Opertions()
    exec_obj = HandleArgvs()
    mail_obj = MailUserDistribution()
    user = kwargs['Badge']
    name = kwargs['xname']
    alias_name = NameConvert(name)
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
    user_obj = obj.Get_ObjectDN(type='user', sAMAccountName=user)
    if user_obj:
        old_user_dn = user_obj[0].entry_dn
        # self.obj.ChangePassword(user_obj[0].entry_dn, ResumePost_line['Badge'])
        if int(user_obj[0].userAccountControl[0]) != 512:
            new_user_dn, ou_dn = exec_obj.OrganizationalStructure(user, kwargs)
            obj.Enable_User(old_user_dn)
            # 把账号加入到通讯组
            obj.AddUserInGroups_PS(user, 'all.list')
            obj.AddUserInGroups_PS(user, org_id)
            obj.AddUserInGroups_PS(user, kwargs['depid_new_id'])
            obj.ChangePassword(user_obj[0].entry_dn, kwargs['Badge'])
            mail_user_dic = exec_obj.inject_dbdata(kwargs, result, ou_dn, org_id, upn, mail_domain,enable_mail_org_id)
            mail_obj.insert_update_maildb(**mail_user_dic)
            exchange.EnableMailboxhigh(user, alias=alias_name, database=result[0])
            obj.AclExchangeReadField(kwargs['Badge'])

            # 检查EHR提供的用户组织架构和AD域中的是否相同
            if old_user_dn != new_user_dn:
                print('用户组织机构发生变化，正在移动')
                res = obj.Move_Object(type='user', new_object_dn=ou_dn,
                                           sAMAccountName=kwargs['Badge'], cn=kwargs['Badge'])
                if res is True:
                    logging.info("移动用户从 %s 到 %s 成功" % (old_user_dn, ou_dn))
                else:
                    logging.error("移动用户从 %s 到 %s 失败" % (old_user_dn, ou_dn))

            # 检查用户的title和职级是否休要更新
            # title,telephoneNumber ,extensionAttribute2
            try:
                if user_obj[0].title[0] != kwargs['JOB_NEW']:
                    obj.Update_User_Info(new_user_dn, title=kwargs['JOB_NEW'])
            except:
                obj.Update_User_Info(new_user_dn, title=kwargs['JOB_NEW'])

            try:
                if user_obj[0].telephoneNumber[0] != str(kwargs['tel']):
                    obj.Update_User_Info(new_user_dn, telephoneNumber=kwargs['tel'])
            except:
                obj.Update_User_Info(new_user_dn, telephoneNumber=kwargs['tel'])
            try:
                if user_obj[0].extensionAttribute2[0] != str(kwargs['JOBGRADENEW']):
                    obj.Update_User_Info(new_user_dn,
                                              extensionAttribute2=kwargs['JOBGRADENEW'])
            except:
                obj.Update_User_Info(new_user_dn,
                                          extensionAttribute2=kwargs['JOBGRADENEW'])

            try:
                if user_obj[0].userPrincipalName[0].split('@')[1] != mail_domain:
                    obj.Update_User_Info(new_user_dn, extensionAttribute1=mail_domain)
                    obj.Update_User_Info(new_user_dn, mail='%s@%s' % (
                        user_obj[0].userPrincipalName[0].split('@')[0], mail_domain))
                    obj.Update_User_Info(new_user_dn, userPrincipalName='%s@%s' % (
                        user_obj[0].userPrincipalName[0].split('@')[0], mail_domain))
            except:
                obj.Update_User_Info(new_user_dn, extensionAttribute1=mail_domain)
                obj.Update_User_Info(new_user_dn, mail='%s@%s' % (
                    user_obj[0].userPrincipalName[0].split('@')[0], mail_domain))
                obj.Update_User_Info(new_user_dn, userPrincipalName='%s@%s' % (
                    user_obj[0].userPrincipalName[0].split('@')[0], mail_domain))

    else:
        # 超过三个月复制、账号可能已被人工删除、需重新创建
        new_user_dn, ou_dn = exec_obj.OrganizationalStructure(user, kwargs)
        obj.Add_User(ou_dn, name, user, kwargs['tel'],
                          kwargs['JOB_NEW'], kwargs['JOBGRADENEW'], domain=mail_domain,
                          alias=alias_name)
        exchange.EnableMailboxhigh(user, alias=alias_name, database=result[0])
        obj.AclExchangeReadField(kwargs['Badge'])
        time.sleep(1)
        # 把账号加入到通讯组
        obj.AddUserInGroups_PS(user, 'all.list')
        obj.AddUserInGroups_PS(user, org_id)
        obj.AddUserInGroups_PS(user, kwargs['depid_new_id'])
        time.sleep(1)
        mail_user_dic = exec_obj.inject_dbdata(kwargs, result, ou_dn, org_id, upn, mail_domain,enable_mail_org_id)
        mail_obj.insert_update_maildb(**mail_user_dic)

    # # 启用Exchange邮箱
    # if org_id in enable_mail_org_id:
    #     res = exchange.GetMailbox(user)
    #     if res.get('code') != 0:
    #         exchange.EnableMailboxhigh(user, alias=alias_name, database=result[0])
    #
    #         # 设置Exchange读取属性
    #         obj.AclExchangeReadField(kwargs['Badge'])


if  __name__ == '__main__':
    ResumePost_list = []
    start=time.time()
    user_obj = UserApi.EHR(6)
    user_data = user_obj.get_data()
    for user_line in user_data['Result']['empadd']['Row']:
        if hours(user_line['xClosedTime']) < 2:
            if "复职" in user_line['xtype']:
                ResumePost_list.append(user_line)
    print(len(ResumePost_list))
    ProcessPool = Pool(10)
    for i in ResumePost_list:
        ProcessPool.apply_async(ResumePost, kwds=i)
    ProcessPool.close()
    ProcessPool.join()
    end = time.time()
    print("总共用时{}秒".format((end - start)))