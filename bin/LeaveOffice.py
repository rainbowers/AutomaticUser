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
from plugins.timetools import hours
from api import ExchangeApi as exchange, UserApi, Adapi



def LeaveOffice(**kwargs):
    obj = Adapi.Ad_Opertions()
    user = kwargs['Badge']
    name = kwargs['xname']
    user_obj = obj.Get_ObjectDN(type='user', sAMAccountName=user)
    if user_obj:
        old_user_dn = user_obj[0].entry_dn
    if user_obj:
        # 禁用邮箱账号
        if int(user_obj[0].userAccountControl[0]) != 514:  # 避免重复禁用账号，导致用户默认组（Domain Users）不存在
            # print("正在禁用 %s 、工号是：%s，请稍等……" % (name, user))
            '''临时关闭删除邮箱功能，后期单独设计此功'''
            # 禁用邮箱
            res = exchange.GetMailbox(user)
            if res.get('code') == 0:
                res = exchange.DisableMailbox(user)
                if res.get('code') == 0:
                    logging.info(
                        "禁用邮箱" + user_obj[0].userPrincipalName[0] + "," + old_user_dn + "成功")
                else:
                    logging.error(
                        "禁用邮箱" + user_obj[0].userPrincipalName[0] + "," + old_user_dn + "失败")

            # 禁用AD域账号
            obj.Disable_User(old_user_dn, user)

            # 移除相关组，只保留Domain Users组
            ou_list = obj.GetUserGroups(user)
            if ou_list:
                obj.RemoveUserOutGroups([user_obj[0].entry_dn], ou_list)


if __name__ == '__main__':
    LeaveOffice_list = []
    start = time.time()
    user_obj = UserApi.EHR(6)
    user_data = user_obj.get_data()
    for user_line in user_data['Result']['empadd']['Row']:
        if hours(user_line['xClosedTime']) < 2:
            if "离职" in user_line['xtype']:
                LeaveOffice_list.append(user_line)
    print(len(LeaveOffice_list))
    ProcessPool = Pool(10)
    for i in LeaveOffice_list:
        ProcessPool.apply_async(LeaveOffice, kwds=i)
    ProcessPool.close()
    ProcessPool.join()
    end = time.time()
    print("总共用时{}秒".format((end - start)))
