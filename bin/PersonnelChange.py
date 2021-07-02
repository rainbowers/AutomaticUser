# -*- coding:utf-8 -*-
__author__ = 'Rainbower'

import sys
import os
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)

import time
import logging
from api import UserApi, Adapi
from multiprocessing import Pool
from lib.core import HandleArgvs
from plugins.tools import is_number
from plugins.timetools import hours

def PersonnelChange(**kwargs):
    obj = Adapi.Ad_Opertions()
    exec_obj = HandleArgvs()
    user = kwargs['Badge']
    org_id = kwargs['SHIYEBU_new_ID']
    new_user_dn, ou_dn = exec_obj.OrganizationalStructure(user, kwargs)
    user_obj = obj.Get_ObjectDN(type='user', sAMAccountName=user)
    if user_obj:
        old_user_dn = user_obj[0].entry_dn
        if old_user_dn != new_user_dn:
            # 更新群组,根据架构的调整更新用户隶属于的组（添加组、删除组）
            user_groups_list = obj.GetUserGroups(kwargs['Badge'])
            first_group_dn = 'CN=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
                org_id, kwargs['SHIYEBU_new'])
            five_group_dn = 'CN=%s,OU=%s,OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
                kwargs['depid_new_id'], kwargs['DEP_NEW'],
                kwargs['zuzhiquyu_NEW'],
                kwargs['zuzhishengfen_new'],
                kwargs['zuzhifenqu_new'], kwargs['SHIYEBU_new'])

            for i in user_groups_list:
                groupname = i.split(',')[0].split('=')[1]
                if is_number(groupname):
                    if len(groupname) == 3:  # 一级部门
                        if groupname != org_id:
                            obj.RemoveUserOutGroups([user_obj[0].entry_dn], [i])  # 移除旧组
                            obj.AddUserInGroups([user_obj[0].entry_dn],
                                                [first_group_dn])  # 添加到新组
                    else:  # 其他非一级部门
                        if groupname != kwargs['depid_new_id']:
                            obj.RemoveUserOutGroups([user_obj[0].entry_dn], [i])  # 移除旧组
                            obj.AddUserInGroups([user_obj[0].entry_dn],
                                                [five_group_dn])  # 添加到新组
            obj.Move_Object(type='user', new_object_dn=ou_dn,
                            sAMAccountName=kwargs['Badge'],
                            cn=kwargs['Badge'])

        # 职位变更
        try:
            if user_obj[0].title[0] != kwargs['JOB_NEW']:
                obj.Update_User_Info(new_user_dn, title=kwargs['JOB_NEW'])
        except:
            obj.Update_User_Info(new_user_dn, title=kwargs['JOB_NEW'])

        # 职级变更
        try:
            if user_obj[0].extensionAttribute2[0] != kwargs['JOBGRADENEW']:
                obj.Update_User_Info(new_user_dn,
                                     extensionAttribute2=kwargs['JOBGRADENEW'])
        except:
            obj.Update_User_Info(new_user_dn,
                                 extensionAttribute2=kwargs['JOBGRADENEW'])


if __name__ == '__main__':
    PersonnelChange_list = []
    start = time.time()
    user_obj = UserApi.EHR(6)
    user_data = user_obj.get_data()
    for user_line in user_data['Result']['empadd']['Row']:
        if hours(user_line['xClosedTime']) < 2:
            if "人事变更" in user_line['xtype']:
                PersonnelChange_list.append(user_line)
    print(len(PersonnelChange_list))
    ProcessPool = Pool(10)
    for i in PersonnelChange_list:
        ProcessPool.apply_async(PersonnelChange, kwds=i)
    ProcessPool.close()
    ProcessPool.join()
    end = time.time()
    print("总共用时{}秒".format((end - start)))
