# -*- coding:utf-8 -*-
__author__ = 'Rainbower'

import sys
import time
import xlrd
import logging
import datetime
from config.settings import Params
from api import ExchangeApi as exchange, UserApi, Adapi
from api.SynchronizationAAD import SyncAAD
from api.ExchangeOnlineApi import ExchangeOnline
from lib.ChineseToPinyin import ChineseToPinyin
from lib.MailUserDistribution import MailUserDistribution
from plugins.OpertionsToExcel import ExcelData
from plugins.NameSerialize import NameConvert
from plugins.tools import is_number
from plugins.timetools import hours
from multiprocessing import Pool


# 日志设置
# LOG_FORMAT = "%(asctime)s  %(levelname)s  %(filename)s  %(lineno)d  %(message)s"
# logging.basicConfig(filename='../logs/debug.log', level=logging.INFO, format=LOG_FORMAT)

class HandleArgvs():
    def __init__(self):
        self.obj= Adapi.Ad_Opertions()
        # self.exchange_ojb = ExchangeApi.Exchange()
        self.mail_obj = MailUserDistribution()
        self.sync_aad = SyncAAD()
        self.online_obj = ExchangeOnline()


    def create_or_update_ou(self,org_line):
        org_dn_dic = {}
        org_name_dic = {}
        org_description_dic = {}

        org_description_dic[1] = org_line['shiyebuid']
        org_description_dic[2] = org_line['zuzhifenquid']
        org_description_dic[3] = org_line['zuzhishenfenid']
        org_description_dic[4] = org_line['zuzhiquyuid']
        org_description_dic[5] = org_line['Uniqueid']

        org_name_dic[1] = org_line['shiyebu']
        org_name_dic[2] = org_line['zuzhifenqu']
        org_name_dic[3] = org_line['zuzhishenfen']
        org_name_dic[4] = org_line['zuzhiquyu']
        try:
            org_name_dic[5] = org_line['title']
        except:
            org_name_dic[5] = org_line['Title']

        org_dn_dic[1] = 'OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (org_line['shiyebu'])
        org_dn_dic[2] = 'OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (org_line['zuzhifenqu'], org_line['shiyebu'])
        org_dn_dic[3] = 'OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
        org_line['zuzhishenfen'], org_line['zuzhifenqu'], org_line['shiyebu'])
        org_dn_dic[4] = 'OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
        org_line['zuzhiquyu'], org_line['zuzhishenfen'], org_line['zuzhifenqu'], org_line['shiyebu'])
        try:
            org_dn_dic[5] = 'OU=%s,OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
            org_line['title'], org_line['zuzhiquyu'], org_line['zuzhishenfen'], org_line['zuzhifenqu'], org_line['shiyebu'])
        except:
            org_dn_dic[5] = 'OU=%s,OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
                org_line['Title'], org_line['zuzhiquyu'], org_line['zuzhishenfen'], org_line['zuzhifenqu'],
                org_line['shiyebu'])

        number = 1
        while number < 6:
            base_dn = org_dn_dic[number]
            if "(" in org_name_dic[number]:
                org_isexist = self.obj.Get_ObjectDN(type='org', base_dn='DC=onesmart,DC=local',
                                                    ou="%s*" % org_name_dic[number].split("(")[0],
                                                    description=org_description_dic[number])
            else:
                org_isexist = self.obj.Get_ObjectDN(type='org', base_dn='DC=onesmart,DC=local',
                                                    ou=org_name_dic[number],
                                                    description=org_description_dic[number])  # 获取组织单位
            if not org_isexist:
                print(org_line)
                # 需EHR系统提供组织架构的最近数据变更记录的接口

                # 增加组织单位
                print("不存在", org_line)
                self.obj.Add_Org(org_dn=org_dn_dic[number], org_id=org_description_dic[number])

                # 创建通用通讯组，只在一级和五级部门下创建通讯组,已组织单位唯一ID命名
                if number == 1:
                    group_name = org_line['shiyebuid']
                    group_dn = 'CN=%s,%s' % (org_line['shiyebuid'], org_dn_dic[1])
                    group_is_exist = self.obj.Get_ObjectDN(type='group', base_dn=org_dn_dic[1],
                                                           CN=group_name,
                                                           description=org_description_dic[1])
                    if not group_is_exist:
                        self.obj.Add_User_Group(group_name, group_dn, -2147483640, org_description_dic[1])
                        time.sleep(0.5)
                        exchange.EnableDistributionGroup(group_name)
                elif number == 5:
                    # 五级部门下通讯组名称为：四级部门名称+五级部门名称，避免因重名而创建失败
                    group_name = org_line['Uniqueid']
                    group_dn = 'CN=%s,%s' % (group_name, org_dn_dic[5])
                    group_is_exist = self.obj.Get_ObjectDN(type='group', base_dn=org_dn_dic[5],
                                                           description=org_description_dic[5])
                    if not group_is_exist:
                        self.obj.Add_User_Group(group_name, group_dn, -2147483640, org_description_dic[5])
                        time.sleep(0.5)
                        exchange.EnableDistributionGroup(group_name)
            else:  # 组织单位已存在、可能需要修改或需要移动、或不做操作
                # 修改组织单位
                old_org_dn = org_isexist[0].entry_dn
                old_org_name = old_org_dn.split(',')[0].split('=')[1]
                new_org_name = org_dn_dic[number].split(',')[0].split('=')[1]
                if old_org_name != new_org_name:
                    # print(org_line)
                    print('正在重命名组织单位：%s 为%s' % (old_org_name, new_org_name))
                    self.obj.Rename_Org(old_org_dn, new_org_name)
                    self.obj.Update_Org_Info(old_org_dn, description=org_description_dic[number])
                else:
                    # print("old_org_dn:", old_org_dn, "new_org_dn:", base_dn)
                    if old_org_dn != org_dn_dic[number]:
                        # print(old_org_dn, org_dn_dic[number])
                        result = self.obj.Move_Object(type='org', new_object_dn=org_dn_dic[number - 1],
                                                      ou=org_dn_dic[number].split(',')[0].split('=')[1],
                                                      description=org_description_dic[number])
                        if result is True:
                            logging.info("移动组织单位%s到%s成功" % (old_org_dn, org_dn_dic[number - 1]))
                        else:
                            logging.error("移动组织单位%s到%s失败" % (old_org_dn, org_dn_dic[number - 1]))
            number += 1


    def operations_ou_full(self):
        '''
            全量数据更新
        :return:
        '''
        obj_data = UserApi.EHR(7)
        org_data = obj_data.get_data()
        org_data = org_data['Result']['org']['Row']

        for org_line in org_data:
            print(org_line)
            # if org_line['type'] == '部门':
            #     self.create_or_update_ou(org_line)

    def operations_ou_diff(self, data=None):
        '''
            差异数据更新
        :param data:
        :return:
        '''
        obj_data = UserApi.EHR(10)
        org_data = obj_data.get_data()
        org_data = org_data['Result']['organadd']['Row']

        for org_line in org_data:
            # day = hours(org_line['ClosedTime'])
            # current_year=str(datetime.datetime.now().year)
            if '失效' in org_line['oprtype'] and hours(org_line['ClosedTime'])<24 and "部门" in org_line['oprPROPERTY']:
                org_dn = 'OU=%s,OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
                    org_line['Title'], org_line['zuzhiquyu'], org_line['zuzhishenfen'], org_line['zuzhifenqu'],
                    org_line['shiyebu'])
                self.obj.Update_dn(org_dn,NewNmae="ou=%s-失效"%org_line['Title'])
                self.obj.Update_Org_Info(org_dn,description='失效')


    def inject_dbdata(self, userdata,maildbnumber,ou_dn,org_id,upn,mail_domain,enable_mail_org_id):
        # print(userdata,maildbnumber,ou_dn,org_id,upn,mail_domain)
        mail_user_dic = {
            'maildb': maildbnumber[0],
            'jobnumber': userdata['Badge'],
            'counter': maildbnumber[1] + 1,
            'jobgrade': userdata['jobgrade'],
            'createtime': time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
            'organization':ou_dn,
        }
        try:
            mail_user_dic['username'] = userdata['Name1']
        except:
            mail_user_dic['username'] = userdata['xname']

        try:
            mail_user_dic['title'] = userdata['JOBTITLE']
        except:
            mail_user_dic['title'] = userdata['JOB_NEW']

        try:
            mail_user_dic['jobgrade'] = userdata['jobgrade']
        except:
            mail_user_dic['jobgrade'] = userdata['JOBGRADENEW']

        try:
            try:
                mail_user_dic['mobile'] = userdata['mobile']
            except:
                mail_user_dic['mobile'] = userdata['tel']
        except:
            mail_user_dic['mobile'] = 'unknown'

        try:
            mail_user_dic['gender'] = userdata['GENDER']
        except:
            mail_user_dic['gender'] = userdata['gender']

        if org_id in enable_mail_org_id:
            mail_user_dic['is_active'] = 0
            if userdata['email']:
                #全量添加用户时、邮件地址使用现有的
                if "<" in userdata['email'] and ">" in userdata['email']:
                    email = userdata['email'].split('<')[1].split('>')[0]
                    mail_user_dic['email'] = email.split('@')[0] + '@' + mail_domain
                else:
                    mail_user_dic['email'] = userdata['email'].split('@')[0] + '@' + mail_domain
            else:
                #增量添加账号时、邮箱地址需要拼接、已存在邮箱则+1 如：zhangshan1@onesmart.org、zhangsan2@onesmart.org
                mail_user_dic['email'] = upn

        else:
            mail_user_dic['is_active'] = 1
            if userdata['email']:
                mail_user_dic['email'] = userdata['email'].split('@')[0] + '@' +  mail_domain
            else:
                mail_user_dic['email'] = upn

        return mail_user_dic

    def OrganizationalStructure(self,user ,user_data):
        '''
            根据用户数据，拼接完整的组织架构地址
        :param user:
        :param user_data:
        :return:
        '''

        org_dic = {
            'shiyebuid': user_data['SHIYEBU_new_ID'],
            'shiyebu': user_data['SHIYEBU_new'],
            'zuzhifenquid': user_data['zuzhifenqu_new_ID'],
            'zuzhifenqu': user_data['zuzhifenqu_new'],
            'zuzhishenfenid': user_data['zuzhishengfen_new_ID'],
            'zuzhishenfen': user_data['zuzhishengfen_new'],
            'zuzhiquyuid': user_data['zuzhiquyu_NEW_id'],
            'zuzhiquyu': user_data['zuzhiquyu_NEW'],
            'Uniqueid': user_data['depid_new_id'],
            'title': user_data['DEP_NEW'],
        }

        # ou_is_exist = self.obj.Get_ObjectDN(type='org',ou=org_dic['title'] ,description=user_data['depid_new_id'])
        base_dn="OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local" % (
                org_dic['zuzhiquyu'], org_dic['zuzhishenfen'], org_dic['zuzhifenqu'],
                org_dic['shiyebu'])
        ou_is_exist = self.obj.Get_ObjectDN(type='org',base_dn=base_dn,ou="%s*"%user_data['DEP_NEW'] )
        if ou_is_exist:
            ou_dn = ou_is_exist[0].entry_dn
            new_user_dn = 'CN=%s,%s' % (user, ou_dn)
        else:
            self.create_or_update_ou(org_dic)
            ou_dn = "OU=%s,OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local" % (
                org_dic['title'], org_dic['zuzhiquyu'], org_dic['zuzhishenfen'], org_dic['zuzhifenqu'],
                org_dic['shiyebu'])
            new_user_dn = 'CN=%s,%s' % (user, ou_dn)
        return new_user_dn,ou_dn

    def operations_user(self,funcid,active=None):
        '''
            第一次时倒入所有用户信息;以后每次增量更新：变更部门、区域、title、新增用户等
        :return:
        '''
        user_obj = UserApi.EHR(funcid)
        user_data = user_obj.get_data()

        self.enable_mail_org_id = []
        self.not_enable_mail_org_id = []
        for k,v in Params['MailDomain']['enable_mail'].items():
            self.enable_mail_org_id.append(k)
        for k,v in Params['MailDomain']['not_enable_mail'].items():
            self.not_enable_mail_org_id.append(k)

        # RTC账号清单获取，用于测试
        file_path = r"%s\%s" % (Params['ExcelPath'], 'rtc.xlsx')
        excel_obj = ExcelData(file_path, 'user')
        user_list = []
        for i in excel_obj.readExcel():
            user_list.append(i['jobnumber'])

        if funcid == 5:
            # 用户全量导入
            for user_line in user_data['Result']['emp']['Row']:
                self.org_id = str(user_line['shiyebuid'])
                self.result = self.mail_obj.select_maildb()
                self.ou_dn = 'OU=%s,OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
                    user_line['DEPTITLE'], user_line['zuzhiquyu'], user_line['zuzhishenfen'],
                    user_line['zuzhifenqu'],user_line['shiyebu'])

                #创建用户
                try:
                    org_id = str(user_line['shiyebuid'])
                except:
                    org_id = str(user_line['SHIYEBU_new_ID'])

                is_exist = self.obj.Get_ObjectDN(type='user', sAMAccountName=user_line['Badge'])

                if not is_exist:
                        # if org_id in self.enable_mail_org_id and "小小地球" in user_line['shiyebu']: #需要启用邮箱的用户
                        if org_id in self.enable_mail_org_id: #需要启用邮箱的用户
                            self.mail_domain = Params['MailDomain']['enable_mail'][org_id]
                            if user_line['email']:
                                if "<" in user_line['email'] and ">" in user_line['email']:
                                    email = user_line['email'].split('<')[1].split('>')[0]
                                else:
                                    email = user_line['email']
                                alias_name = email.split('@')[0]

                                self.upn = email.split('@')[0] + '@' + self.mail_domain
                                self.obj.Add_User(self.ou_dn, user_line['Name1'], user_line['Badge'],mobile=user_line['mobile'],title=user_line['JOBTITLE'],jobgrade=user_line['jobgrade'],domain=self.mail_domain,alias=alias_name)

                                # 把账号加入到通讯组
                                self.obj.AddUserInGroups_PS(user_line['Badge'], 'all.list')
                                self.obj.AddUserInGroups_PS(user_line['Badge'], self.org_id)
                                self.obj.AddUserInGroups_PS(user_line['Badge'], user_line['DepID'])
                            else:
                                #只创建AD账号
                                self.obj.Add_User(self.ou_dn, user_line['Name1'], user_line['Badge'],
                                                  mobile=user_line['mobile'], title=user_line['JOBTITLE'],
                                                  jobgrade=user_line['jobgrade'], domain=self.mail_domain)
                            mail_user_dic = self.inject_dbdata(user_line)

    def Induction(self, **kwargs):
        print(kwargs)
        user = kwargs['Badge']
        name = kwargs['xname']
        alias_name = NameConvert(name)
        rank = kwargs['JOBGRADENEW']
        result = self.mail_obj.select_maildb()
        org_id = str(kwargs['SHIYEBU_new_ID'])
        try:
            mail_domain = Params['MailDomain']['enable_mail'][org_id]
        except:
            mail_domain = Params['MailDomain']['not_enable_mail'][org_id]
        if org_id in self.enable_mail_org_id:
            upn = alias_name + '@' + Params['MailDomain']['enable_mail'][org_id]
        new_user_dn, ou_dn = self.OrganizationalStructure(user, kwargs)
        user_obj = self.obj.Get_ObjectDN(type='user', sAMAccountName=user)
        if org_id in self.enable_mail_org_id:
            if not user_obj:
                print(kwargs)
                if "外籍" in kwargs['JOB_NEW']:
                    alias_name = kwargs['Badge']
                self.obj.Add_User(ou_dn, kwargs['xname'], kwargs['Badge'], kwargs['tel'],
                                  kwargs['JOB_NEW'], kwargs['JOBGRADENEW'], domain=mail_domain,
                                  alias=alias_name)
                time.sleep(1)
                # 把账号加入到通讯组
                self.obj.AddUserInGroups_PS(user, 'all.list')
                self.obj.AddUserInGroups_PS(user, org_id)
                self.obj.AddUserInGroups_PS(user, kwargs['depid_new_id'])
                time.sleep(1)

                mail_user_dic = self.inject_dbdata(kwargs, result, ou_dn, org_id, upn, mail_domain)
                # print('入职用户信息：',mail_user_dic)

                # 启用邮箱
                exchange.EnableMailboxhigh(user, alias=alias_name, database=result[0])

                # 设置Exchange读取属性
                self.obj.AclExchangeReadField(kwargs['Badge'])

                # 入库
                self.mail_obj.insert_update_maildb(**mail_user_dic)

    def LeaveOffice(self, **kwargs):
        print(kwargs)
        user = kwargs['Badge']
        name = kwargs['xname']
        user_obj = self.obj.Get_ObjectDN(type='user', sAMAccountName=user)
        if user_obj:
            old_user_dn = user_obj[0].entry_dn
        if user_obj:
            # 禁用邮箱账号
            if int(user_obj[0].userAccountControl[0]) != 514:  # 避免重复禁用账号，导致用户默认组（Domain Users）不存在
                print("正在禁用 %s 、工号是：%s，请稍等……" % (name, user))
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
                self.obj.Disable_User(old_user_dn, user)

                # 移除相关组，只保留Domain Users组
                ou_list = self.obj.GetUserGroups(user)
                if ou_list:
                    self.obj.RemoveUserOutGroups([user_obj[0].entry_dn], ou_list)

    def ResumePost(self, **kwargs):
        '''AD账号默认不删除、复职用户需启用AD账号；职级、职位、部门'''
        user = kwargs['Badge']
        name = kwargs['xname']
        alias_name = NameConvert(name)
        result = self.mail_obj.select_maildb()
        org_id = str(kwargs['SHIYEBU_new_ID'])
        try:
            mail_domain = Params['MailDomain']['enable_mail'][org_id]
        except:
            mail_domain = Params['MailDomain']['not_enable_mail'][org_id]
        if org_id in self.enable_mail_org_id:
            upn = alias_name + '@' + Params['MailDomain']['enable_mail'][org_id]
        user_obj = self.obj.Get_ObjectDN(type='user', sAMAccountName=user)
        if user_obj:
            old_user_dn = user_obj[0].entry_dn
            # self.obj.ChangePassword(user_obj[0].entry_dn, ResumePost_line['Badge'])
            if int(user_obj[0].userAccountControl[0]) != 512:
                new_user_dn, ou_dn = self.OrganizationalStructure(user, kwargs)
                self.obj.Enable_User(old_user_dn)
                # 把账号加入到通讯组
                self.obj.AddUserInGroups_PS(user, 'all.list')
                self.obj.AddUserInGroups_PS(user, org_id)
                self.obj.AddUserInGroups_PS(user, kwargs['depid_new_id'])
                self.obj.ChangePassword(user_obj[0].entry_dn, kwargs['Badge'])
                mail_user_dic = self.inject_dbdata(kwargs, result, ou_dn, org_id, upn, mail_domain)
                self.mail_obj.insert_update_maildb(**mail_user_dic)
                # 检查EHR提供的用户组织架构和AD域中的是否相同
                if old_user_dn != new_user_dn:
                    print('用户组织机构发生变化，正在移动')
                    res = self.obj.Move_Object(type='user', new_object_dn=ou_dn,
                                               sAMAccountName=kwargs['Badge'], cn=kwargs['Badge'])
                    if res is True:
                        logging.info("移动用户从 %s 到 %s 成功" % (old_user_dn, ou_dn))
                    else:
                        logging.error("移动用户从 %s 到 %s 失败" % (old_user_dn, ou_dn))

                # 检查用户的title和职级是否休要更新
                # title,telephoneNumber ,extensionAttribute2
                try:
                    if user_obj[0].title[0] != kwargs['JOB_NEW']:
                        self.obj.Update_User_Info(new_user_dn, title=kwargs['JOB_NEW'])
                except:
                    self.obj.Update_User_Info(new_user_dn, title=kwargs['JOB_NEW'])

                try:
                    if user_obj[0].telephoneNumber[0] != str(kwargs['tel']):
                        self.obj.Update_User_Info(new_user_dn, telephoneNumber=kwargs['tel'])
                except:
                    self.obj.Update_User_Info(new_user_dn, telephoneNumber=kwargs['tel'])
                try:
                    if user_obj[0].extensionAttribute2[0] != str(kwargs['JOBGRADENEW']):
                        self.obj.Update_User_Info(new_user_dn,
                                                  extensionAttribute2=kwargs['JOBGRADENEW'])
                except:
                    self.obj.Update_User_Info(new_user_dn,
                                              extensionAttribute2=kwargs['JOBGRADENEW'])

                try:
                    if user_obj[0].userPrincipalName[0].split('@')[1] != mail_domain:
                        self.obj.Update_User_Info(new_user_dn, extensionAttribute1=mail_domain)
                        self.obj.Update_User_Info(new_user_dn, mail='%s@%s' % (
                            user_obj[0].userPrincipalName[0].split('@')[0], mail_domain))
                        self.obj.Update_User_Info(new_user_dn, userPrincipalName='%s@%s' % (
                            user_obj[0].userPrincipalName[0].split('@')[0], mail_domain))
                except:
                    self.obj.Update_User_Info(new_user_dn, extensionAttribute1=mail_domain)
                    self.obj.Update_User_Info(new_user_dn, mail='%s@%s' % (
                        user_obj[0].userPrincipalName[0].split('@')[0], mail_domain))
                    self.obj.Update_User_Info(new_user_dn, userPrincipalName='%s@%s' % (
                        user_obj[0].userPrincipalName[0].split('@')[0], mail_domain))

        else:
            # 超过三个月复制、账号可能已被人工删除、需重新创建
            new_user_dn, ou_dn = self.OrganizationalStructure(user, kwargs)
            self.obj.Add_User(ou_dn, name, user, kwargs['tel'],
                              kwargs['JOB_NEW'], kwargs['JOBGRADENEW'], domain=mail_domain,
                              alias=alias_name)
            time.sleep(1)
            # 把账号加入到通讯组
            self.obj.AddUserInGroups_PS(user, 'all.list')
            self.obj.AddUserInGroups_PS(user, org_id)
            self.obj.AddUserInGroups_PS(user, kwargs['depid_new_id'])
            time.sleep(1)
            mail_user_dic = self.inject_dbdata(kwargs, result, ou_dn, org_id, upn, mail_domain)
            self.mail_obj.insert_update_maildb(**mail_user_dic)
        # 启用Exchange邮箱
        if org_id in self.enable_mail_org_id:
            res = exchange.GetMailbox(user)
            if res.get('code') != 0:
                exchange.EnableMailboxhigh(user, alias=alias_name, database=result[0])

                # 设置Exchange读取属性
                self.obj.AclExchangeReadField(kwargs['Badge'])

    def PersonnelChange(self, **kwargs):
        user = kwargs['Badge']
        org_id = kwargs['SHIYEBU_new_ID']
        new_user_dn, ou_dn = self.OrganizationalStructure(user, kwargs)
        user_obj = self.obj.Get_ObjectDN(type='user', sAMAccountName=user)
        if user_obj:
            old_user_dn = user_obj[0].entry_dn
            # 组织变更
            if old_user_dn != new_user_dn:
                print("正在进行人事变更")

                # 更新群组,根据架构的调整更新用户隶属于的组（添加组、删除组）
                user_groups_list = self.obj.GetUserGroups(kwargs['Badge'])
                first_group_dn = 'CN=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
                    org_id, kwargs['SHIYEBU_new'])
                five_group_dn = 'CN=%s,OU=%s,OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
                    kwargs['depid_new_id'], kwargs['DEP_NEW'],
                    kwargs['zuzhiquyu_NEW'],
                    kwargs['zuzhishengfen_new'],
                    kwargs['zuzhifenqu_new'], kwargs['SHIYEBU_new'])

                try:
                    for i in user_groups_list:
                        groupname = i.split(',')[0].split('=')[1]
                        if is_number(groupname):
                            if len(groupname) == 3:  # 一级部门
                                if groupname != org_id:
                                    self.obj.RemoveUserOutGroups([user_obj[0].entry_dn], [i])  # 移除旧组
                                    self.obj.AddUserInGroups([user_obj[0].entry_dn],
                                                             [first_group_dn])  # 添加到新组
                            else:  # 其他非一级部门
                                if groupname != kwargs['depid_new_id']:
                                    self.obj.RemoveUserOutGroups([user_obj[0].entry_dn], [i])  # 移除旧组
                                    self.obj.AddUserInGroups([user_obj[0].entry_dn],
                                                             [five_group_dn])  # 添加到新组
                    self.obj.Move_Object(type='user', new_object_dn=ou_dn,
                                         sAMAccountName=kwargs['Badge'],
                                         cn=kwargs['Badge'])
                except:
                    pass

            # 职位变更
            try:
                if user_obj[0].title[0] != kwargs['JOB_NEW']:
                    self.obj.Update_User_Info(new_user_dn, title=kwargs['JOB_NEW'])
            except:
                self.obj.Update_User_Info(new_user_dn, title=kwargs['JOB_NEW'])

            # 职级变更
            try:
                if user_obj[0].extensionAttribute2[0] != kwargs['JOBGRADENEW']:
                    self.obj.Update_User_Info(new_user_dn,
                                              extensionAttribute2=kwargs['JOBGRADENEW'])
            except:
                self.obj.Update_User_Info(new_user_dn,
                                          extensionAttribute2=kwargs['JOBGRADENEW'])

    def Increment_user(self,active):
        # 用户增量导入
        Induction_list = []
        LeaveOffice_list = []
        ResumePost_list = []
        PersonnelChange_list = []

        user_obj = UserApi.EHR(6)
        user_data = user_obj.get_data()
        for user_line in user_data['Result']['empadd']['Row']:
            if hours(user_line['xClosedTime']) < 8:
                if "入职" in user_line['xtype']:
                    Induction_list.append(user_line)
                elif "离职" in user_line['xtype']:
                    LeaveOffice_list.append(user_line)
                elif "复职" in user_line['xtype']:
                    ResumePost_list.append(user_line)
                elif "人事变更" in user_line['xtype']:
                    PersonnelChange_list.append(user_line)

        if active=="Induction":
            ProcessPool=Pool(10)
            for i in Induction_list:
                ProcessPool.apply_async(self.Induction,kwds=i)
            ProcessPool.close()
            ProcessPool.join()

    def adduserinmailgroup(self):
        '''
        :return:
        '''
        obj_data = UserApi.EHR(7)
        org_data = obj_data.get_data()
        org_data = org_data['Result']['org']['Row']

        ou_list = []
        first_temp_dic = {'all.list':'OU=精锐教育集团,DC=onesmart,DC=local'}
        five_temp_dic = {}
        for org_line in org_data:
            if org_line['type'] == '部门':

                first_ou_dn = 'OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (org_line['shiyebu'])
                five_ou_dn = 'OU=%s,OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
                            org_line['title'], org_line['zuzhiquyu'], org_line['zuzhishenfen'], org_line['zuzhifenqu'],
                            org_line['shiyebu'])
                first_temp_dic[org_line['shiyebuid']] = first_ou_dn
                five_temp_dic[org_line['Uniqueid']] = five_ou_dn
                if first_temp_dic not in ou_list:
                    ou_list.append(first_temp_dic)
                if five_temp_dic not in ou_list:
                    ou_list.append(five_temp_dic)

        self.obj.AddUserInGroups_Muit_PS(ou_list)

    def create_special_account(self):
        get_data = ExcelData(r'D:\LdapAutoSyn\excel\onesmart.org特殊邮箱清单.xlsx','特殊邮箱清单')
        data = get_data.readExcel()
        f = open("./special_users.txt", 'a')
        for i in data:
            user = i['通讯录显示名'].split('@')[0]+"\n"
            org = 'OU=特殊账号,OU=精锐教育集团,DC=onesmart,DC=local'
            self.obj.Add_User(org,i['用户姓名'],i['通讯录显示名'].split('@')[0],domain='onesmart.org',alias=i['通讯录显示名'].split('@')[0])
            f.write(user)
        f.close()

    def CreateMailGrupAndAddUser(self):
        '''
            根据excel中的数据创建历年积累下来的邮件组，并根据邮件组中所包含的用户添加到组中
        :return:
        '''

        for filename in Params['ExcelFileList']:
            print(filename)
            file_path = r"%s\%s" %(Params['ExcelPath'],filename)
            data = xlrd.open_workbook(file_path)
            for sheetname in data.sheet_names():
                group_name = sheetname.split('@')[0]
                group_dn = "CN=%s,OU=特殊邮件组,OU=精锐教育集团,DC=onesmart,DC=local" % group_name
                print(group_dn)
                res = exchange.EnableDistributionGroup(group_name)
                if res.get('code') == 0:
                    logging.info("用户" + group_name + "邮箱启用成功")
                excel_obj = ExcelData(file_path, sheetname)
                user_list = []
                for linedata in excel_obj.readExcel():
                    upn = linedata['*工号']
                    name = linedata['姓名']
                    user_exist = self.obj.Get_ObjectDN(type='user', sAMAccountName=upn)
                    if user_exist:
                        res = self.obj.AddUserInGroups([user_exist[0].entry_dn], [group_dn])
                        if res:
                            print(upn +"添加成功")
                        else:
                            print([user_exist[0].entry_dn], [group_dn])
                        time.sleep(0.5)
                    else:
                        print(upn,name)

    def CreateHeadoffice2161(self):
        file_path = r"%s\%s" % (Params['ExcelPath'], Params['HeadOffice2161'])
        data = xlrd.open_workbook(file_path)
        for sheetname in data.sheet_names():
            group_name = sheetname.split('@')[0]
            group_dn = "CN=%s,OU=特殊邮件组,OU=精锐教育集团,DC=onesmart,DC=local" % group_name
            self.obj.Add_User_Group(group_name, group_dn, -2147483640, sheetname)
            res = exchange.EnableDistributionGroup(group_name)
            if res.get('code') == 0:
                logging.info("通讯安全组" + group_name + "邮箱启用成功")
            else:
                logging.error("通讯安全组" + group_name + "邮箱启用失败")

            excel_obj = ExcelData(file_path, sheetname)
            user_list = []
            for linedata in excel_obj.readExcel():
                email = linedata['address']
                user_exist = self.obj.Get_ObjectDN(type='user', userPrincipalName=email)
                if not user_exist:
                    print(email)
                    res = exchange.AddDistributionGroupMember(identity="headoffice2161.list", mailaddress=email)
                    if res.get('code') != 0:
                        print(res)
                res = exchange.GetMailbox(email.split('@')[0])
                if res.get('code') == 0:
                    user_exist = self.obj.Get_ObjectDN(type='user', userPrincipalName=email)
                    if user_exist:
                        user_list.append(user_exist[0].entry_dn)
                    else:
                        res = exchange.NewMailContacthight(Name=email.split('@')[0],ExternalEmailAddress=email)
                        if res.get('code') == 0:
                            time.sleep(2)
                            exchange.AddDistributionGroupMember(identity=email.split('@')[0],mailaddress=email)
            self.obj.AddUserInGroups(user_list, [group_dn])

    def CheckEmailUser(self):
        not_enable_list = ['103','378','401','599','339']
        # RTC账号清单获取，用于测试
        file_path = r"%s\%s" % (Params['ExcelPath'], 'rtc.xlsx')
        excel_obj = ExcelData(file_path, 'user')
        user_list = []
        for i in excel_obj.readExcel():
            user_list.append(i['jobnumber'])

        user_obj = UserApi.EHR(5)
        user_data = user_obj.get_data()
        for user_line in user_data['Result']['emp']['Row']:

            if user_line['email'] and '@onesmart.org' in user_line['email'] and user_line['Badge'] not in user_list:
                if "<" in user_line['email'] and ">" in user_line['email']:
                    email = user_line['email'].split('<')[1].split('>')[0]
                else:
                    email = user_line['email']
                is_mail_exsit = self.obj.Get_ObjectDN(type='user',base_dn='OU=精锐教育集团,DC=onesmart,DC=local',sAMAccountName=user_line['Badge'],mail=email)
                is_user_exsit = self.obj.Get_ObjectDN(type='user',base_dn='OU=精锐教育集团,DC=onesmart,DC=local',sAMAccountName=user_line['Badge'])
                if not is_mail_exsit and is_user_exsit and str(user_line['shiyebuid']) not in not_enable_list:
                    print(user_line['shiyebu'],user_line['Badge'],user_line['email'])

    def CheckUserPassword(self):

        user_obj = UserApi.EHR(5)
        user_data = user_obj.get_data()
        for user_line in user_data['Result']['emp']['Row']:
            password = '%s@Jr2020' % str(user_line['Badge'])[3:]
            if "<" in user_line['email'] and ">" in user_line['email']:
                email = user_line['email'].split('<')[1].split('>')[0]
            else:
                email = user_line['email']

            self.obj.authenticate_user(email,password)

    def DisableUser(self):
        '''
            批量禁用AD账号及Exchange邮箱（mailbox）
        :return:
        '''
        file_path = r"%s\%s" % (Params['ExcelPath'], '禁用的用户-2020-10-14.xlsx')
        data = xlrd.open_workbook(file_path)
        for sheetname in data.sheet_names():
            excel_obj = ExcelData(file_path, sheetname)
            for linedata in excel_obj.readExcel():
                uid = str(linedata['sAMAccountName'])
                user_exist = self.obj.Get_ObjectDN(type='user', sAMAccountName=uid)
                if user_exist:
                    print(user_exist[0].entry_dn)
                    # 禁用邮箱
                    res = exchange.DisableMailbox(uid)
                    if res.get('code') == 0:
                        print("禁用邮箱" + uid + "成功")
                    else:
                        print("禁用邮箱" + uid + "失败")

                    # 禁用AD域账号
                    self.obj.Disable_User(user_exist[0].entry_dn, uid.strip())

                    # 移除相关组，只保留Domain Users组
                    ou_list = self.obj.GetUserGroups(uid.strip())
                    print(ou_list)
                    if ou_list:
                        self.obj.RemoveUserOutGroups([user_exist[0].entry_dn], ou_list)

    def CreateMailUser(self,jobnumber,maildomain):
        user_obj = UserApi.EHR(5)
        user_data = user_obj.get_data()
        for i in user_data['Result']['emp']['Row']:
            try:
                if jobnumber == i['Badge']:
                    print("正在创建%s域账号，请稍等……"%jobnumber)
                    self.result = self.mail_obj.select_maildb()
                    user=i['Name1']
                    org = 'OU=%s,OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
                        i['DEPTITLE'], i['zuzhiquyu'], i['zuzhishenfen'],
                        i['zuzhifenqu'], i['shiyebu'])
                    # if i['email']:
                    #     alias_name = i['email'].split('@')[0]
                    # else:
                    #     alias_name = NameConvert(user)
                    alias_name = NameConvert(user)
                    print(user,alias_name)
                    self.obj.Add_User(org,user,jobnumber,mobile=i['mobile'],title=i['JOBTITLE'],jobgrade=i['jobgrade'],domain=maildomain,alias=alias_name)
                    # 把账号加入到通讯组
                    print("正在添加邮箱群组，请稍等……")
                    self.obj.AddUserInGroups_PS(jobnumber, 'all.list')
                    self.obj.AddUserInGroups_PS(jobnumber, i['shiyebuid'])
                    self.obj.AddUserInGroups_PS(jobnumber, i['DepID'])
                    print("正在设置邮箱权限，请稍等……")
                    self.obj.AclExchangeReadField(jobnumber)
                    print("正在启用邮箱请稍等……")
                    res = exchange.EnableMailboxhigh(user, alias=alias_name, database=self.result[0])
                    if res.get('code') == 0:
                        print("邮箱启用成功，地址为：%s@%s"%(alias_name,maildomain))
            except Exception as e:
                print("账号创建失败：%s"%e)

    def xxdq(self):
        data_obj = self.obj.Get_ObjectDN(type='user',base_dn='OU=小小地球,OU=精锐教育集团,DC=onesmart,DC=local')
        ig_list =['2400519','9900256','0302479','0202216','0301723','0301201','9300409','9500001','0201514','0302716']
        for i in data_obj:
            if i['extensionAttribute1'] == "ftkenglish.com":
                if i['mail']:# and i['sAMAccountName'] not in ig_list:
                    print(i['displayName'],i['sAMAccountName'],i.entry_dn,i['mail'],i['pwdLastSet'])

    def xxdq2(self):
        user_obj = UserApi.EHR(5)
        user_data = user_obj.get_data()
        for i in user_data['Result']['emp']['Row']:

            if i['email'] and i['shiyebuid'] == 401:
                # if i['email'].strip() != data_obj[0]['mail'] and "ftkenglish.com" in i['email']:
                a=['2402868', '2402831', '2402834', '2402865', '2402841', '2402864', '2400208', '2401022', '2402867', '9903996', '2402817', '2401947', '2400253', '2402824', '2402839', '2402866', '2402857', '2402828', '2402818', '2402842', '2402837', '2402832', '2402852', '2402861', '2402840', '2402820', '2402833', '2402815', '2402823', '2402822', '2402836', '2402819', '2402853', '2402826', '2402814', '2402827', '2402830', '2402855', '2402856', '2402854', '2402821', '2402829', '2402860', '2402862', '2400551', '2400484', '9901559', '2402825', '2400748', '0219377', '2402816', '2402838']
                if i['Badge'] in a:
                    data_obj = self.obj.Get_ObjectDN(type='user', base_dn='OU=小小地球,OU=精锐教育集团,DC=onesmart,DC=local',
                                                         sAMAccountName=i['Badge'])
                    print(i['Badge'],i['jobgrade'])
                    if data_obj:
                        self.obj.Update_User_Info(data_obj[0].entry_dn, extensionAttribute2=i['jobgrade'])


    def JrseGradeFIleInGroup(self):
        '''
            精锐少儿5级及以上上有人员加入到jrse5.list@onesmart.rg群组。
        :return:
        '''
        shiyebu_list = [378,103,401,148,652]
        user_obj = UserApi.EHR(5)
        user_data = user_obj.get_data()
        user_list = []
        for i in user_data['Result']['emp']['Row']:
            if i['shiyebuid'] in shiyebu_list:
                if int(i['jobgrade']) >=5:
                    userinfo = self.obj.Get_ObjectDN(type='user',base_dn='OU=%s,OU=精锐教育集团,DC=onesmart,DC=local'%(i['shiyebu']),sAMAccountName=i['Badge'])
                    if userinfo:
                        user_list.append(userinfo[0].entry_dn)

        groupinfo=self.obj.Get_ObjectDN(type='group',base_dn='OU=特殊邮件组,OU=精锐教育集团,DC=onesmart,DC=local',cn='jrse5.list')
        group_dn = groupinfo[0].entry_dn
        self.obj.AddUserInGroups(user_list, [group_dn])

    def CreateAllUser(self):
        user_obj = UserApi.EHR(5)
        user_data = user_obj.get_data()
        UserList=[]
        for i in user_data['Result']['emp']['Row']:
            data_obj = self.obj.Get_ObjectDN(type='user', base_dn='OU=精锐教育集团,DC=onesmart,DC=local',
                                             sAMAccountName=i['Badge'])
            jobnumber=i['Badge']
            maildomain='onesmart.org'
            if not data_obj:
                UserList.append(i)
                self.result = self.mail_obj.select_maildb()
                user = i['Name1']
                org = 'OU=%s,OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
                    i['DEPTITLE'], i['zuzhiquyu'], i['zuzhishenfen'],
                    i['zuzhifenqu'], i['shiyebu'])
                alias_name = NameConvert(user)
                self.obj.Add_User(org, user, jobnumber, mobile=i['mobile'], title=i['JOBTITLE'], jobgrade=i['jobgrade'],
                                  domain=maildomain, alias=alias_name)
                self.obj.AclExchangeReadField(jobnumber)

        print(UserList)


    def SetPasswordTime(self):
        '''
            设置密码过期时间，特殊账号密码永不过期
        :return:
        '''
        user_obj=self.obj.Get_ObjectDN(type='user',base_dn='OU=特殊账号,OU=精锐教育集团,DC=onesmart,DC=local')
        for userinfo in user_obj:
            print(userinfo.entry_dn)
            self.obj.Update_User_Info(userinfo.entry_dn, userAccountControl=66048)

    def CheckMailAndUPN(self):
        '''
            检查mail地址和userPrincipalName是否一致
        :return:
        '''
        user_obj = UserApi.EHR(5)
        user_data = user_obj.get_data()
        for i in user_data['Result']['emp']['Row']:
            if i['shiyebuid']==103 or i['shiyebuid']==378:
                user_info = self.obj.Get_ObjectDN(type='user',sAMAccountName=i['Badge'])
                if user_info[0].userPrincipalName != user_info[0].mail:
                    print(user_info[0].entry_dn)
                else:
                    print(user_info[0].sAMAccountName,user_info[0].displayName,user_info[0].userPrincipalName,user_info[0].entry_dn)


    def CreateSpecialMailUser(self):
        file_path = r"%s\%s" % (Params['ExcelPath'], "地球校区邮件账号.xlsx")
        data = xlrd.open_workbook(file_path)
        for sheetname in data.sheet_names():
            excel_obj = ExcelData(file_path, sheetname)
            user_list = []
            for linedata in excel_obj.readExcel():
                desc = linedata['校区']
                upn = linedata['账号']
                user_exist = self.obj.Get_ObjectDN(type='user', sAMAccountName=upn.split('@')[0])
                org="OU=特殊账号,OU=精锐教育集团,DC=onesmart,DC=local"

                if user_exist:
                    print(user_exist[0].entry_dn)
                    self.obj.Update_User_Info(user_exist[0].entry_dn, userAccountControl=66048)
                    self.obj.ChangePassword(org,upn.split('@')[0])

    def UpdateDepartment(self):
        '''
            更新用户部门信息
        :return:
        '''
        user_obj = UserApi.EHR(5)
        user_data = user_obj.get_data()
        for user_line in user_data['Result']['emp']['Row']:
            org = 'OU=%s,OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
                    user_line['DEPTITLE'], user_line['zuzhiquyu'], user_line['zuzhishenfen'], user_line['zuzhifenqu'],
                    user_line['shiyebu'])
            department_info = "%s/%s/%s/%s/%s" %(user_line['shiyebu'],user_line['zuzhifenqu'],user_line['zuzhishenfen'],user_line['zuzhiquyu'],user_line['DEPTITLE'])
            if "." in department_info and "|" in department_info:
                department_info=department_info.replace(".",'').replace("|",'')
            user_info = self.obj.Get_ObjectDN(type='user',base_dn="OU=精锐教育集团,DC=onesmart,DC=local",sAMAccountName=user_line['Badge'])
            if user_info:
                user_org = user_info[0].entry_dn
                self.obj.Update_User_Info(user_org, department=department_info)
                print(user_line['Badge'], department_info)

    def DisableYoubihuiMail(self):
        '''
            禁用优毕慧邮箱的AD账号，邮箱暂时保留
        :return:
        '''
        file_path = r"%s\%s" % (Params['ExcelPath'], "优毕慧在职邮箱.XLSX")
        data = xlrd.open_workbook(file_path)
        for sheetname in data.sheet_names():
            excel_obj = ExcelData(file_path, sheetname)
            user_list = []
            for linedata in excel_obj.readExcel():
                upn = linedata['电子邮件']
                user_exist = self.obj.Get_ObjectDN(type='user', userPrincipalName=upn)
                if user_exist:
                    print(user_exist[0].entry_dn,user_exist[0].sAMAccountName)
                    self.obj.Disable_User(user_exist[0].entry_dn,user_exist[0].sAMAccountName)

if  __name__ == '__main__':
    test = HandleArgvs()
    # test.operations_ou_full()
    test.operations_ou_diff()
    # test.operations_user(5)
    # test.operations_user(6,active='Induction')
    # test.operations_user(6,active='LeaveOffice')
    # test.operations_user(6,active='ResumePost')
    # test.operations_user(6,active='PersonnelChange')

    # test.adduserinmailgroup()
    # test.enabile_exchange_user()
    # test.create_special_account()
    # test.CreateMailGrupAndAddUser()
    # test.CreateHeadoffice2161()

    # test.CreateRTCUserAndEmail()
    # test.CheckEmailUser()
    # test.DisableUser()

    # test.xxdq()
    # test.xxdq2()

    # test.JrseGradeFIleInGroup()
    #test.CreateAllUser()
    # test.SetPasswordTime()
    # test.CheckMailAndUPN()
    # test.CreateSpecialMailUser()
    # test.UpdateDepartment()
    # test.DisableYoubihuiMail()