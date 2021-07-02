# -*- coding:utf-8 -*-
__author__ = 'Rainbower'

import time
import logging
from config.settings import Params
from api import ExchangeApi as exchange, UserApi, Adapi
from api.SynchronizationAAD import SyncAAD
from api.ExchangeOnlineApi import ExchangeOnline
from lib.ChineseToPinyin import ChineseToPinyin
from lib.MailUserDistribution import MailUserDistribution
from plugins.OpertionsToExcel import ExcelData

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
        org_description_dic = {}

        org_description_dic[1] = org_line['shiyebuid']
        org_description_dic[2] = org_line['zuzhifenquid']
        org_description_dic[3] = org_line['zuzhishenfenid']
        org_description_dic[4] = org_line['zuzhiquyuid']
        org_description_dic[5] = org_line['Uniqueid']

        org_dn_dic[1] = 'OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (org_line['shiyebu'])
        org_dn_dic[2] = 'OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (org_line['zuzhifenqu'], org_line['shiyebu'])
        org_dn_dic[3] = 'OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
        org_line['zuzhishenfen'], org_line['zuzhifenqu'], org_line['shiyebu'])
        org_dn_dic[4] = 'OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
        org_line['zuzhiquyu'], org_line['zuzhishenfen'], org_line['zuzhifenqu'], org_line['shiyebu'])
        org_dn_dic[5] = 'OU=%s,OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
        org_line['title'], org_line['zuzhiquyu'], org_line['zuzhishenfen'], org_line['zuzhifenqu'], org_line['shiyebu'])

        # print(org_description_dic[5],org_dn_dic[5])

        number = 1
        while number < 6:
            base_dn = org_dn_dic[number]
            is_exist = self.obj.Get_ObjectDN(type='org', base_dn='OU=精锐教育集团,DC=onesmart,DC=local',
                                             description=org_description_dic[number])  # 获取组织单位

            if not is_exist:
                # 增加组织单位
                # print(org_line)
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

                elif number == 5:
                    # 五级部门下通讯组名称为：四级部门名称+五级部门名称，避免因重名而创建失败
                    group_name = org_line['Uniqueid']
                    group_dn = 'CN=%s,%s' % (group_name, org_dn_dic[5])
                    group_is_exist = self.obj.Get_ObjectDN(type='group', base_dn=org_dn_dic[5],
                                                           description=org_description_dic[5])
                    if not group_is_exist:
                        self.obj.Add_User_Group(group_name, group_dn, -2147483640, org_description_dic[5])

            else:  # 组织单位已存在、可能需要修改或需要移动、或不做操作

                # 修改组织单位
                old_org_dn = is_exist[0].entry_dn
                old_org_name = old_org_dn.split(',')[0].split('=')[1]
                new_org_name = org_dn_dic[number].split(',')[0].split('=')[1]
                if old_org_name != new_org_name:
                    print('正在重命名组织单位：%s 为%s' % (old_org_name, new_org_name))
                    self.obj.Rename_Org(old_org_dn, new_org_name)
                else:
                    print("old_org_dn:", old_org_dn, "new_org_dn:", base_dn)
                    if old_org_dn != org_dn_dic[number]:
                        print(old_org_dn, org_dn_dic[number])
                        result = self.obj.Move_Object(type='org', new_object_dn=org_dn_dic[number - 1],
                                                      ou=org_dn_dic[number].split(',')[0].split('=')[1],
                                                      description=org_description_dic[number])

                        if result is True:
                            logging.info("移动组织单位%s到%s成功" % (old_org_dn, org_dn_dic[number - 1]))
                        else:
                            logging.error("移动组织单位%s到%s失败" % (old_org_dn, org_dn_dic[number - 1]))
            number += 1


    def operations_ou(self,data=None):
        '''
            组织单位增、修改、移动
        :return:
        '''
        if data:
            self.create_or_update_ou(data)
        else:
            obj_data = UserApi.EHR(7)
            org_data = obj_data.get_data()
            org_data = org_data['Result']['org']['Row']
            # obj_ldap = ldaps.Ad_Opertions()

            for org_line in org_data:
                if org_line['type'] == '部门':
                    self.create_or_update_ou(org_line)

                # 已提炼为公共方法
                # org_dn_dic = {}
                # org_description_dic = {}
                #
                # org_description_dic[1] = org_line['shiyebuid']
                # org_description_dic[2] = org_line['zuzhifenquid']
                # org_description_dic[3] = org_line['zuzhishenfenid']
                # org_description_dic[4] = org_line['zuzhiquyuid']
                # org_description_dic[5] = org_line['Uniqueid']
                #
                # org_dn_dic[1] = 'OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (org_line['shiyebu'])
                # org_dn_dic[2] = 'OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (org_line['zuzhifenqu'],org_line['shiyebu'])
                # org_dn_dic[3] = 'OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (org_line['zuzhishenfen'], org_line['zuzhifenqu'],org_line['shiyebu'])
                # org_dn_dic[4] = 'OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % ( org_line['zuzhiquyu'], org_line['zuzhishenfen'], org_line['zuzhifenqu'],org_line['shiyebu'])
                # org_dn_dic[5] = 'OU=%s,OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (org_line['title'], org_line['zuzhiquyu'], org_line['zuzhishenfen'], org_line['zuzhifenqu'],org_line['shiyebu'])
                #
                # # print(org_description_dic[5],org_dn_dic[5])
                #
                # number = 1
                # while number < 6:
                #     base_dn = org_dn_dic[number]
                #     is_exist = self.obj.Get_ObjectDN(type='org',base_dn='OU=精锐教育集团,DC=onesmart,DC=local',description=org_description_dic[number]) #获取组织单位
                #
                #     if not is_exist:
                #         #增加组织单位
                #         # print(org_line)
                #         self.obj.Add_Org(org_dn=org_dn_dic[number],org_id=org_description_dic[number])
                #         # 创建通用通讯组，只在一级和五级部门下创建通讯组,已组织单位唯一ID命名
                #         if number == 1:
                #             group_name = org_line['shiyebuid']
                #             group_dn = 'CN=%s,%s' % (org_line['shiyebuid'], org_dn_dic[1])
                #             group_is_exist = self.obj.Get_ObjectDN(type='group', base_dn=org_dn_dic[1],
                #                                                    CN=group_name,
                #                                                    description=org_description_dic[1])
                #             if not group_is_exist:
                #                 self.obj.Add_User_Group(group_name, group_dn, 2, org_description_dic[1])
                #
                #         elif number == 5:
                #             # 五级部门下通讯组名称为：四级部门名称+五级部门名称，避免因重名而创建失败
                #             group_name = org_line['Uniqueid']
                #             group_dn = 'CN=%s,%s' % (group_name, org_dn_dic[5])
                #             group_is_exist = self.obj.Get_ObjectDN(type='group', base_dn=org_dn_dic[5],
                #                                                    description=org_description_dic[5])
                #             if not group_is_exist:
                #                 self.obj.Add_User_Group(group_name, group_dn, 2, org_description_dic[5])
                #
                #     else: #组织单位已存在、可能需要修改或需要移动、或不做操作
                #
                #         #修改组织单位
                #         old_org_dn = is_exist[0].entry_dn
                #         old_org_name = old_org_dn.split(',')[0].split('=')[1]
                #         new_org_name = org_dn_dic[number].split(',')[0].split('=')[1]
                #         if old_org_name != new_org_name:
                #             print('正在重命名组织单位：%s 为%s' %(old_org_name,new_org_name))
                #             self.obj.Rename_Org(old_org_dn,new_org_name)
                #         else:
                #             print("old_org_dn:",old_org_dn, "new_org_dn:",base_dn)
                #             if old_org_dn != org_dn_dic[number]:
                #                 print(old_org_dn,org_dn_dic[number])
                #                 result = self.obj.Move_Object(type='org', new_object_dn=org_dn_dic[number - 1],
                #                                      ou=org_dn_dic[number].split(',')[0].split('=')[1],
                #                                      description=org_description_dic[number])
                #
                #                 if result is True:
                #                     logging.info("移动组织单位%s到%s成功" %(old_org_dn,org_dn_dic[number-1]))
                #                 else:
                #                     logging.error("移动组织单位%s到%s失败" % (old_org_dn, org_dn_dic[number - 1]))
                #     number += 1

    def inject_dbdata(self, userdata):
        # print(userdata)
        mail_user_dic = {
            'maildb': self.result[0],
            # 'username': userdata['Name1'],
            'jobnumber': userdata['Badge'],
            'counter': self.result[1] + 1,
            'jobgrade': userdata['jobgrade'],
            # 'mobile': user_line['mobile'],
            # 'title': userdata['JOBTITLE'],
            'createtime': time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
            # 'email': user_line['email'],
            'organization':self.ou_dn,
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
            mail_user_dic['mobile'] = userdata['mobile']
        except:
            mail_user_dic['mobile'] = userdata['tel']
        else:
            mail_user_dic['mobile'] = 'unknown'

        try:
            mail_user_dic['gender'] = userdata['GENDER']
        except:
            mail_user_dic['gender'] = userdata['gender']

        if self.org_id in self.enable_mail_org_id:
            mail_user_dic['is_active'] = 0
            if userdata['email']:
                #全量添加用户时、邮件地址使用现有的
                if "<" in userdata['email'] and ">" in userdata['email']:
                    mail_user_dic['email'] = userdata['email'].split('<')[1].split('>')[0]
                else:
                    mail_user_dic['email'] = userdata['email']
            else:
                #增量添加账号时、邮箱地址需要拼接、已存在邮箱则+1 如：zhangshan1@onesmart.org、zhangsan2@onesmart.org
                mail_user_dic['email'] = self.upn

        else:
            mail_user_dic['is_active'] = 1
            if userdata['email']:
                mail_user_dic['email'] = userdata['email']
            else:
                mail_user_dic['email'] = self.upn

        return mail_user_dic

    def operations_user(self,funcid):
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



        if funcid == 5:
            # 用户全量导入
            for user_line in user_data['Result']['emp']['Row']:
                #['UserDB01','UserDB02','UserDB03','UserDB04','UserDB05','UserDB06','UserDB07','UserDB08','UserDB09','UserDB010','UserDB011','UserDB012']

                self.org_id = str(user_line['shiyebuid'])
                self.result = self.mail_obj.select_maildb()
                # f4 = open("../users/4users.txt", 'a')
                # f5 = open("../users/5users.txt", 'a')
                # if not user_line['email']:
                #     pass
                # else:
                #     result = self.mail_obj.select_maildb()
                #     user_str = "%s %s %s\n" % (user_line['Badge'], user_line['email'].split('@')[0],result[0])
                #
                #     mail_user_dic = {
                #             'maildb': result[0],
                #             'username': user_line['Name1'],
                #             'jobnumber': user_line['Badge'],
                #             'counter': result[1] + 1,
                #             'jobgrade': user_line['jobgrade'],
                #             # 'mobile': user_line['mobile'],
                #             'title': user_line['JOBTITLE'],
                #             'createtime':time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
                #             #'email': user_line['email'],
                #         }
                #     if "<" in user_line['email'] and ">" in user_line['email']:
                #         mail_user_dic['email'] = user_line['email'].split('<')[1].split('>')[0]
                #     else:
                #         mail_user_dic['email'] = user_line['email']
                #
                #     print(mail_user_dic)
                #     if org_id in enable_mail_org_id:
                #         mail_user_dic['is_active'] = 0
                #         if user_line['jobgrade'] < 5:
                #             f4.write(user_str)
                #         elif user_line['jobgrade'] >= 5:
                #             f5.write(user_str)
                #     else:
                #         mail_user_dic['is_active'] = 1
                #         if user_line['jobgrade'] < 5:
                #             f4.write(user_str)
                #         elif user_line['jobgrade'] >= 5:
                #             f5.write(user_str)
                #     self.mail_obj.insert_update_maildb(**mail_user_dic)



                # mail_user_dic = {
                #     'maildb': self.result[0],
                #     'username': user_line['Name1'],
                #     'jobnumber': user_line['Badge'],
                #     'counter': self.result[1] + 1,
                #     'jobgrade': user_line['jobgrade'],
                #     # 'mobile': user_line['mobile'],
                #     'title': user_line['JOBTITLE'],
                #     'createtime': time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
                #     'email': user_line['email'],
                # }
                # if "<" in user_line['email'] and ">" in user_line['email']:
                #     mail_user_dic['email'] = user_line['email'].split('<')[1].split('>')[0]
                # else:
                #     mail_user_dic['email'] = user_line['email']

                self.ou_dn = 'OU=%s,OU=%s,OU=%s,OU=%s,OU=%s,OU=精锐教育集团,DC=onesmart,DC=local' % (
                    user_line['DEPTITLE'], user_line['zuzhiquyu'], user_line['zuzhishenfen'],
                    user_line['zuzhifenqu'],user_line['shiyebu'])

                #创建用户
                org_id = str(user_line['shiyebuid'])
                # print(ou_dn, user_line['Name1'], user_line['Badge'],user_line['JOBTITLE'])
                # is_exist = self.obj.Get_ObjectDN(type='user', CN=user_line['Badge'])

                is_exist = self.obj.Get_ObjectDN(type='user', sAMAccountName=user_line['Badge'])


                #用户的UPN拼接
                if org_id in self.enable_mail_org_id:
                    if user_line['email']:
                        #EHR接口提供的数据中有邮件地址的用户，直接使用邮件地址拼接UPN地址
                        if "<" in user_line['email'] and ">" in user_line['email']:
                            email = user_line['email'].split('<')[1].split('>')[0]

                        else:
                            email = user_line['email']
                        alias_name = email.split('@')[0]
                else:
                    if user_line['email']:
                        # EHR接口提供的数据中有邮件地址的用户，直接使用邮件地址拼接UPN地址
                        if "<" in user_line['email'] and ">" in user_line['email']:
                            email = user_line['email'].split('<')[1].split('>')[0]
                        else:
                            email = user_line['email']
                        alias_name = email.split('@')[0]
                    else:
                        # EHR接口提供的数据中没有邮件地址的账户
                        alias_name = ChineseToPinyin(user_line['Name1'])
                        try:
                            self.mail_user_number = self.mail_obj.select_email(alias_name)[0]
                        except:
                            self.mail_user_number = self.obj.Get_ObjectDN(type='user', mail='%s*' % alias_name)

                        # print(type(self.mail_user_number),self.mail_user_number)
                        # 邮箱前缀和UPN确保唯一
                        if self.mail_user_number > 0:
                            alias_name = alias_name + str(self.mail_user_number + 1)

                try:
                    self.upn = alias_name + '@' + Params['MailDomain']['enable_mail'][org_id]
                except:
                    self.upn = alias_name + '@' + Params['MailDomain']['not_enable_mail'][org_id]

                mail_user_dic = self.inject_dbdata(user_line)
                # print(mail_user_dic)
                # print(self.upn,mail_user_dic)
                # if mail_user_dic['jobnumber'] == '1700149':
                #     print(org_id,self.enable_mail_org_id)
                #     print(self.upn, mail_user_dic)


                # if user_line['DepID'] != 5486:
                #     self.obj.Delete_Object(type='user', sAMAccountName=user_line['Badge'])

                if not is_exist:

                    if user_line['DepID'] != 5486:
                        if org_id in self.enable_mail_org_id: #需要启用邮箱的用户
                            mail_domain = Params['MailDomain']['enable_mail'][org_id]
                            self.obj.Add_User(self.ou_dn, user_line['Name1'], user_line['Badge'],user_line['mobile'],user_line['JOBTITLE'],user_line['jobgrade'],domain=mail_domain,alias=self.upn.split('@')[0])
                            # alias_name = user_line['email'].split('@')[0]
                            self.mail_obj.insert_update_maildb(**mail_user_dic)

                            # time.sleep(3)
                            # if not user_line['email']:
                            #     pass
                            # else:
                            #     if org_id in self.enable_mail_org_id:
                            #         # mail_user_dic['is_active'] = 0
                            #         alias_name = user_line['email'].split('@')[0]
                            #         mail_user_dic = self.inject_dbdata(user_line)
                            #         self.mail_obj.insert_update_maildb(**mail_user_dic)
                            #
                            #         #启用邮箱
                            #         # if user_line['jobgrade'] > 0:
                            #         #     res = exchange.EnableMailboxhigh(user_line['Badge'],alias=alias_name,database=self.result[0])
                            #         #     if res.get('code') == 0:
                            #         #         mail_user_dic['is_active'] = 0
                            #         #         logging.info("用户" + user_line['Badge'] + "邮箱启用成功")
                            #         #         self.mail_obj.insert_update_maildb(**mail_user_dic)
                            #         #     else:
                            #         #         logging.error("用户" + user_line['Badge'] + "邮箱启用失败")
                            #         # elif user_line['jobgrade'] >= 5:
                            #         #     smtpaddress = "smtp:%s@jingruigroup.partner.mail.onmschina.cn" %alias_name
                            #         #     res = exchange.EnableRemoteMailboxhigh(user_line['Badge'],alias=alias_name,remoteroutingaddress=smtpaddress)
                            #         #     if res.get('code') == 0:
                            #         #         mail_user_dic['is_active'] = 0
                            #         #         logging.info("用户"+user_line['Badge']+"邮箱启用成功")
                            #         #         self.mail_obj.insert_update_maildb(**mail_user_dic)
                            #         #     else:
                            #         #         logging.error("用户" + user_line['Badge'] + "邮箱启用失败")

                        elif org_id in self.not_enable_mail_org_id:
                            mail_domain = Params['MailDomain']['not_enable_mail'][org_id]
                            self.obj.Add_User(self.ou_dn, user_line['Name1'], user_line['Badge'],user_line['mobile'],user_line['JOBTITLE'],user_line['jobgrade'],domain=mail_domain,alias=self.upn.split('@')[0])
                            # mail_user_dic['is_active'] = 1
                            self.mail_obj.insert_update_maildb(**mail_user_dic)

                else:
                    old_user_dn = is_exist[0].entry_dn
                    new_user_dn = "CN=%s,%s" %(user_line['Name1'],self.ou_dn)
                    if old_user_dn != new_user_dn:
                        # res = self.obj.Move_Object(type='user',new_object_dn=ou_dn,cn=user_line['Badge'])
                        res = self.obj.Move_Object(type='user',new_object_dn=self.ou_dn,cn=user_line['Name1'])
                        if res is True:
                            logging.info("移动用户从 %s 到 %s 成功" %(old_user_dn,self.ou_dn))
                        else:
                            logging.error("移动用户从 %s 到 %s 失败" %(old_user_dn,self.ou_dn))
            # f4.close()
            # f5.close()

        elif funcid == 6:
            # 用户增量导入
            for user_line  in user_data['Result']['empadd']['Row']:
                user = user_line['Badge']
                name = user_line['xname']
                alias_name = ChineseToPinyin(name)
                rank = user_line['JOBGRADENEW']
                user_obj = self.obj.Get_ObjectDN(type='user', sAMAccountName=user)
                if user_obj:
                    old_user_dn = user_obj[0].entry_dn

                # print(alias_name,old_user_dn)

                self.result = self.mail_obj.select_maildb()

                self.org_id = str(user_line['SHIYEBU_new_ID'])

                # mail_user_number = self.obj.Get_ObjectDN(type=user, mail='%s*' % alias_name)
                try:
                    self.mail_user_number = self.mail_obj.select_email(alias_name)[0]
                except:
                    self.mail_user_number = self.obj.Get_ObjectDN(type='user', mail='%s*' % alias_name)

                # print(type(self.mail_user_number),self.mail_user_number)
                #邮箱前缀和UPN确保唯一
                if self.mail_user_number > 0:
                    alias_name = alias_name+str(self.mail_user_number+1)



                result = self.mail_obj.select_maildb()

                if "入职" in user_line['xtype'] or "复职" in user_line['xtype'] or "人事变更" in user_line['xtype']:
                    org_dic = {
                        'shiyebuid': user_line['SHIYEBU_new_ID'],
                        'shiyebu': user_line['SHIYEBU_new'],
                        'zuzhifenquid': user_line['zuzhifenqu_new_ID'],
                        'zuzhifenqu': user_line['zuzhifenqu_new'],
                        'zuzhishenfenid': user_line['zuzhishengfen_new_ID'],
                        'zuzhishenfen': user_line['zuzhishengfen_new'],
                        'zuzhiquyuid': user_line['zuzhiquyu_NEW_id'],
                        'zuzhiquyu': user_line['zuzhiquyu_NEW'],
                        'Uniqueid': user_line['depid_new_id'],
                        'title': user_line['DEP_NEW'],
                    }
                    org_id = str(user_line['SHIYEBU_new_ID'])
                    # print(org_id,user_line)
                    # 根据事业部ID决定新建AD用户的UPN及邮箱后缀
                    try:
                        mail_domain = Params['MailDomain']['enable_mail'][org_id]
                    except:
                        mail_domain = Params['MailDomain']['not_enable_mail'][org_id]

                    if org_id in self.enable_mail_org_id:
                        self.upn = alias_name + '@' + Params['MailDomain']['enable_mail'][org_id]
                    # 获取用户隶属于的组织单位OU-DN
                    ou_is_exist = self.obj.Get_ObjectDN(type='org', OU=user_line['DEP_NEW'],
                                                        description=user_line['depid_new_id'])
                    if ou_is_exist:
                        self.ou_dn = ou_is_exist[0].entry_dn
                        new_user_dn = 'CN=%s,%s' % (name, self.ou_dn)
                    else:
                        self.operations_ou(org_dic)
                        self.ou_dn = "OU=%s,OU=%s,OU=%s,OU=%s,OU=%s,DC=onesmart,DC=local"%(org_dic['title'],org_dic['zuzhiquyu'],org_dic['zuzhishenfen'],org_dic['zuzhifenqu'],org_dic['shiyebu'])
                        new_user_dn = 'CN=%s,%s' %(name, self.ou_dn)
                    mail_user_dic = self.inject_dbdata(user_line)

                    # print(mail_user_dic)
                    if user_obj:
                        user_dn = user_obj[0].entry_dn




                if "入职" in user_line['xtype']:

                    if org_id in self.enable_mail_org_id:
                        # org_dic = {
                        #     'shiyebuid':user_line['SHIYEBU_new_ID'],
                        #     'shiyebu':user_line['SHIYEBU_new'],
                        #     'zuzhifenquid':user_line['zuzhifenqu_new_ID'],
                        #     'zuzhifenqu':user_line['zuzhifenqu_new'],
                        #     'zuzhishenfenid':user_line['zuzhishengfen_new_ID'],
                        #     'zuzhishenfen':user_line['zuzhishengfen_new'],
                        #     'zuzhiquyuid':user_line['zuzhiquyu_NEW_id'],
                        #     'zuzhiquyu':user_line['zuzhiquyu_NEW'],
                        #     'Uniqueid':user_line['depid_new_id'],
                        #     'title':user_line['DEP_NEW'],
                        # }
                        #检测组织架构是否存在
                        # self.operations_ou(org_dic)

                        # 根据事业部ID决定新建AD用户的UPN及邮箱后缀
                        # mail_domain = Params['MailDomain']['enable_mail'][org_id]

                        # print('1111111111111111',self.ou_dn, user_line['xname'], user_line['Badge'], user_line['tel'],
                        #                   user_line['JOB_NEW'], user_line['JOBGRADENEW'], mail_domain,alias_name)
                        # 创建AD账号
                        user_is_exist = self.obj.Get_ObjectDN(type='user',base_dn=self.ou_dn,sAMAccountName=user_line['Badge'])
                        if not user_is_exist:
                            self.obj.Add_User(self.ou_dn, user_line['xname'], user_line['Badge'], user_line['tel'],
                                              user_line['JOB_NEW'], user_line['JOBGRADENEW'], domain=mail_domain,
                                              alias=alias_name)
                            time.sleep(1)
                            # 把账号加入到通讯组
                            self.obj.AddUserInGroups_PS(user,'all.list')
                            self.obj.AddUserInGroups_PS(user,user_line['SHIYEBU_new_ID'])
                            self.obj.AddUserInGroups_PS(user,user_line['depid_new_id'])
                            time.sleep(2)

                            print(user_line)
                            mail_user_dic = self.inject_dbdata(user_line)
                            print(mail_user_dic)

                            # 启用邮箱
                            exchange.EnableMailboxhigh(user,alias=alias_name,database=result[0])
                            self.mail_obj.insert_update_maildb(**mail_user_dic)

                            # 以下代码是实现账号在线上线下开通
                            # if int(rank) < 5:  # 启用Exchange邮箱
                            #     # self.mail_obj.insert_update_maildb(dbname=result[0], jobnumber=user_line['Badge'],username=user_line['xname'], counter=result[1] + 1)
                            #
                            #     # self.enabile_exchange_user('enable_exchange_user',user, alias_name, mail_user_number,result[0])
                            #     exchange.EnableMailboxhigh(user,alias=alias_name,database=result[0])
                            #     self.mail_obj.insert_update_maildb(**mail_user_dic)
                            #
                            # elif int(rank) >= 5:  # 启用Office365邮箱
                            #     # self.enabile_exchange_user('enable_exchange_online_user',user, alias_name, mail_user_number,result[0])
                            #
                            #
                            #     smtpaddress = "smtp:%s@jingruigroup.partner.mail.onmschina.cn" % alias_name
                            #     res = exchange.EnableRemoteMailboxhigh(user, alias=alias_name,
                            #                                            remoteroutingaddress=smtpaddress)
                            #
                            #     # 通过AAD同步AD信息到exchangeonline上，时间约7秒
                            #     self.sync_aad.SyncUserinfoToOnline()
                            #     time.sleep(300)
                            #
                            #     # 邮箱授权
                            #     self.online_obj.AddAssignLicense(self.upn)
                            #
                            #     if res.get('code') == 0:
                            #         mail_user_dic['is_active'] = 0
                            #         logging.info("用户" + user_line['Badge'] + "邮箱启用成功")
                            #         self.mail_obj.insert_update_maildb(**mail_user_dic)
                            #     else:
                            #         logging.error("用户" + user_line['Badge'] + "邮箱启用失败")


                    elif org_id in self.not_enable_mail_org_id:
                        # mail_domain = Params['MailDomain']['not_enable_mail'][org_id]
                        # 创建AD账号
                        self.obj.Add_User(self.ou_dn, user_line['xname'], user_line['Badge'], user_line['tel'],
                                          user_line['JOB_NEW'], user_line['JOBGRADENEW'], domain=mail_domain,
                                          alias=alias_name)


                # elif "离职" in user_line['xtype']:
                #     #禁用站好、程序不自动不删除
                #
                #     print(user_line)
                #     # 禁用邮箱账号
                #     res = exchange.DisableMailUser(identity=user)
                #     if res.get('code') == 0:
                #         logging.info("用户" + user + "邮箱" + self.upn + "禁用成功")
                #     else:
                #         logging.error("用户" + user + "邮箱" + self.upn + "禁用失败")
                #
                #     # 禁用AD域账号
                #     self.obj.Disable_User(user_dn, user)
                #
                #     # #禁用Exchange账号、目前所有账号开在本地Exchange上，不需考线上线下问题
                #     # if int(rank) < 5:
                #     #     res = exchange.DisableMailUser(identity=user)
                #     #     if res.get('code') == 0:
                #     #         logging.info("用户" + user + "邮箱" + self.upn + "禁用成功")
                #     #     else:
                #     #         logging.error("用户" + user + "邮箱" + self.upn + "禁用失败")
                #     #
                #     # elif int(rank) >= 5:
                #     #     self.online_obj.DelAssignLicense(self.upn)
                #     #     res = exchange.DisableRemoteMailUser(identity=user)
                #     #     if res.get('code') == 0:
                #     #         logging.info("用户" + user + "邮箱" + self.upn + "禁用成功")
                #     #     else:
                #     #         logging.error("用户" + user + "邮箱" + self.upn + "禁用失败")
                #
                # elif "复职" in user_line['xtype']:
                #     '''
                #         AD账号默认不删除、复职用户需启用AD账号；职级、职位、部门
                #    '''
                #     print(user_line)
                #     # 启用AD域账号
                #     user_is_exsit = self.obj.Get_ObjectDN(type='user',base_dn='',sAMAccountName=user_line['Badge'])
                #
                #     #检查用户组织架构是否存在、是否需要更新、移动
                #     # self.operations_ou(org_dic)
                #     if user_is_exsit:
                #         self.obj.Enable_User(user_dn, CN=user)
                #
                #         # 检查EHR提供的用户组织架构和AD域中的是否相同
                #         if old_user_dn != new_user_dn:
                #             res = self.obj.Move_Object(type='user', new_object_dn=self.ou_dn, cn=user_line['Badge'])
                #             if res is True:
                #                 logging.info("移动用户从 %s 到 %s 成功" % (old_user_dn, self.ou_dn))
                #             else:
                #                 logging.error("移动用户从 %s 到 %s 失败" % (old_user_dn, self.ou_dn))
                #     else:
                #         #超过三个月复制、账号可能已被人工删除、需重新创建
                #         self.obj.Add_User(self.ou_dn, name,user, user_line['tel'],
                #                                       user_line['JOB_NEW'], user_line['JOBGRADENEW'], domain=mail_domain,
                #                                       alias=alias_name)
                #     #检查用户的title和职级是否休要更新
                #     #title,telephoneNumber ,extensionAttribute2
                #     if user_obj[0].title != user_line['JOB_NEW']:
                #         self.obj.Update_User_Info(new_user_dn,title=user_line['JOB_NEW'])
                #     elif user_obj[0].telephoneNumber != user_line['tel']:
                #         self.obj.Update_User_Info(new_user_dn,telephoneNumber=user_line['tel'])
                #     elif user_line[0].extensionAttribute2 != user_line['JOBGRADENEW']:
                #         self.obj.Update_User_Info(new_user_dn,extensionAttribute2=user_line['JOBGRADENEW'])
                #
                #     # 启用Exchange邮箱
                #     if org_id in self.enable_mail_org_id:
                #         exchange.EnableMailboxhigh(user,alias=alias_name,database=result[0])
                #         # if rank < 5:
                #         #     self.enabile_exchange_user('enable_exchange_user', user, alias_name, self.mail_user_number)
                #         # else:
                #         #     self.enabile_exchange_user('enable_exchange_online_user', user, alias_name, self.mail_user_number)
                #
                # elif "人事变更" in user_line['xtype']:
                #     #组织架构变更
                #     if old_user_dn != new_user_dn:
                #         res = self.obj.Move_Object(type='user', new_object_dn=self.ou_dn, cn=user_line['Badge'])
                #         if res is True:
                #             logging.info("移动用户从 %s 到 %s 成功" % (old_user_dn, self.ou_dn))
                #         else:
                #             logging.error("移动用户从 %s 到 %s 失败" % (old_user_dn, self.ou_dn))
                #
                #     #职位变更
                #     if user_obj[0].title != user_line['JOB_NEW']:
                #         self.obj.Update_User_Info(new_user_dn,title=user_line['JOB_NEW'])
                #
                #     # 职级变更
                #     if user_line[0].extensionAttribute2 != user_line['JOBGRADENEW']:
                #         self.obj.Update_User_Info(new_user_dn, extensionAttribute2=user_line['JOBGRADENEW'])

    def adduserinmailgroup(self):
        '''
        :return:
        '''
        obj_data = UserApi.EHR(7)
        org_data = obj_data.get_data()
        org_data = org_data['Result']['org']['Row']

        ou_list = []
        first_temp_dic = {'all.list':'OU=精锐在线教育集团,DC=onesmart,DC=local'}
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

    # def enabile_exchange_user(self,action,user,alias_name,mail_user_number,dbname):
    #     exchange_ojb = ExchangeApi.Exchange()
    #     if len(mail_user_number) > 0:
    #         mail_aliasname = alias_name + (len(mail_user_number) + 1)
    #         self.exchange_ojb.enable_user(action, user, mail_aliasname,dbname)
    #     else:
    #         self.exchange_ojb.enable_user(action, user, alias_name,dbname)


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

if  __name__ == '__main__':
    test = HandleArgvs()
    # test.operations_ou()
    # test.operations_user(5)
    test.operations_user(6)
    # test.adduserinmailgroup()

    # test.create_special_account()