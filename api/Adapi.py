# -*- coding:utf-8 -*-
__author__ = 'Rainbower'

import os
import time
import winrm
import logging
from ldap3.core.exceptions import LDAPBindError
from ldap3 import Server, Connection, ALL, NTLM, MODIFY_REPLACE, SUBTREE, ALL_ATTRIBUTES
from ldap3.extend.microsoft.addMembersToGroups import ad_add_members_to_groups
from ldap3.extend.microsoft.removeMembersFromGroups import ad_remove_members_from_groups

from config import settings

logdir='\\'.join(os.path.abspath(os.path.dirname(__file__)).split('\\')[:-1])
# 日志设置
LOG_FORMAT = "%(asctime)s  %(levelname)s  %(filename)s  %(lineno)d  %(message)s"
logging.basicConfig(filename='%s/logs/%s.log'%(logdir,time.strftime("%Y-%m-%d", time.localtime())), level=logging.INFO, format=LOG_FORMAT)

class Ad_Opertions():
    '''
    操作AD域的类
    '''
    def __init__(self,):
        '''
            建立ad连接
        '''
        self.domain = settings.Params['Adconfig']['domain']
        self.DC = ','.join(['DC=' + dc for dc in self.domain.split('.')]) # onesmart.local -> DC=onesmart,DC=local
        self.pre = self.domain.split('.')[0].upper()  # 用户登陆的前缀
        self.ip = settings.Params['Adconfig']['host']
        self.admin = settings.Params['Adconfig']['user']
        self.pwd = settings.Params['Adconfig']['password']
        self.server = Server(self.ip,port=636,use_ssl=True,get_info=ALL)
        self.conn = Connection(self.server, user=self.pre+'\\'+self.admin, password=self.pwd, auto_bind=True)
        self.win = winrm.Session('http://' + self.ip + ':5985/wsman',auth=(self.admin,self.pwd))

    def Get_ObjectDN(self, type=None, base_dn=None, **kwargs):
            """
            参考: https://ldap3.readthedocs.io/tutorial_searches.html
            获取用户dn或组织单位, 通过 args
            可以支持多个参数: Get_ObjectDN(type='user',mail="junpeng.du@onesmart.org", uid="1601016")，Get_User_OR_Org_DNInfo(type='org',ou="Andy1")
            会根据 kwargs 生成 search的内容，进行查询: 多个条件是 & and查询,返回第一个查询到的结果,建议使用唯一标识符进行查询,这个函数基本可以获取所有类型的数据
            :param type:
            :param base_dn:
            :param kwargs:
            :return:
            """

            search = ""
            Type_dic = {'user':'user','org':'OrganizationalUnit','group':'group'}
            for k, v in kwargs.items():
                search += "(%s=%s)" % (k, v)
            if not base_dn:
                base_dn = self.DC
            else:
                base_dn = base_dn
            if search:
                search_filter = '(&(objectclass=%s){})'.format(search) %Type_dic[type]
            else:
                search_filter = '(objectclass=%s)'%Type_dic[type]

            # print(search_filter)
            status = self.conn.search(search_base=base_dn,  # 在哪个组织单位中查询(可以不知道ou、但是dc必须指定)
                                      search_filter=search_filter,  # 搜索请求的过滤器
                                      search_scope=SUBTREE,  # 指定搜索上下文的范围
                                      attributes=ALL_ATTRIBUTES  # 查询数据的哪些属性
                                      # attributes=['sAMAccountName', 'displayName','mail','title','telephoneNumber', 'objectClass', 'userAccountControl','userAccountControl','userPrincipalName','extensionAttribute1','extensionAttribute2','pwdLastSet','msDS-UserPasswordExpiryTimeComputed','ou','descripion'],  # 查询数据的哪些属性
                                      # paged_size=20000  # 一次查询多少条
                                      )
            if status:
                return self.conn.entries
            # else:
            #     return False

    def Add_User(self, org, name, uid, mobile=None, title=None,jobgrade=None,domain=None, gidnumber=501,alias=None,):
        '''
        增加用户
        :param org:         组织单位
        :param name:        显示名
        :param uid:         用户名
        :param mobile:      电话
        :param jobgrade:    邮件域
        :param title:       职务
        :param gidnumber:   组ID
        :param alias:
        :return:
        '''

        password = '%s@Jr2020'% str(uid)[3:]
        # password = 'Onesmart@123'

        org_base = ','.join(['OU=' + ou for ou in org.split('.')]) + ',' + self.DC
        user_att = {
            'Description': name,                            #描述
            'displayName' : name,                           #显示名
            'sAMAccountName': uid,                          # 登录名
            'userAccountControl': '544',                    #启用账号
            'pwdLastSet': -1,                               #取消下次登录需要修改密码
            # 'Department':'',                              #部门
            # 'Company':'',                                 #公司
        }
        if not mobile or len(mobile)!=11:
                pass
        else:
            user_att['telephoneNumber'] = mobile

        if alias:
            user_att['userPrincipalName'] = "%s@%s" % (alias, domain.strip())
        else:
            user_att['userPrincipalName'] = "%s@%s" % (uid, domain.strip())

        if title:
            user_att['title'] = title
        if domain:
            user_att['extensionAttribute1'] = domain.strip()
        if jobgrade:
            user_att['extensionAttribute2'] = jobgrade
        res = self.conn.add('CN=%s,%s'%(uid,org), object_class='user',attributes=user_att)

        if res:
            self.conn.extend.microsoft.modify_password('CN=%s,%s' % (uid, org), password)
            self.conn.modify('CN=%s,%s' % (uid, org), {'userAccountControl': [('MODIFY_REPLACE', 512)]})

            logging.info('增加账号CN=%s,%s成功'%(uid,org))
            logging.info("%s的邮箱地址为：%s@%s"%(uid,alias,domain.strip()))
        else:
            print('增加账号 %s 发生错误：'%name,user_att, self.conn.result)
            logging.error('增加账号' + uid+ '发生错误：' + self.conn.result['message'])



    def Disable_User(self,user_dn, uid):
        """
        禁用账号
        user_dn：用户组织路径
        uid：账号
        """
        res = self.conn.modify(user_dn,{'userAccountControl': [(MODIFY_REPLACE, [514])]})
        if res:
            print('禁用账号 %s 成功！'%uid)
            logging.info('禁用账号 %s 成功！'%uid)
        else:
            print('禁用账号 %s 发生错误：'%uid, self.conn.result['description'])
            logging.error('禁用账号 %s 发生错误：'%uid, self.conn.result['description'])

    def Enable_User(self, user_dn):
        """
        启用账号
        org：增加到该组织下
        uid：账号
        """

        # org_base = 'CN=%s,'%uid + ','.join(['OU=' + ou for ou in org.split('.')]) + ',' + self.DC
        '''
            66048:启用账号并设置密码永不过期
            66050:禁用账号并设置密码永不过期
            512:启用账号
            514:禁用账号
        '''
        res = self.conn.modify(user_dn,{'userAccountControl': [(MODIFY_REPLACE, [512])]})
        if res:
            print('启用账号 %s 成功！'%user_dn)
            logging.info('启用账号 %s 成功！'%user_dn)
        else:
            print('启用账号 %s 发生错误：'%user_dn, self.conn.result['description'])
            logging.error('启用账号 %s 发生错误：'%user_dn, self.conn.result['description'])




    def Update_User_Info(self, user_dn, action=MODIFY_REPLACE, **kwargs):
        """
        更新用户属性
        :param dn: 用户dn 可以通过get_userdn_by_args，get_userdn_by_mail 获取
        :param action: MODIFY_REPLACE 对字段原值进行替换  MODIFY_ADD 在指定字段上增加值   MODIFY_DELETE 对指定字段的值进行删除
        :param kwargs: 要进行变更的信息内容 uid userPassword mail sn gidNumber uidNumber mobile title
        :return: True or False
        """
        allow_key = "uid userPassword mail sn gidNumber uidNumber mobile title extensionAttribute2 extensionAttribute1 telephoneNumber userPrincipalName userAccountControl department proxyAddresses".split(" ")
        update_args = {}
        for k, v in kwargs.items():
            if k not in allow_key:
                msg = "字段: %s, 不允许进行修改, 不生效" % k
                return False
            update_args.update({k: [(action, [v])]})
        status = self.conn.modify(user_dn, update_args)
        if status is True:
            logging.info("更新用户属性:%s, %s 成功" %(user_dn,update_args))
        else:
            logging.error("更新用户属性:%s, %s 失败" %(user_dn,update_args))
        return status

    def Update_User_CN(self, user_dn, new_user_cn):
        s = self.conn.modify_dn(user_dn, 'cn=%s' % new_user_cn)
        return s

    def Add_User_Group(self,group_name,group_dn,group_type,group_id):
        '''
        创建组，组作用域（本地、全局、通用）
        :param group_dn: 例：(CN=baby,DC=onesmart,DC=local)在onesmart.local下创建baby组，默认是全局安全组
        :param group_type: 全局安全组(-2147483646)、全局通讯组(2);通用安全组(-2147483640)、通用通讯组(8);本地域安全组(-2147483644)、本地域通讯组(4)
        :return:
        '''

        group_att= {
            'sAMAccountName': group_name,
            'groupType': group_type,
            'description': group_id,
        }
        status = self.conn.add(group_dn, object_class='group', attributes=group_att)
        if status:
            logging.info('增加邮件组：' + group_dn + " 成功！ ")
        else:
            # return False
            logging.info('增加邮件组：' + group_dn + " 失败！ ")


    def AddUserInGroups(self,userdn_list,groupdn_list):
        '''
        把单个用户或者多个用户一次性添加到组中
        :param userdn_list:  实例：['CN=1601016,OU=Andy,DC=onesmart,DC=lcoal','CN=1601017,OU=Andy,DC=onesmart,DC=lcoal']
        :param groupdn_list: 实例：['CN=MailGroup1,DC=onesmart,DC=local','CN=MailGroup2,DC=onesmart,DC=local']
        :return:
        '''

        status = ad_add_members_to_groups(self.conn,userdn_list,groupdn_list)
        if status:
            return True
        else:
            return False

    def AddUserInGroups_PS(self,username,groupname):
        '''
            通过调用powershell把用户添加到通讯组
        :param username:
        :param groupname:
        :return:
        '''

        try:
            command_lsit = ["Import-Module ActiveDirectory"]
            command_lsit.append("Add-ADGroupMember -Identity %s -Members %s" %(groupname,username))
            for i in command_lsit:
                result = self.win.run_ps(i)
                if result.status_code == 0:
                    logging.info("执行命令：" + i + " 成功!")
                else:
                    logging.info("执行命令：" + i + " 失败!")

        except Exception as e:
            logging.error(e)

    def AddUserInGroups_Muit_PS(self,ou_list):
        '''
            通过调用powershell把批量添加用户到通讯组
        :param ou_list:
        :return:
        '''

        result = self.win.run_ps("Import-Module ActiveDirectory")
        if result.status_code == 0:
            logging.info("执行命令：" + "Import-Module ActiveDirectory" + " 成功!")
        else:
            logging.info("执行命令：" + "Import-Module ActiveDirectory" + " 失败!")

        for ou in ou_list:
            print(ou)
            for k,v in ou.items():
                try:
                    commands = "Get-ADUser -Filter * -SearchBase '%s' | foreach {Add-ADGroupMember -Identity %s -Members $_.name}" %(v,k)

                    result = self.win.run_ps(commands)
                    if result.status_code == 0:
                        logging.info("执行命令：" + commands + " 成功!")
                    else:
                        logging.info("执行命令：" + commands + " 失败!")

                except Exception as e:
                    logging.error(e)


    def RemoveUserOutGroups(self,userdn_list,groupdn_list,fix=False):
        '''
        把单个用户或者多个用户一次性添加到组中
        :param userdn_list:  实例：['CN=1601016,OU=Andy,DC=onesmart,DC=lcoal','CN=1601017,OU=Andy,DC=onesmart,DC=lcoal']
        :param groupdn_list: 实例：['CN=MailGroup1,DC=onesmart,DC=local','CN=MailGroup2,DC=onesmart,DC=local']
        :param fix: 检查组中是否有需要操作的用户,True:不管是否成功都返回True,False:成功返回True失败返回False
        :return:
        '''

        status = ad_remove_members_from_groups(self.conn,userdn_list,groupdn_list,fix)
        if status:
            return True
        else:
            return False

    def Add_Org(self,org_dn=None,org_id=None):
        '''
        增加组织
        oorg: 组织，格式为：aaa.bbb 即bbb组织下的aaa组织，不包含域地址
        '''
        attr = {
            # 'co':'',            #国家/地区
            # 'st':'',            #省/自治区
            # 'l':'',             #市/县
            # 'street':'',        #街道
            # 'postalCode':'',    #邮编
            'description':org_id, #描述
            # 'name':'',          #名称
            # 'ou':'',            #组织单位

        }

        res = self.conn.add(org_dn, object_class='OrganizationalUnit',attributes=attr)  # 成功返回True，失败返回False
        if res:
            logging.info("增加组织" + org_dn + "成功！")
        else:
            logging.error("增加组织" + org_dn + "发生错误: " + self.conn.result['description'])

    def Org_Delete_Attribute(self,flag,*args):
        '''
        对组织单位设置"防止对象被意外删除"属性
        :param flag: 0 or 1,禁用或启用
        :param agrs: 实例：['OU=Andy,DC=onesmart,DC=local','OU=Andy,DC=onesmart,DC=local']
        :return:
        '''

        try:

            if args:
                enable_del = ["Import-Module ActiveDirectory"]
                disable_del = ["Import-Module ActiveDirectory"]
                for ou_dn in args:
                    enable_del.append("Get-ADOrganizationalUnit -SearchBase '%s' -filter *|Set-ADOrganizationalUnit -ProtectedFromAccidentalDeletion $false" % ou_dn)
                    disable_del.append("Get-ADOrganizationalUnit -SearchBase '%s' -filter *|Set-ADOrganizationalUnit -ProtectedFromAccidentalDeletion $true" % ou_dn)

                flag_map = {0: enable_del, 1: disable_del}
                for cmd in flag_map[flag]:
                    ret = self.win.run_ps(cmd)
                    if ret.status_code == 0:      # 调用成功 减少日志写入
                        if flag == 0:
                            logging.info("防止对象被意外删除×")
                        elif flag == 1:
                            logging.info("防止对象被意外删除√")
                    else:
                        return False
            else:
                enable_del = ["Import-Module ActiveDirectory",
                              "Get-ADOrganizationalUnit -filter * -Properties ProtectedFromAccidentalDeletion | where {"
                              "$_.ProtectedFromAccidentalDeletion -eq $true} |Set-ADOrganizationalUnit "
                              "-ProtectedFromAccidentalDeletion $false"]
                disable_del = ["Import-Module ActiveDirectory",
                               "Get-ADOrganizationalUnit -filter * -Properties ProtectedFromAccidentalDeletion | where {"
                               "$_.ProtectedFromAccidentalDeletion -eq $false} |Set-ADOrganizationalUnit "
                               "-ProtectedFromAccidentalDeletion $true"]
                flag_map = {0: enable_del, 1: disable_del}
                for cmd in flag_map[flag]:
                    ret = self.win.run_ps(cmd)
                    if ret.status_code == 0:# 调用成功 减少日志写入
                        if flag == 0:
                            logging.info("防止对象被意外删除×")
                        elif flag == 1:
                            logging.info("防止对象被意外删除√")
                    else:
                        return False


        except Exception as e:
            print(e)
            # logging.error(e)
            return False

    def Rename_Org(self,org_dn,newname):
        '''
            重命名组织单位
        :param org_dn:
        :param newname:
        :return:
        '''
        try:
            command_lsit = ["Import-Module ActiveDirectory"]
            command_lsit.append("Rename-ADObject '%s' -NewName '%s'" %(org_dn,newname))
            for i in command_lsit:
                result = self.win.run_ps(i)
                if result.status_code == 0:
                    logging.info("执行命令：" + i + " 成功!")
                else:
                    logging.info("执行命令：" + i + " 失败!")

        except Exception as e:
            logging.error(e)

    def Update_Org_Info(self, org_dn, action=MODIFY_REPLACE, **kwargs):
        """
        更新组织单位属性
        :param dn: 组织单位dn 可以通过get_userdn_by_args，get_userdn_by_mail 获取
        :param action: MODIFY_REPLACE 对字段原值进行替换  MODIFY_ADD 在指定字段上增加值   MODIFY_DELETE 对指定字段的值进行删除
        :param kwargs: 要进行变更的信息内容 uid userPassword mail sn gidNumber uidNumber mobile title
        :return:
        """
        allow_key = "co st l street postalCode description descripion name ou".split(" ")
        update_args = {}
        for k, v in kwargs.items():
            if k not in allow_key:
                msg = "字段: %s, 不允许进行修改, 不生效" % k
                print(msg)
                return False
            update_args.update({k: [(action, [v])]})
        status = self.conn.modify(org_dn, update_args)
        if status:
            logging.info("%s更新为%s成功" % (org_dn, update_args))
        else:
            logging.error("%s更新为%s失败" % (org_dn, update_args))
        return status

    def Update_dn(self,Old_dn=None,NewNmae=None):
        '''
        OU or User 重命名方法
        :param Old_dn: 需要修改的object的完整dn路径
        :param NewNmae: 新的名字，User格式："cn=新名字";ou格式："ou=新名字"
        :return:
        '''
        status = self.conn.modify_dn(Old_dn,NewNmae)
        if status:
            logging.info("%s重命名为%s成功" %(Old_dn,NewNmae))
        else:
            logging.error("%s重命名为%s失败" %(Old_dn,NewNmae))


    def Delete_Object(self, type=None, base_dn=None, **kwargs):
        '''
        删除用户、用户组、组织单位
        :param object_dn:
        :param uid:
        :return:
        '''

        is_exsit = self.Get_ObjectDN(type, **kwargs)
        if is_exsit:
            object_dn = is_exsit[0].entry_dn
            res = self.conn.delete(object_dn)

            if res:
                print('删除 %s 成功！' % object_dn)
            else:
                print('删除 %s 发生错误：' % object_dn, self.conn.result['description'])
        else:
            print("用户 %s 不存在"%kwargs)


    def Move_Object(self,type=None,new_object_dn=None, **kwargs):
        '''
        移动用户、用户组、组织单位到指定的组织单位
        :param type: 被移动的对象类型：user、org、group
        :param new_object_dn: 目标路径，如：OU=上海,OU=上海,OU=精锐.个性化,OU=精锐在线教育集团,DC=onesmart,DC=local
        :param kwargs: 被移动的对象（OU=上海一区,description=121）
        :return:
        '''
        is_exist = self.Get_ObjectDN(type,**kwargs)
        if is_exist:
            object_dn = is_exist[0].entry_dn
            if type == 'user' or type == 'group':
                #status = self.conn.modify_dn(object_dn,'CN=%s'%kwargs['sAMAccountName'], new_superior=new_object_dn)
                print(new_object_dn,kwargs)
                status = self.conn.modify_dn(object_dn,'CN=%s'%kwargs['sAMAccountName'], new_superior=new_object_dn)
                if status is True:
                    logging.info("移动：" + kwargs['sAMAccountName'] + "从" + object_dn + "到:" + new_object_dn + "成功")
                else:
                    logging.error("移动：" + kwargs['sAMAccountName'] + "从" + object_dn + "到:" + new_object_dn + "失败")
            else:
                status = self.conn.modify_dn(object_dn,'OU=%s'%kwargs['ou'], new_superior=new_object_dn)
                if status is True:
                    logging.info("移动：" + kwargs['ou'] + "从" + object_dn + "到:" + new_object_dn + "成功")
                else:
                    logging.error("移动：" + kwargs['ou'] + "从" + object_dn + "到:" + new_object_dn + "失败")
        else:
            logging.error(str(kwargs)+"不存在")

    def authenticate_user(self, upn, password):
        """
        验证用户名密码
        通过邮箱进行验证密码
        :param mail:
        :param password:
        :return:
        """

        entry = self.Get_ObjectDN(type='user',userPrincipalName=upn)
        if entry:
            bind_dn = entry[0].entry_dn
            try:
                Connection(self.server, bind_dn, password, auto_bind=True)
                print('验证通过')
                return True
            except LDAPBindError:
                print("验证失败")
                return False

        else:
            print("user: %s not exist! " % upn)
            return False

    def ChangePassword(self,org,uid):
        '''
            重置密码
        :param org:
        :param uid:
        :return:
        '''
        password = str(uid)[-4:] + "@Jr2020"
        status = self.conn.extend.microsoft.modify_password(org, password)
        if status:
            logging.info("%s密码重置为%s成功" %(org,password))
        else:
            logging.error("%s密码重置为%s失败" %(org,password))

    def AclExchangeReadField(self,jobnumber):
        '''
        设置exchange不可读字段
        :param jobnumber:
        :return:
        '''
        commands = r"C:\Users\andy\PycharmProjects\ADExchange2\scripts\AclExchange.ps1 %s" % jobnumber
        try:
            result = self.win.run_ps(commands)
            if result.status_code == 0:
                logging.info("执行命令：" + commands + " 成功!")
            else:
                logging.info("执行命令：" + commands + " 失败!")

        except Exception as e:
            # logging.error(e)
            print(e)

    def GetUserGroups(self,jobnumber):
        '''
            获取用户所有组
        :param jobnumber:
        :return: 列表
        '''
        ou_list = []
        try:
            result = self.win.run_ps("Get-ADPrincipalGroupMembership -Identity %s|Select-Object |ForEach-Object distinguishedName"%jobnumber).std_out.decode('utf-8').split('\n')
            for i in result:
                if 'Domain Users' not in i and i:
                    ou_list.append(i.strip())
        except Exception as e:
            print(e)
        return ou_list

    def GetGroupsUser(self,groupname):
        '''
        获取组中所有用户信息
        '''
        try:
            result = self.win.run_ps("Get-ADGroupMember %s|ft SamAccountName"%groupname)
            print(type(result))
            print(result)
        except Exception as e:
            print(e)



if __name__ == '__main__':
    import sys
    obj = Ad_Opertions()
    data = obj.Get_ObjectDN(type='org',base_dn="OU=至慧上海一区,OU=至慧上海,OU=至慧高端少儿思维,OU=至慧高端少儿思维,OU=精锐教育集团,DC=onesmart,DC=local",ou="至慧上海浦东唐镇中心*")
    # org_dn = data[0].entry_dn
    print(data)
    # obj.Update_dn(Old_dn=org_dn,NewNmae="ou=至慧上海浦东唐镇中心-失效")
    # # obj.Update_Org_Info(data[0].entry_dn, ou="test111",description='18157366080')

    # dep="精锐在线|精锐.班组课/精锐在线|精锐.班组课/精锐在线|精锐.班组课广东/精锐在线|精锐.班组课广东一区/精锐在线班组课深圳福田科技中学中心"
    # dep=dep.replace(".",'').replace('|','')
    # print(dep)
    # data=obj.Get_ObjectDN(type='user',cn='9300975')
    # print(data[0].entry_dn)
    # res=obj.Update_User_Info(data[0].entry_dn, department=dep)
    # print(res)
    # # obj.ChangePassword(data[0].entry_dn,'tesT005')
    # obj.Update_User_Info(data[0].entry_dn,extensionAttribute1="ftkenglish.com")
    # obj.Update_User_Info(data[0].entry_dn,mail="test005@ftkenglish.com")
    # obj.Update_User_Info(data[0].entry_dn,userPrincipalName="test005@ftkenglish.com")
    # obj.Update_User_Info(data[0].entry_dn,proxyAddresses="SMTP:test005@ftkenglish.com,smtp:test005@jingruigroup.partner.mail.onmschina.cn")
    sys.exit()

    #搜索用户
    # data = obj.Get_ObjectDN(type='org',ou='RTC',description=1234)
    # data = obj.Get_ObjectDN(type='org',base_dn='OU=精锐.个性化,DC=onesmart,DC=local',ou='财务组')
    # data = obj.Get_ObjectDN(type='org',base_dn='OU=精锐·个性化,DC=onesmart,DC=local',ou='财务组')
    # print(data)
    # print(len(data))
    # sys.exit()
    # result = obj.Add_Org(org_dn='OU=财务组,OU=南京总部2,OU=南京,OU=江苏,OU=精锐·个性化,DC=onesmart,DC=local',org_id='00000000')

    # print(result)

    # data = obj.Get_ObjectDN(type='user',userPrincipalName='1601016@onesmart.local')
    # print(data[0].entry_dn)
    # print(data)
    # # data=obj.Get_ObjectDN(type='org',description='121')
    # # print(data)
    # # print(len(data))
    #
    #
    # sys.exit()

    # #添加账号
    # ou_dn = "OU=技术支持&运维中心,OU=集团研发技术总部,OU=集团总部,OU=集团总部,OU=集团,OU=精锐教育集团,DC=onesmart,DC=local"
    # obj.Add_User(ou_dn, '杜俊朋', '1601016','18157366080','junpeng.du@onesmart.org','运维',domain='onesmart.org')

    # s = obj.authenticate_user("1601016@jronline.org.cn", "Andy.com..,")
    # print(s)

    # data = obj.Get_ObjectDN(type='user', base_dn='OU=精锐教育集团,DC=onesmart,DC=local',sAMAccountName='0210207')
    # print(len(data))

    # obj.Disable_User('CN=杨小坚,OU=小小地球广州万达中心,OU=小小地球广州一区,OU=小小地球广东,OU=小小地球广东,OU=小小地球,OU=精锐教育集团,DC=onesmart,DC=local','2400883')

    # new_object_dn = "OU=技术支持&运维中心,OU=集团研发技术总部,OU=集团总部,OU=集团总部,OU=集团,OU=精锐教育集团,DC=onesmart,DC=local"
    # obj.Move_Object(type='user', new_object_dn=new_object_dn, cn="杜俊朋",
    #                 mail="junpeng.du@jronline.org.cn")

    # data = obj.Get_ObjectDN(type='user',base_dn='OU=精锐教育集团,DC=onesmart,DC=local')
    # print(len(data))
    #
    # for i in data:
    #     print(i.entry_dn)
    #     # obj.Delete_Object(type='user',CN=i.entry_dn.split(',')[0].split('=')[1])

    # obj.Disable_User('CN=刘惠琳,OU=上海浦东高桥中心,OU=上海四区,OU=上海,OU=上海,OU=精锐.个性化,OU=精锐教育集团,DC=onesmart,DC=local','0219092')
    # obj.Move_Object(type='user', new_object_dn='OU=至慧无锡万象城中心,OU=至慧少儿数学无锡一区,OU=至慧少儿数学苏州,OU=至慧少儿数学,OU=至慧少儿数学,OU=精锐教育集团,DC=onesmart,DC=local', sAMAccountName="9903402")

    # org = "OU=资金中心,OU=集团总部,OU=集团总部,OU=集团总部,OU=集团,OU=精锐教育集团,DC=onesmart,DC=local"
    # obj.Add_User(org,'王官凌','2400497',mobile='13061660712',title='出纳',jobgrade=2,domain='onesmart.org',alias='guanling.wang')
    # # 把账号加入到通讯组
    # obj.AddUserInGroups_PS('2400497', 'all.list')
    # obj.AddUserInGroups_PS('2400497', '101')
    # obj.AddUserInGroups_PS('2400497', '5043')
    # obj.AclExchangeReadField('2400497')

    # data = obj.Get_ObjectDN(type='user',sAMAccountName='0211524')
    # if int(data[0].userAccountControl[0]) != 514:
    #     print(123)

    # obj.Update_User_Info(data[0].entry_dn,userPrincipalName='tingting.xu5@ftkenglish.com')
    # obj.Update_User_Info(data[0].entry_dn,extensionAttribute1='ftkenglish.com')
    # obj.Update_User_Info(data[0].entry_dn,mail='tingting.xu5@ftkenglish.com')
    # obj.Enable_User(data[0].entry_dn)
    # obj.Disable_User(data[0].entry_dn,'0220239')
    # time.sleep(1)
    # obj.authenticate_user('wei.xu2@onesmart.org','1566@Jr2020')

    #

    #
    # data  = obj.Get_ObjectDN(type='user',userPrincipalName='junpeng.du@onesmart.org')
    # print(data)
    # print(type(data[0].userPrincipalName[0]))
    # res = obj.GetUserGroups(1601016)
    # print(res)
    # data = obj.Get_ObjectDN(type='org', base_dn='OU=精锐教育集团,DC=onesmart,DC=local',
    #                  description=3983)
    # if not data:
    #     print(1)
    # else:
    #     print(2)
    data = obj.Get_ObjectDN(type='org', base_dn='OU=精锐教育集团,DC=onesmart,DC=local',description=102)
    print(data)

    sys.exit()

    #禁用账号
    obj.Disable_User('Andy1','1601016')

    #启用账号
    obj.Enable_User('Andy1','1601016')

    sys.exit()

    #更新用户信息
    data = obj.Get_ObjectDN(type='user',cn="1601016",mail="junpeng.du@onesmart.org",)
    #修改密码
    obj.Update_User_Info(data[0].entry_dn, userPassword="Andy.com..,")
    #修改职位
    obj.Update_User_Info(data[0].entry_dn, title="系统运维")


    #创建用户组
    obj.Add_User_Group('rainbower','CN=rainbower,DC=onesmart,DC=local',2)

    # #把用户加入到组
    obj.AddUserInGroups(['CN=1601016,OU=Andy1,DC=onesmart,DC=local'],['CN=Group001,DC=onesmart,DC=local'])



    #创建组织单位
    obj.Add_Org('Andy',00000000)
    obj.Add_Org('Andy1',11111111)
    obj.Add_Org('Andy2',22222222)
    obj.Add_Org('Andy3',33333333)
    obj.Add_Org('Andy4',44444444)
    obj.Add_Org('Andy5',55555555)


    #更新组织单位
    data = obj.Get_ObjectDN(type='org',ou="Andy",description=00000000)
    obj.Update_Org_Info(data[0].entry_dn,description=10000000)

    # 给指定的组织单位禁用"对象防止删除"属性
    obj.Org_Delete_Attribute(0,'OU=Andy1,DC=onesmart,DC=local','OU=Andy2,DC=onesmart,DC=local')
    # 给指定的组织单位启用"对象防止删除"属性
    obj.Org_Delete_Attribute(1,'OU=Andy1,DC=onesmart,DC=local','OU=Andy2,DC=onesmart,DC=local')

    # 给所有组织单位禁用"对象防止删除"属性
    obj.Org_Delete_Attribute(0)
    # 给所有的组织单位启用"对象防止删除"属性
    obj.Org_Delete_Attribute(1)



    #删除用户、用户组、组织单位
    obj.Delete_Object(type='user',cn="1601016",mail="junpeng.du@onesmart.org")
    obj.Delete_Object(type='group',cn='test123')
    obj.Delete_Object(type='org',ou='Andy3')


    # 移动用户、用户组、组织单位到指定组织单位下
    #移动用户
    obj.Move_Object(type='user',new_object_dn='OU=Andy1,dc=onesmart,dc=local',cn="1601016",mail="junpeng.du@onesmart.org")

    #移动组
    obj.Move_Object(type='group',new_object_dn='OU=Andy1,dc=onesmart,dc=local',cn='Group001')

    #移动组织单位
    obj.Move_Object(type='org',new_object_dn='OU=Andy,dc=onesmart,dc=local',ou="Andy1",description=11111111)
    obj.Move_Object(type='org',new_object_dn='OU=Andy,dc=onesmart,dc=local',ou="Andy2",description=22222222)
    obj.Move_Object(type='org',new_object_dn='OU=Andy,dc=onesmart,dc=local',ou="Andy3",description=33333333)
    obj.Move_Object(type='org',new_object_dn='OU=Andy,dc=onesmart,dc=local',ou="Andy4",description=44444444)
    obj.Move_Object(type='org',new_object_dn='OU=Andy,dc=onesmart,dc=local',ou="Andy5",description=55555555)
