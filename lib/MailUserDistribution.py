# -*- coding:utf-8 -*-
__author__ = 'Rainbower'

import pymysql
from config.settings import Params

import logging
# 日志设置
LOG_FORMAT = "%(asctime)s  %(levelname)s  %(filename)s  %(lineno)d  %(message)s"
logging.basicConfig(filename='../logs/debug.log', level=logging.INFO, format=LOG_FORMAT)

class MailUserDistribution():
    def __init__(self):
        self.port = Params['MysqlConfig']['port']
        self.host = Params['MysqlConfig']['host']
        self.user = Params['MysqlConfig']['user']
        self.password = Params['MysqlConfig']['password']
        self.conn = pymysql.connect(self.host,self.user,self.password,'jr_ldap_exchange',self.port)
        self.cursor = self.conn.cursor()
    def insert_update_maildb(self,**kwargs):
        if type(kwargs) is dict:
            #maildb jobnumber username gender mobile email jobgrade title organization is_active  createtime
            select_sql = 'select jobnumber from mail_user_distribution where jobnumber="%s"' % kwargs['jobnumber']
            insert_sql = 'insert into mail_user_distribution(maildb,jobnumber,username,gender,mobile,email,jobgrade,title,organization,is_active,createtime) values("%s","%s","%s","%s","%s","%s","%s","%s","%s","%s","%s")' % (
                kwargs['maildb'], kwargs['jobnumber'], kwargs['username'], kwargs['gender'], kwargs['mobile'],
                kwargs['email'], kwargs['jobgrade'], kwargs['title'], kwargs['organization'], kwargs['is_active'],
                kwargs['createtime'])

            update_sql = "update mail_user_distribution_counter set counter=%s where maildb='%s'" % (
                kwargs['counter'], kwargs['maildb'])

            update_user = 'update mail_user_distribution set mobile="%s",email="%s",jobgrade="%s",title="%s",organization="%s",is_active="%s",createtime="%s" where jobnumber="%s"' % (
                kwargs['mobile'], kwargs['email'], kwargs['jobgrade'], kwargs['title'], kwargs['organization'],
                kwargs['is_active'], kwargs['createtime'], kwargs['jobnumber'])

            # print(update_user)
            try:
                self.conn.ping(reconnect=True)
                #cursor = self.conn.cursor()
                is_exist = self.cursor.execute(select_sql)
                if is_exist == 0:
                    self.cursor.execute(insert_sql)
                    self.cursor.execute(update_sql)

                    logging.info(kwargs['jobnumber']+"入库成功")
                else:
                    # print(kwargs['jobnumber']+"已存在、更新信息")
                    self.cursor.execute(update_user)

                self.conn.commit()
                self.conn.close()

            except Exception as e:
                logging.info(kwargs['jobnumber'] + "入库失败" + e)
                self.conn.rollback()


    def select_maildb(self):
        '''
            查找counter值最小的数据
        :return: dbname,counter
        '''

        sql = "select maildb,counter from mail_user_distribution_counter where counter=(select min(counter) from mail_user_distribution_counter) limit 1"
        try:
            self.conn.ping(reconnect=True)
            self.cursor.execute(sql)
            result = self.cursor.fetchone()
        except:
            self.conn.rollback()

        return result


    def select_email(self,name):
        '''
            查找邮箱地址是否存在
        :param name:
        :return:
        '''
        namelist = []
        sql = 'select email from mail_user_distribution where email like "%s%s";' %(name,'%')
        try:
            self.conn.ping(reconnect=True)
            self.cursor.execute(sql)
            results = self.cursor.fetchall()
            for i in results:
                namelist.append(i[0].split('@')[0])
        except:
            self.conn.rollback()

        return namelist

if __name__ == '__main__':
    a = MailUserDistribution()
    a.select_email()