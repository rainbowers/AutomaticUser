# -*- coding:utf-8 -*-
__author__ = 'Rainbower'

import winrm
import logging
from config.settings import Params

# 日志设置
LOG_FORMAT = "%(asctime)s  %(levelname)s  %(filename)s  %(lineno)d  %(message)s"
logging.basicConfig(filename='../logs/logs.log', level=logging.INFO, format=LOG_FORMAT)


class SyncAAD():
    def __init__(self):
        self.ip = Params['Aadconfig']['host']
        self.user = Params['Aadconfig']['user']
        self.pwd = Params['Aadconfig']['password']
        self.win = winrm.Session('http://' + self.ip + ':5985/wsman', auth=(self.user, self.pwd),transport='ntlm')

    def SyncUserinfoToOnline(self):
        '''
            同步用户信息到云端
        :return:
        '''

        status = self.win.run_ps("Start-ADSyncSyncCycle -PolicyType Delta")
        if status.status_code == 0:
            logging.info("用户信息同步成功")
        else:
            logging.info("用户信息同步失败")


if __name__ ==  '__main__':
    sync_obj = SyncAAD()
    sync_obj.SyncUserinfoToOnline()

