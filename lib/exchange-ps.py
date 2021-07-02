# -*- coding:utf-8 -*-
__author__ = 'Rainbower'

import winrm
import logging
from config import settings

# 日志设置
LOG_FORMAT = "%(asctime)s  %(levelname)s  %(filename)s  %(lineno)d  %(message)s"
logging.basicConfig(filename='../logs/log.log', level=logging.INFO, format=LOG_FORMAT)

class Exchange():
    def __init__(self):
        self.host = settings.Params['ExchangeConfig']['host']
        self.user = settings.Params['ExchangeConfig']['user']
        self.password = settings.Params['ExchangeConfig']['password']
        self.win = winrm.Session('http://' + self.host + ':5985/wsman', auth=(self.user, self.password))
        # self.win = winrm.Session('http://' + self.host + ':5985/wsman', auth=(self.user, self.password),transport='ntlm')

    def enable_user(self, action, username, aliasname):
        enable_command = r"%s\%s.ps1 %s %s" % (settings.Params['script_dir'], action, username, aliasname)
        print(enable_command)
        status = self.win.run_ps(enable_command)
        if status.status_code == 0:
            logging.info(username + "邮箱开通成功")
        else:
            logging.error(username + "邮箱开通失败")

    def disble_user(self, action, username):
        disable_command = r"%s\%s.ps1 %s" % (settings.Params['script_dir'], action, username,)
        print(disable_command)
        status = self.win.run_ps(disable_command)
        if status.status_code == 0:
            logging.info(username + "邮箱禁用成功")
        else:
            logging.error(username + "邮箱禁用失败")
