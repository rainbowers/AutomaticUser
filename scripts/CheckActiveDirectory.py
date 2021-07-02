# -*- coding:utf-8 -*-
__author__ = 'Rainbower'
import os
# class CheckActiveDirectoryStatus():
def CheckLinkStatus():
    '''
        目录复制链路状态
        0：服务正常
        1：服务失败
    :return:
    '''
    result = os.popen("repadmin /showrepl").read()
    if  result.count("尝试成功") < 5:
        print(0)
    else:
        print(1)

def CheckDCStatus():
    '''
        AD健康检测
        0：服务正常
        1：服务失败
    :return:
    '''
    result = os.popen("dcdiag /q").read()
    print(result)


# CheckLinkStatus()
CheckDCStatus()