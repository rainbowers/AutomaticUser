# -*- coding:utf-8 -*-
__author__ = 'Rainbower'
import sys,os
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)

from lib import core
exec_obj = core.HandleArgvs()

help_msg = '''
    allorg      添加及更新组织单位
    alluser     添加全量用户
    updateuser  更新增量用户
    useringroup 更新所有用户到对应的组内
    addsaccount 创建特殊账号
    
'''
if len(sys.argv) >1:
    if sys.argv[1] == 'allorg':
        exec_obj.operations_ou_full()
    if sys.argv[1] == 'alluser':
        exec_obj.operations_user(5)
    if sys.argv[1] == 'createuser':
        exec_obj.CreateMailUser(sys.argv[2],sys.argv[3])
    if sys.argv[1] == 'updatedepartment':
        exec_obj.UpdateDepartment()
    if sys.argv[1] == 'oudiff':
        exec_obj.operations_ou_diff()
    if sys.argv['']:
        exec_obj.create_special_account()
else:
    print(help_msg)