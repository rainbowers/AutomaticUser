# -*- coding:utf-8 -*-
__author__ = 'Rainbower'

from api.Adapi import Ad_Opertions
from lib.ChineseToPinyin import ChineseToPinyin
from lib.MailUserDistribution import MailUserDistribution

def ListToList(slist):
    '''
        从数字列表中找出不连续的数字，并重新组成列表
    :param slist:
    :return:
    '''
    lists = []
    slist.sort()
    b = (x for x in range(slist[0], slist[-1] + 1))
    for i in b:
        if  i not in slist:
            lists.append(i)
            continue
    if lists:
        return lists
    else:
        return [slist[-1]+1]

def StringToInt(string):
    '''
        把字符窜中包含数字的部分提取出来，如：ying.zhang2,ying.zhang10
    :param string:
    :return:
    '''
    lists =[]
    for i in string:
        if i>="0" and i<="9":
            lists.append(i)
    if lists:
        return int(''.join(lists))

def NameConvert(name):
    '''
        把中文名字转换为拼音("名.姓")、检查是否已存在，并重组；例如：meixuan.li2
    :param name:
    :return:
    '''
    ad_obj = Ad_Opertions()
    Alias_name = ChineseToPinyin(name)

    result = ad_obj.Get_ObjectDN(type='user', userPrincipalName='%s*' % Alias_name)
    if result:
        a = []
        Name_list = []
        for name in result:
            Name_list.append(name.userPrincipalName[0].split('@')[0])

        for i in Name_list:
            res = StringToInt(i)
            if res:
                if res >= 2:
                    a.append(1)
                    a.append(StringToInt(i))
                else:
                    a.append(StringToInt(i))
            else:
                a.append(1)
        data = ListToList(a)
        #print(Alias_name+str(data[0]))
        return Alias_name+str(data[0])
    else:
        return Alias_name
if __name__ == "__main__":
    res = NameConvert('张紫微')
    print(res)