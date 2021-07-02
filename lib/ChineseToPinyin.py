# -*- coding:utf-8 -*-

from pypinyin import lazy_pinyin

def is_Chinese(word):
    for ch in word:
        if '\u4e00' <= ch <= '\u9fff':
            return True
    return False

def ChineseToPinyin(str):
    '''
    中文拼音、支持多音字、支持简单的繁体
    :param str:
    :return:
    '''
    is_true = is_Chinese(str)
    if is_true:
        data = lazy_pinyin(str,errors='ignore')
        surname = data[0]
        name = ''.join(data[1:])
        mail_name = '%s.%s' %(name,surname)
    else:
        mail_name = str
    return mail_name

if __name__ == '__main__':
    res = ChineseToPinyin('沪东外国语学校')

    print(res)