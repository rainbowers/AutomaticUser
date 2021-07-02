# -*- coding:utf-8 -*-

import datetime


def get_timespamp(date):
    return datetime.datetime.strptime(date,"%Y-%m-%dT%H:%M:%S").timestamp()

def hours(stime):
    current_time = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    cdate=datetime.datetime.strptime(current_time,"%Y-%m-%dT%H:%M:%S")
    sdata=datetime.datetime.strptime(stime,"%Y-%m-%dT%H:%M:%S")
    num=(cdate-sdata).total_seconds()/3600
    return float(num)