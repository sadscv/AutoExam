#!/usr/bin/env python  
# -*- coding:utf-8 _*-
""" 
@author: sadscv
@time: 2019/03/29 19:14
@file: test_schedule.py

@desc: 
"""
from src.optimization.domain.schedule import Schedule
from src.optimization.domain.timeslot import Timeslot


def test_get_timeslot():
    ts = [Timeslot() for i in range(3)]
    sc = Schedule(tmax=3, num_students=1)
    # sc._timeslots = ts
    print(sc.get_timeslot(0))


if __name__ == '__main__':
    test_get_timeslot()
