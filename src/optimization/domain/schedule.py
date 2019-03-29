#!/usr/bin/env python  
# -*- coding:utf-8 _*-
""" 
@author: sadscv
@time: 2019/03/23 20:23
@file: schedule.py

@desc: 
"""
import random

from src.optimization.domain.exam import Exam
from src.optimization.domain.timeslot import Timeslot


class Schedule(object):
    def __init__(self, tmax, num_students):
        self.MAX_SWAP_TRIES = 10
        self._cost = 0
        self.tmax = tmax
        self._timeslots = self.init_timeslots()
        self.num_students = num_students

    def init_timeslots(self) -> list[Timeslot]:
        ts = []
        for i in range(self.tmax):
            ts[i] = Timeslot(i)
        return ts

    @property
    def timeslots(self):
        return self._timeslots

    @timeslots.setter
    def timeslots(self, ts: Timeslot):
        self._timeslots = ts

    def tmax(self):
        return len(self._timeslots)

    def get_num_students(self):
        return self.num_students

    @property
    def cost(self):
        return self._cost

    @cost.setter
    def cost(self, cost):
        self._cost = cost

    def update_cost(self, penalty):
        self._cost += penalty

    def get_total_collisions(self):
        tot = 0
        for t in self._timeslots:
            tot += t.collisions
        return tot

    def get_timeslot(self, i: int) -> Timeslot:
        return self._timeslots[i]

    def get_random_timeslot(self, bucket: [None, int, list[int]] = None) -> \
            Timeslot:
        """
        Return a random Timeslot up to a certain bound which bucket provided,
        if bucket not given. it's up bound is self._timeslot
        :return:
        """
        if not bucket:
            return self._timeslots[random.randint(len(self._timeslots))]
        if isinstance(bucket, int):
            return self._timeslots[random.randint(bucket)]
        else:
            return self._timeslots[random.randint(len(bucket))]

    def get_timeslot_with_exams(self):
        """
        Return a timeslot which must contain exams.
        :return:
        """
        t = Timeslot()
        while t.is_free():
            t = self.get_random_timeslot()
        return t

    def num_exams(self):
        """
        Return the number of exams in the timeslots of this schedule
        :return:
        """
        counter = 0
        for t in self.timeslots:
            counter += len(t.get_exams())
        return counter

    def clean(self):
        self.init_timeslots()

    @staticmethod
    def move(ex: Exam, source: Timeslot, dest: Timeslot) -> bool:
        """
        Todo write CostFunction before handling this code
        :param ex:
        :param source:
        :param dest:
        :return:
        """
        if dest.is_compatible(ex):
            pass

    def check_feasible_swap(self, t1: Timeslot, ex1: Exam,
                            t2: Timeslot,  ex2: Exam):
        return self.check_swap(t1, ex1, ex2) and self.check_swap(t2, ex2, ex1)

    @staticmethod
    def check_swap(t1: Timeslot, ex1: Exam, ex2: Exam) -> bool:
        t1.remove_exam(ex1)
        compatible = t1.is_compatible(ex2)
        t1.add_exam(ex1)
        return compatible

    def swap(self, t1: Timeslot, ex1: Exam, t2: Timeslot, ex2: Exam):
        """
        Swap two exam's from each other's timeslot
        :param t1:
        :param ex1:
        :param t2:
        :param ex2:
        :return:
        """
        self.move(ex1, t1, t2)
        self.move(ex2, t2, t1)

    def random_placement(self, tobe_placed: Exam, bound: [int, list[int]]=None):
        """
        Tries to place an exam into a random Timeslot. considering a window
        of timeslots of width
        :param tobe_placed: the exam that has to be placed
        :param bound: the width of the window
        :return:
        """
        if not bound:
            destination = self.get_random_timeslot()
        else:
            destination = self.get_random_timeslot(bound)
        if destination.is_compatible(tobe_placed):
            destination.add_exam(tobe_placed)
            return True


