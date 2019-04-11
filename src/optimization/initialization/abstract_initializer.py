#!/usr/bin/env python  
#-*- coding:utf-8 _*-  
""" 
@author: sadscv
@time: 2019/04/05 21:41
@file: abstract_initializer.py

@desc: 
"""
from threading import Thread
from typing import List

from src.optimization.domain.exam import Exam
from src.optimization.domain.schedule import Schedule


class AbstractInitializer(Thread):
    def __init__(self, exams: List[Exam], optimizer: Optimizer,
                 num_students: int, tmax: int):
        super().__init__()
        self.exams = exams
        self.optimizer = optimizer
        self.schedule = Schedule(tmax, num_students)
        self.num_students = num_students

    def run(self):
        self.initialize()
        self.schedule.compute_cost()
        self.notify_new_initial_solution()

    def initialize(self):
        """
        Computes a new schedule
        :return:
        """
        pass

    def notify_new_initial_solution(self):
        self.optimizer.update_on_new_initial_solution(self.schedule)

    def get_percent_placed(self):
        """
        Return the percentage of exams currently placed in the schedule.
        :return:
        """
        return self.schedule.num_exams() / len(self.exams)

    def __str__(self):
        return self.schedule.num_exams() + "out of" + len(self.exams) + "placed"

