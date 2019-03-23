#!/usr/bin/env python  
# -*- coding:utf-8 _*-
""" 
@author: sadscv
@time: 2019/03/17 15:27
@file: student.py

@desc:
"""
from .exam import Exam


class Student(object):
    def __init__(self, stu_id):
        self._id = stu_id
        self.exams = []
        self.exams_ids = []

    @property
    def id(self):
        return self._id

    @id.setter
    def id(self, stu_id):
        if not isinstance(stu_id, int):
            assert ValueError('stu_id must be an integer')
        self._id = stu_id

    def add_exam(self, e: Exam) -> None:
        self.exams.append(e)

    def exam_id(self, eid):
        self.exams_ids.append(eid)

    def get_exams_ids(self) -> list:
        return self.exams_ids

    def check_exam(self, eid: Exam) -> bool:
        return eid in self.exams

    def check_exam_id(self, eid: int) -> bool:
        return eid in self.exams_ids

    def get_exams(self):
        return self.exams

    def __str__(self):
        return self.id()


