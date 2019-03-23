#!/usr/bin/env python  
# -*- coding:utf-8 _*-
""" 
@author: sadscv
@time: 2019/03/15 21:12
@file: exam.py

@desc: 
"""
from typing import Any, Tuple


class Exam(object):
    def __init__(self, exam_id, num_students):
        self.conflicting_exams = {}
        self._id = exam_id
        self.num_students = num_students

    def get_num_students(self):
        return self.num_students

    @property
    def id(self):
        return self._id

    @id.setter
    def id(self, exam_id: int):
        if exam_id >= 0:
            self._id = exam_id

    def add_conflicting_exams(self, eid: int, num_students=None) -> None:
        """
        Add a new exam to this exam collisions set
        if there wasn't any previous collision between this exam and eid,
        then create a new row inside the collision set and set the number of
        students involved in the collision to one. Otherwise, increase the
        old value by one.

        :param num_students: number of students conflicted while adding that
        exam
        :param eid: the id of the exam we want to add to the conflicting
        exams set
        """
        if num_students:
            self.conflicting_exams[eid] = num_students
        else:
            num_stu_involved = 1
            if eid in self.conflicting_exams:
                num_stu_involved += self.conflicting_exams[eid]
            self.conflicting_exams[eid] = num_stu_involved

    def get_conflict_exam(self) -> dict:
        return self.conflicting_exams

    def stu_involved_in_collision(self, exam_id: int) -> int:
        """
        Get the number of students that are enrolled in both this exam and
        'exam_id' exam.

        :param exam_id: The exam we want to check the collision for.
        :return: The number of students involved in the collision. If no
        collision exist, return 0;
        """
        if exam_id in self.conflicting_exams:
            return self.conflicting_exams[exam_id]
        return 0

    def is_compatible(self, exam_id: int) -> bool:
        """
        Tells if this exam is compatible(i.e. not conflicting) with Exam(
        exam_id)
        :param exam_id:
        :return: T/F if the two exams are conflicting or not
        """
        return exam_id not in self.conflicting_exams

    def conflict_exam_size(self):
        return len(self.conflicting_exams)

    def compare_to(self, other: Exam) -> int:
        return self.conflict_exam_size() - other.conflict_exam_size()

    def __str__(self):
        return "Ex" + self._id + " "

    def __eq__(self, other: object) -> bool:
        if type(other) == Exam:
            return other.get_id() == self.id
        return False

    def __hash__(self):
        """
        override __hash__ func to make sure when two Exam instance with
        same id comparing (like exam1 == exam2)would return true.
        :return:
        """
        hashcode = 5
        hashcode = 31 * hashcode + hash(self._id)
        return hashcode

    def __copy__(self):
        # cls = self.__class__
        # result = cls.__new__(cls)
        # result.__dict__.update(self.__dict__)
        clone = Exam(self.id, self.num_students)
        for key, value in self.conflicting_exams.items():
            clone.add_conflicting_exams(key)
            clone.add_conflicting_exams(key, value)
        return clone



