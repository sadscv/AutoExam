#!/usr/bin/env python  
# -*- coding:utf-8 _*-
""" 
@author: sadscv
@time: 2019/03/17 16:05
@file: timeslot.py

@desc: 
"""
import random

from src.optimization.domain.exam import Exam


class Timeslot(object):
    def __init__(self, timeslot_id):
        self.num_collisions = 0
        self.exams = []
        self._id = timeslot_id

    def get_exams(self):
        return self.exams

    @property
    def id(self):
        return self._id

    @id.setter
    def id(self, timeslot_id):
        self._id = timeslot_id

    @property
    def collisions(self) -> int:
        return self.num_collisions

    @collisions.setter
    def collisions(self, n: int):
        self.num_collisions = n

    def get_exam(self, i: int) -> Exam:
        return self.exams[i]

    def get_random_exam(self):
        if len(self.exams) == 0:
            return None
        else:
            rd = random.randint(0, len(self.exams) - 1)
            return self.exams[rd]

    def add_exam(self, e: Exam):
        if e is not None:
            self.exams.append(e)

    def remove_exam(self, e: Exam):
        # Todo: check if following remove would cause bug
        # i.e. many identical Exam occur at one timeslot'exams.
        self.exams.remove(e)

    def force_exam(self, pending_exam: Exam):
        """
        force an exam to e scheduled in this timeslot.
        :param pending_exam: the exam we want to force in this timeslot
        :return: None
        """
        collision_generated = 0
        for e in self.exams:
            collision_generated += pending_exam.stu_involved_in_collision(e.id)
        self.exams.add(pending_exam)
        self.num_collisions += collision_generated

    def is_compatible(self, e: Exam) -> bool:
        """
        Tell if the timeslot is compatible with Exam e.
        i.e. if it doesn't contain any of e's conflicting exam, return True
        :param e:
        :return:
        """
        if self.is_free():
            return True
        compatible = True
        for already_in in self.exams:
            # a &= b equals to a = a && b
            compatible &= e.is_compatible(already_in.id)
            if not compatible:
                return False
        return True

    def is_free(self):
        # Todo:test this func
        """
        Tell if there are exams currently scheduled in this timeslot
        :return: True if this timeslot currently free
        """
        return True if self.exams else False

    def conflict_weight(self, conflict_weights: dict) -> int:
        """
        Count the weight of the conflict between an exam with a certain set
        of conflicting exams (represented by the conflictWeight) and this
        timeslot
        # Todo: try to figure out what's this func built for
        :param conflict_weights: a dict [Exam.id, number of commonStudent]
        :return:
        """
        count = 0
        for e in self.exams:
            if conflict_weights[e.id] is not None:
                count += conflict_weights[e.id]
        return count

    def num_exams(self):
        return len(self.exams)

    def __eq__(self, other: Timeslot) -> bool:
        if self == other:
            return True
        if other is None:
            return False

        if type(other) == Timeslot:
            print('warning, here may cause error')
            return other.exams == self.exams
        return False

    def __hash__(self):
        """
        Because of  exam appears in only one timeslot, we cn use the first
        exam at an timeslot to identify timeslot univocally.
        :return:
        """
        if not self.exams:
            return 0
        return self.exams[0].id + 1

    def __copy__(self):
        """
        override copy method
        :return:
        """
        clone = type(self)(self.id)
        for e in self.exams:
            clone.add_exam(e.__copy__())
        clone.collisions = self.num_collisions
        return clone

    def __deepcopy__(self, memo):
        cls = self.__class__
        result = cls.__new__(cls)
        memo[id(self)] = result
        for k, v in self.__dict__.items():
            setattr(result, k, __import__("copy").deepcopy(v, memo))
        return result

    def contains(self, e: [int, Exam]) -> bool:
        """
        Returns T/F if the exam or e_id <e> is assigned to this timeslot or not
        :param e:
        :return:
        """
        if isinstance(e, Exam):
            return self.exams.__contains__(e)
        for exam in self.exams:
            if exam.id == e:
                return True
        return False

    def clean(self):
        """
        Remove all the exams assigned to the timeslot
        :return: None
        """
        self.exams = []

    def add_exams(self, new_exams: list):
        """
        Append extra exams to current timeslot's exam_list
        :param new_exams:
        :return: None
        """
        tmp_set = set(self.exams)
        tmp_set.add(set(new_exams))
        return list(tmp_set)

    def try_insert_exams(self, to_insert: list[Exam]) -> list[Exam]:
        """
        For each exam in to_insert, if the timeslot doesn't contain it,
        and timeslot is compatible with it. we assign this exam to current 
        timeslot, and also add it to the list of positioned exam.
        :param to_insert:
        :return:
        """
        positioned = []
        for e in to_insert:
            if (not self.contains(e)) and self.is_compatible(e):
                self.add_exam(e)
                positioned.append(e)
        return positioned

    def get_exam_by_id(self, e_id: int) -> Exam:
        for e in self.exams:
            if e.id == e_id:
                return e
        return None

    def __str__(self):
        string = ""
        for e in self.exams:
            string += e.__str__ + "\t"
        return string
