#!/usr/bin/env python
# -*- coding:utf-8 _*-
""" 
@author: sadscv
@time: 2019/03/31 17:03
@file: cost_function.py

@desc: 
"""
from typing import List

from src.optimization.domain.exam import Exam
from src.optimization.domain.schedule import Schedule
from src.optimization.domain.timeslot import Timeslot


class CostFunction(object):
    COST = [0, 16, 8, 4, 2, 1]

    def get_cost(self, timeslots: [Schedule, List[Timeslot]]) -> int:
        """
        Return the cost of a certain timeslot configuration.
        N.B. it still has to be weighted by the total number of students.
        [(t1,t2)(t2,t3)(t3,t4)(t4,t5)(t5,t6)(t6,t7)]*COST[1]
        [(t1,t3)(t2,t4)(t3,t5)(t4,t6)(t5,t7)]*COST[2]
        [(t1,t4)(t2,t5)(t3,t6)(t4,t7)]*COST[3]
        [(t1,t5)(t2,t6)(t3,t7)]*COST[4]
        [(t1,t6)(t2,t7)]*COST[5]
        :param timeslots:
        :return:
        """
        if isinstance(timeslots, Schedule):
            timeslots = timeslots.timeslots
        penalty = 0
        for i in range(1, 6):
            for t in range(len(timeslots) - i):
                # calculate each `e`'s conflict weights on this `t`
                for e in timeslots[t].get_exams():
                    penalty += timeslots[t + i].conflict_weight(
                        e.get_conflict_exam()) * self.COST[i]
        return penalty

    def get_exam_cost(self, e: Exam, src: int,
                      timeslots: List[Timeslot]) -> int:
        """
        Return the cost associated to a specific exam, The closer conflict Exam
        's distance to `e`, the higher penalty gives it.
        :param e: The exam for which we want to compute the associated cost
        :param src: The index of the timeslot associated to the selected exam
        :param timeslots: The starting set of timeslots.
        :return: The cost associated to the given exam.
        """
        range_i = self.get_range(src, len(timeslots))
        penalty = 0
        for i in range(range_i[0], range_i[1] + 1):
            if i != src:
                distance = src - i
                penalty += timeslots[i].conflict_weight(e.get_conflict_exam(
                )) * self.COST[abs(distance)]
        return penalty

    def get_timeslot_cost(self, index: int, timeslots: List[Timeslot]) -> int:
        """
        Return the cost associated to a specific timeslot

        :param index: The index associated to the selected timeslot
        :param timeslots: The starting set of timeslots
        :return:
        """
        range_t = self.get_range(index, len(timeslots))
        cost = 0
        for i in range(range_t[0], range_t[1] + 1):
            distance = index - i
            for e in timeslots[i].get_exams():
                cost += timeslots[i].conflict_weight(e.get_conflict_exam()) * \
                        self.COST[abs(distance)]
        return cost

    def get_exam_move_penalty(self, e: Exam, src: int, dest: int,
                              timeslots: [Schedule, List[Timeslot]]) -> int:
        """
        Returns the cost while swapping an exam from `src` to `dest` timeslot
        :param e: The exam we want to move
        :param src: The index of current timeslot of given `e`
        :param dest: The index of destination time we want to move
        :param timeslots: The starting set of timeslots
        :return: The penalty assigned to a specific move.
        """
        if isinstance(timeslots, Schedule):
            timeslots = timeslots.timeslots
        penalty = 0
        penalty -= self.get_exam_cost(e, src, timeslots)
        penalty += self.get_exam_cost(e, dest, timeslots)
        return penalty

    def get_timeslot_swap_penalty(self, i: int, j: int,
                                  timeslots: [List[Timeslot], Schedule]) -> int:
        """
        Return the cost while swapping timeslot with index `i` and index `j`
        :param i:
        :param j:
        :param timeslots:
        :return:
        """
        if isinstance(timeslots, Schedule):
            timeslots = timeslots.timeslots
        penalty = 0
        penalty += self.get_new_timeslot_penalty(i, j, timeslots)
        penalty += self.get_new_timeslot_penalty(j, i, timeslots)
        return penalty

    def get_new_timeslot_penalty(self, current_index: int, other_index: int,
                                 timeslots: List[Timeslot]) -> int:
        """
        Calculate the penalty given by virtually moving the timeslot from
        `other_index` to `current_index`.It subtracts the penalty due to the
        position of the timeslot that was in current_index
        :param current_index:
        :param other_index:
        :param timeslots:
        :return:
        """
        range_i = self.get_range(current_index, len(timeslots))
        penalty = 0
        for i in range(range_i[0], range_i[1]):
            if i != current_index and i != other_index:
                distance = current_index - i
                for e in timeslots[current_index].get_exams():
                    penalty -= timeslots[i].conflict_weight(
                        e.get_conflict_exam()) * self.COST[abs(distance)]
                for e in timeslots[other_index].get_exams():
                    penalty += timeslots[i].conflict_weight(
                        e.get_conflict_exam()) * self.COST[abs(distance)]
        return penalty

    def get_invert_timeslots_penalty(self, start: int, end: int,
                                     timeslots: [List[Timeslot],
                                                 Schedule]) -> int:
        """
        Return the cost of inverting the order of timeslots between `start`
        and `end` in the array of `timeslots`
        :param start:
        :param end:
        :param timeslots:
        :return:
        """
        penalty = 0
        if isinstance(timeslots, Schedule):
            timeslots = timeslots.timeslots
        for i in range(start, end):
            index = timeslots[i].id
            invert_index = (end - 1) - (index - start)
            if index != invert_index:
                penalty -= self.get_timeslot_cost(index, timeslots)
                penalty += self.calc_reposition_penalty(start, end,
                                                        invert_index,
                                                        index, timeslots)
        return penalty

    def calc_reposition_penalty(self, start: int, end: int, invert_index: int,
                                index: int, timeslots: List[Timeslot]) -> int:
        """
        Todo to read the func tomorrow
        :param start:
        :param end:
        :param invert_index:
        :param index:
        :param timeslots:
        :return:
        """
        penalty = 0
        range_i = self.get_range(invert_index, len(timeslots))
        for i in range(range_i[0], range_i[1]):
            if i != invert_index:
                distance = invert_index - i
                t = timeslots[i]

                if start <= i < end:
                    t = timeslots[end - 1 - (i - start)]
                for e in timeslots[index].get_exams():
                    penalty += t.conflict_weight(e.get_conflict_exam()) * \
                               self.COST[abs(distance)]
        return penalty

    def get_mutual_timeslots_cost(self, t1: Timeslot, t2: Timeslot) -> int:
        timeslots = [t1, t2]
        return self.get_cost(timeslots)

    @staticmethod
    def get_range(index: int, tmax: int) -> List[int]:
        """
        Return the range of indexes that are 5 timeslots before and after
        `index`. It takes care of the timeslots bound
        :param index:
        :param tmax:
        :return:
        """
        bound = [max(index - 5, 0), min(index + 5, tmax - 1)]
        return bound
