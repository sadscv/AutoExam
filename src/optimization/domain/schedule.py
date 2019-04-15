#!/usr/bin/env python  
# -*- coding:utf-8 _*-
""" 
@author: sadscv
@time: 2019/03/23 20:23
@file: schedule.py

@desc: 
"""
import random
from typing import List

from src.optimization.domain.cost_function import CostFunction
from src.optimization.domain.exam import Exam
from src.optimization.domain.timeslot import Timeslot


class Schedule(object):
    _timeslots: List[Timeslot]

    def __init__(self, tmax, num_students):
        self.MAX_SWAP_TRIES = 10
        self._cost = 0
        self._tmax = tmax
        self._timeslots = self.init_timeslots()
        self.cost_func = CostFunction()
        self.num_students = num_students

    def init_timeslots(self) -> List[Timeslot]:
        ts = []
        for i in range(self._tmax):
            ts.append(Timeslot(i))
        return ts

    @property
    def timeslots(self):
        return self._timeslots

    @timeslots.setter
    def timeslots(self, ts: List[Timeslot]):
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

    def get_random_timeslot(self, bucket: [int, List[int]] = None) -> \
            Timeslot:
        """
        Return a random Timeslot up to a certain bound which bucket provided,
        if bucket not given. it's up bound is self._timeslot. if bucket is
        list[int], return a random timeslot in the bucket
        :return:
        """
        if not bucket:
            return self._timeslots[random.randint(len(self._timeslots))]
        if isinstance(bucket, int):
            return self._timeslots[random.randint(bucket)]
        else:
            return self._timeslots[bucket[random.randint(len(bucket))]]

    def get_timeslot_include_exams(self):
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

    def move(self, ex: Exam, source: Timeslot, dest: Timeslot) -> bool:
        """
        :param ex:
        :param source:
        :param dest:
        :return:
        """
        if dest.is_compatible(ex):
            penalty = self.cost_func. \
                get_exam_move_penalty(ex, source.id, dest.id, self.timeslots)
            self.update_cost(penalty)
            source.remove_exam(ex)
            dest.add_exam(ex)
            return True
        return False

    def check_feasible_swap(self, t1: Timeslot, ex1: Exam, t2: Timeslot, ex2: Exam):
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

    def random_placement(self, tobe_placed: Exam, bound: [int, List[int]] = None):
        """
        Tries to place an exam into a random Timeslot. considering a window
        of timeslots of `bound`
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
        return False

    def force_random_placement(self, tobe_placed: Exam):
        """
        Force an exam to be placed into a random schedule. even if it
        generates a collision.
        :param tobe_placed: Exam we want to schedule.
        :return:
        """
        self.get_random_timeslot().force_exam(tobe_placed)

    def random_move(self, bound: [int, List[int]] = None) -> bool:
        """
        Try to move an exam from one timeslot to another timeslot randomly
        :return: T/F
        """
        tj = self.get_timeslot_include_exams()
        tk = self.get_random_timeslot(bound)

        ex = tj.get_random_exam()
        return self.move(ex, tj, tk)

    def random_swap(self):
        """
        Try to swap two exam randomly
        :return:
        """
        tj = self.get_timeslot_include_exams()
        tk = self.get_timeslot_include_exams()
        ex1 = tj.get_random_exam()
        ex2 = tk.get_random_exam()
        if self.check_feasible_swap(tj, ex1, tk, ex2):
            self.swap(tj, ex1, tk, ex2)
            return True
        return False

    def free_placement(self, tobe_placed: Exam) -> bool:
        """
        Place `tobe_placed` Exam into the first empty timeslot we found.
        :param tobe_placed:
        :return:
        """
        for destination in self.timeslots:
            if destination.is_free():
                destination.add_exam(tobe_placed)
                return True
        return False

    def free_timeslot_available(self):
        """
        Count the amount of all timeslots which is empty
        :return:
        """
        count = 0
        for destination in self.timeslots:
            if destination.is_free():
                count += 1
        return count

    def is_feasible_move(self, exam: Exam, dest_index: int) -> bool:
        """
        Check if the Exam `exam` canbe move to the 'dest_index' timeslot
        :param exam:
        :param dest_index:
        :return:
        """
        return self.timeslots[dest_index].is_compatible(exam)

    def swap_timeslots(self, i: int, j: int):
        """
        :param i:
        :param j:
        :return:
        """
        penalty = self.cost_func.get_timeslot_swap_penalty(i, j,
                                                           self.timeslots)
        ti = self.get_timeslot(i)
        tj = self.get_timeslot(j)
        ti.id = j
        tj.id = i
        self.timeslots[i] = tj
        self.timeslots[j] = ti
        self.update_cost(penalty)

    @staticmethod
    def get_timeslot_pair_hash(t1: object, t2: object, reverse: bool = False):
        if reverse:
            return t2.__hash__() + "-" + t1.__hash__()
        return t1.__hash__() + "-" + t2.__hash__()

    def __copy__(self):
        s = Schedule(self._tmax, self.num_students)
        s.cost = self.cost
        tclone = []
        for i in range(self._tmax):
            tclone.append(self.timeslots[i].__copy__())
        s.timeslots = tclone

    def __str__(self):
        s = ''
        for i in range(len(self.timeslots)):
            s += "Timeslot{}{}\n".format(i, self.timeslots[i].__str__())
        return s

    def __cmp__(self, other: 'Schedule') -> int:
        return self._cost - other._cost

    def __eq__(self, other):
        if not issubclass(other, Schedule):
            return False
        return self.__cmp__(other) == 0

    def __hash__(self):
        hashcode = 5
        hashcode = 29 * hashcode + self.cost
        return hashcode

    def mutate_exams(self):
        """
        Tries to
        :return:
        """
        swapped = False
        iter_num = 0
        while (not swapped) and (iter_num is not self.MAX_SWAP_TRIES):
            swapped = self.random_swap_with_penalty()
            iter_num += 1
        return swapped

    def random_swap_with_penalty(self):
        """
        Tries to swap two exam in two different timeslot when the swap could
        reduce penalty.
        :return:
        """
        ti = self.get_timeslot_include_exams()
        tj = self.get_timeslot_include_exams()
        ex1 = tj.get_random_exam()
        ex2 = tj.get_random_exam()

        if self.check_feasible_swap(ti, ex1, tj, ex2):
            penalty = self.get_swap_cost(ti.id, ex1, tj.id, ex2)
            if penalty < 0:
                self.swap(ti, ex1, tj, ex2)
                return True
        return False

    def get_swap_cost(self, tid1, ex1, tid2, ex2):
        """
        :param tid1:
        :param ex1:
        :param tid2:
        :param ex2:
        :return:
        """
        return (self.cost_func.get_exam_move_penalty(ex1, tid1, tid2,
                                                     self.timeslots) + \
                self.cost_func.get_exam_move_penalty(ex2, tid2, tid1,
                                                     self.timeslots))

    def try_mutate_timeslots(self):
        """
        Try to select 2 random timeslots, if the cost function decrease,
        swap the order of two timeslots
        :return:
        """
        tk = self.get_random_timeslot()
        tj = self.get_random_timeslot()
        penalty = self.cost_func.get_timeslot_swap_penalty(tk.id, tj.id,
                                                           self.timeslots)
        if penalty < 0:
            self.mutate_timeslot(tj, tk)
            self.update_cost(penalty)

    def mutate_timeslot(self, tj, tk):
        """
        Swap the postion of two timeslots
        :param tj:
        :param tk:
        :return:
        """
        tj_id = tj.id
        tk_id = tk.id
        self.timeslots[tj_id].id = tk_id
        self.timeslots[tk_id].id = tj_id
        self.timeslots[tj_id] = tk
        self.timeslots[tk_id] = tj

    def try_invert_timeslots(self, cut_points):
        penalty = self.cost_func.get_invert_timeslots_penalty(
            cut_points[0], cut_points[1], self.timeslots)
        if penalty < 0:
            self.invert_timeslots(cut_points[0], cut_points[1])
            self.update_cost(penalty)

    def invert_timeslots(self, start, end):
        """
        Assign timeslots's exams between `start` and `end` to it's
        reverse order timeslot.
        :param start:
        :param end:
        :return:
        """
        length = end - start
        tmp_slots = []
        for i in range(length):
            tmp_slots.append(Timeslot(i))
        for i in range(length):
            tmp_slots[length - 1 - i].add_exams(self.timeslots[start +
                                                               i].get_exams())
            self.timeslots[start + i].clean()
        for i in range(length):
            self.timeslots[start + i].add_exams(tmp_slots[i].get_exams())

    def try_crossover(self, parent2exams: List[Exam]):
        """
        Choose a random exam among `parent2exams`. it's the cost function
        decreased, do the crossover.we call it `穿越`

        :param parent2exams:
        :return:
        """
        if parent2exams is None:
            return False
        candidate = parent2exams[random.randint(len(parent2exams))]
        for t in self.timeslots:
            if t.contains(candidate):
                penalty = self.calc_crossover_penalty(parent2exams, t)
                if penalty < 0:
                    self.do_crossover(parent2exams, t)
                    self.update_cost(penalty)
                    return True
        return False

    def calc_crossover_penalty(self, parent2exams: List[Exam], dest: Timeslot)\
            -> int:
        """
        For each exam in `parent2exams`. if the exam is not in timeslot
        `dest` and compatible to move to dest, calculate the penalty related
        to the movements.
        :param parent2exams:
        :param dest:
        :return:
        """
        penalty = 0
        for e in parent2exams:
            src = self.get_timeslot_by_exams(e)
            if src != dest and dest.is_compatible(e):
                # Todo , test self.timeslots will cause bugs.
                penalty += self.cost_func. \
                    get_exam_move_penalty(e, src.id, dest.id, self.timeslots)
        return penalty

    def do_crossover(self, parent2exams, t):
        """
        Tries to insert the exams of `parent2exams` into timeslot `t`. if at
        least one exams has been moved, it fix the schedule, removing the
        exam from it's previous position
        :param parent2exams:
        :param t:
        :return:
        """
        positioned_exams = t.try_insert_exams(parent2exams)
        # Todo test if positioned_exams is empty
        if positioned_exams:
            self.fix_schedule(t, positioned_exams)

    def fix_schedule(self, dest_timeslot, positioned_exams):
        """
        here the fix means repair, fix the mismatch of timeslots and exams.
        :param dest_timeslot:
        :param positioned_exams:
        :return:
        """
        for e in positioned_exams:
            for t in self.timeslots:
                if t.contains(e) and t is not dest_timeslot:
                    t.remove_exam(e)

    def get_timeslot_by_exams(self, e: Exam):
        """
        Retrieve the timeslots which contains exam `e`
        :param e:
        :return:
        """
        for t in self.timeslots:
            if t.contains(e):
                return t
        return None

    def compute_cost(self, cost_map: dict = None):
        if cost_map:
            # Todo, complete this fun with cost_map parameter
            self._cost = self.cost_func.get_cost(self.timeslots, cost_map)
            # pass
        self._cost = self.cost_func.get_cost(self.timeslots)
