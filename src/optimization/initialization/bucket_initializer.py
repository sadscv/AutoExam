#!/usr/bin/env python
# -*- coding:utf-8 _*-
""" 
@author: sadscv
@time: 2019/04/06 18:00
@file: bucket_initializer.py

@desc: 
"""
import random
from typing import List, Any

from src.optimization.domain.exam import Exam
from src.optimization.initialization.abstract_initializer import \
    AbstractInitializer


class BucketInitializer(AbstractInitializer):
    already_placed: List[Exam]
    not_yet_placed: List[Exam]
    MOVE_TRIES: int
    PLACING_TRIES: int
    buckets: List[int]
    STUCK = 2
    THRESHOLD = 1
    current_width: int
    ORDER = [0, 3, 1, 4, 2, 5]

    def __init__(self, tmax: int, num_students: int, exams: List[Exam],
                 optimizer: Optmizer):
        super().__init__(exams, optimizer, num_students, tmax)
        self.tmax = tmax
        self.organize_exams()
        self.not_yet_placed = exams
        self.already_placed = []
        self.MOVE_TRIES = 2 * tmax
        self.PLACING_TRIES = 2 * tmax
        self.buckets = []
        self.initialize_buckets()

    def initialize_buckets(self):
        """
        Here we assign first `tmax` exams into 6 buckets.
        :return:
        """
        for i in range(6):
            current_bucket = []
            current = self.ORDER[i]
            while current < self.tmax:
                current_bucket.append(current)
                current += 6
            self.buckets[i] = current_bucket

    def initialize(self):
        """
        1. Try to place as many exams as possible randomly considering a
        restricted window of timeslots.
        2. When get stuck with an exam(can't placed it within `PLACING_TRIES` tries.
        try to move the already placed exams into other timeslots.
        3. As long as the number of tires reached `THRESHOLD`, widen the
        window a step.
        4. Stop until all exams placed. If get stuck. restart the process.

        :return:
        """
        num_tries = im_stuck = bucket_i = 0
        current_bucket: List[int] = self.buckets[bucket_i]
        bucket_i += 1
        while self.not_yet_placed:
            # try to process not_yet_placed[] as much as possible.
            self.compute_random_schedule(current_bucket)
            # when we can't move an exam from an random timeslot to an timeslot
            # within current_bucket, we stop the random initialize.
            while not self.random_move(current_bucket):
                pass
            num_tries += 1
            if num_tries > self.THRESHOLD:
                if bucket_i < 6:
                    current_bucket.append(self.buckets[bucket_i])
                    self.current_width = len(current_bucket)
                    bucket_i += 1
                else:
                    im_stuck += 1
                    if im_stuck == self.STUCK:
                        im_stuck = 0
                        bucket_i = 0
                        self.restart()
                        bucket_i += 1
                        current_bucket = self.buckets[bucket_i]
                num_tries = 0
        return self.schedule

    def organize_exams(self):
        """
        Sort exams based on their number of conflicts, but it adds some
        noise. instead of having a perfectly ordered list of exams,
        we shuffle groups of 5 exams.
        :return:
        """
        sorted(self.exams, key=lambda exam: exam.conflict_exam_size)
        n_exams = len(self.exams)
        randomized: List[Exam] = []
        new_indexes = self.get_random_sequence(n_exams, shuffle_size=5)
        for i in range(n_exams):
            randomized[new_indexes[i]] = self.exams[i]
        self.exams = randomized

    @classmethod
    def get_random_sequence(cls, n_exams, shuffle_size):
        """
        Returns an list of number from 0 to `n_exams`(exclusive),
        in a "slightly random" order.
        i.e. groups of size `shuffle_size` are shuffled.
        :param n_exams:
        :param shuffle_size:
        :return:
        """
        seq = [i for i in range(n_exams)]
        i = 0
        while shuffle_size * (i + 1) < n_exams:
            i += 1
            offset = i * shuffle_size
            for j in range(shuffle_size/2 + 1):
                m = random.randint(shuffle_size) + offset
                n = random.randint(shuffle_size) + offset
                seq[m], seq[n] = seq[n], seq[m]
        return seq

    def compute_random_schedule(self, current_bucket):
        """
        Tries to generate a schedule by means of random computation.
        :param current_bucket: The bucket of timeslots which `not_yet_placed[0]`
                               trying to insert into.
        :return:
        """
        i = 0
        while self.not_yet_placed and i < len(self.not_yet_placed):
            to_be_placed = self.not_yet_placed[i]
            # cuz not_yet_placed[0] always been the most involved exam,
            # so we dont need to ++i in the while loop when we
            # successfully place previous `to_be_placed` exam.
            if not self.try_random_placement(to_be_placed, current_bucket):
                i += 1
                # use break to interrupt the while loop, in case of most of the
                # exams are assigned into one color all at once.
                break

    def random_move(self, current_bucket):
        """
        Tries to move 1 exam from a random timeslot to `current_bucket`
        :param current_bucket:
        :return:
        """
        for i in range(self.MOVE_TRIES):
            if self.schedule.random_move(current_bucket):
                return True
        return False

    def restart(self):
        pass

    def try_random_placement(self, to_be_placed, current_bucket):
        """
        Tries to place an exam into a random timeslot within given number of
        tries.
        :param to_be_placed:
        :param current_bucket:
        :return:
        """
        for i in range(self.PLACING_TRIES):
            if self.schedule.random_placement(to_be_placed, current_bucket):
                self.not_yet_placed.remove(to_be_placed)
                self.already_placed.append(to_be_placed)
                return True
        return False

