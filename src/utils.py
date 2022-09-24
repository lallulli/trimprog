# coding: utf-8
# Roma e20
# Written by Luca Allulli

from contextlib import contextmanager
import os

def skipping_iter(iterable, n=1):
    """
    Turns an iterable into another iterable, that skips first line of original one

    :param iterable: An iterable
    :param n: Number of items to skip
    """
    i = iterable.__iter__()
    for j in range(n):
        i.__next__()
    try:
        while True:
            yield i.__next__()
    except StopIteration:
        pass


@contextmanager
def _others(iterable):
    try:
        while True:
            yield iterable.next()
    except StopIteration:
        pass


def first_others(iterable):
    """
    Return (first, others) where first is the first item, others is a generator
    """
    i = iterable.__iter__()
    first = i.next()
    return first, _others(i)


def create_dir_if_not_existing(path, recurse=False):
    path = os.path.abspath(path)
    if not os.path.exists(path):
        parent, sub = os.path.split(path)
        if recurse and not os.path.exists(parent):
            create_dir_if_not_existing(parent, True)
        os.mkdir(path)


@contextmanager
def chdir(new_dir):
    """
    Context manager. Chdir to a temporary dir, and return to previous working dir on exit.
    """
    old = os.getcwd()
    try:
        os.chdir(new_dir)
        yield
    finally:
        os.chdir(old)
