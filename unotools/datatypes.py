# -*- coding: utf-8 -*-


def Sequence(*args):
    if len(args) == 1 and args[0] is None:
        return tuple()
    return tuple(args)
