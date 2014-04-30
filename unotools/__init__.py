# -*- coding: utf-8 -*-

try:
    from functools import singledispatch
except ImportError:  # under 3.3
    from singledispatch import singledispatch

import uno
import unohelper
from com.sun.star.script.provider import XScriptContext

from unotools.context import LocalContext, ScriptContext
from unotools.errors import ConnectionError


class Pipe:
    def __init__(self, name):
        self.name = name


class Socket:
    def __init__(self, host, port):
        self.host = host
        self.port = port


@singledispatch
def connect(identifier, **kwargs):
    context = None
    try:
        local_context = LocalContext()
        option = kwargs.get('option')
        if option is not None:
            identifier = '{},{}'.format(identifier, option)
        conn_str = 'uno:{};urp;StarOffice.ComponentContext'.format(identifier)
        context = local_context.resolver.resolve(conn_str)
        if context:
            return ScriptContext(context)
    except Exception as e:
        msg = 'failed to connect: {}'.format((identifier, kwargs))
        raise ConnectionError(msg) from e

    raise ConnectionError('cannot connect: {}'.format((identifier, kwargs)))


@connect.register(Pipe)
def connect_with_pipe(pipe, **kwargs):
    return connect('pipe,name={}'.format(pipe.name), **kwargs)


@connect.register(Socket)
def connect_with_socket(socket, **kwargs):
    identifier = 'socket,host={},port={}'.format(socket.host, socket.port)
    return connect(identifier, **kwargs)
