# -*- coding: utf-8 -*-
import argparse
import logging

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


def parse_argument(argv):
    """
    >>> argv = '-s server -p 8080 -o tcpNoDelay=1'.split()
    >>> parse_argument(argv)  # doctest: +NORMALIZE_WHITESPACE
    Namespace(host='server', option='tcpNoDelay=1', pipe=None, port=8080,
              verbose=False)
    """
    parser = argparse.ArgumentParser()
    parser.set_defaults(host='localhost', option=None, pipe=None, port=8100,
                        verbose=False)
    parser.add_argument('-i', '--pipe', dest='pipe',
                        metavar='PIPE', help='set pipe name')
    parser.add_argument('-o', '--option', dest='option',
                        metavar='OPTION', help='set option')
    parser.add_argument('-p', '--port', dest='port', type=int,
                        metavar='PORT_NUMBER', help='set port number')
    parser.add_argument('-s', '--host', dest='host',
                        metavar='HOST', help='set host name')
    parser.add_argument('-v', '--verbose', action='store_true',
                        help='set verbose mode')
    args = parser.parse_args(argv)

    if args.verbose:
        logging.root.setLevel(logging.DEBUG)
    logging.debug('args: {0}'.format(args))
    return args
