# -*- coding: utf-8 -*-
import uno
import unohelper
from com.sun.star.script.provider import XScriptContext

from unotools.utils import cached_property, get_annotation_to_kwargs
from unotools.utils import set_kwargs


class LocalContext(unohelper.Base):

    def __init__(self):
        self.context = uno.getComponentContext()

    @cached_property
    def service_manager(self):
        return self.context.getServiceManager()

    @cached_property
    def desktop(self):
        service_name = 'com.sun.star.frame.Desktop'
        return self.create_instance_with_context(service_name)

    @cached_property
    def resolver(self):
        service_name = 'com.sun.star.bridge.UnoUrlResolver'
        return self.create_instance_with_context(service_name)

    @cached_property
    def document(self):
        return self.desktop.getCurrentComponent()

    @cached_property
    def core_reflection(self):
        return self.create_instance('com.sun.star.reflection.CoreReflection')

    def load_component_from_url(url, properties=()):
        return self.desktop.loadComponentFromURL(url, '_blank', 0, properties)

    def create_instance(self, name: str):
        return self.service_manager.createInstance(name)

    def create_instance_with_context(self, name: str):
        args = (name, self.context)
        return self.service_manager.createInstanceWithContext(*args)

    def create_struct(self, type_name: str):
        rv, struct = self.core_reflection.forName(type_name).createObject(None)
        return struct

    def make_struct_data(self, type_name: str, **kwargs):
        struct = self.create_struct(type_name)
        set_kwargs(struct, kwargs)
        return struct

    def make_point(x: int, y: int):
        kwargs = self._get_kwargs('make_point', locals())
        return self.make_struct_data('com.sun.star.awt.Point', **kwargs)

    def make_property_value(self, name: str=None, value: str=None,
                            handle: int=None, state: object=None):
        type_name = 'com.sun.star.beans.PropertyValue'
        kwargs = self._get_kwargs('make_property_value', locals())
        return self.make_struct_data(type_name, **kwargs)

    def make_rectangle(self, x: int, y: int, width: int, height: int):
        kwargs = self._get_kwargs('make_rectangle', locals())
        return self.make_struct_data('com.sun.star.awt.Rectangle', **kwargs)

    def make_size(self, width: int, height: int):
        kwargs = self._get_kwargs('make_size', locals())
        return self.make_struct_data('com.sun.star.awt.Size', **kwargs)

    def _get_kwargs(self, func_name, values):
        return get_annotation_to_kwargs(self.__class__, func_name, values)


class ScriptContext(LocalContext, XScriptContext):

    def __init__(self, context):
        self.context = context
