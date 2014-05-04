# -*- coding: utf-8 -*-
import uno
import unohelper
from com.sun.star.bridge import XUnoUrlResolver
from com.sun.star.frame import XDesktop
from com.sun.star.lang import XComponent
from com.sun.star.lang import XMultiComponentFactory
from com.sun.star.reflection import XIdlReflection
from com.sun.star.script.provider import XScriptContext
from com.sun.star.uno import XInterface

from unotools.utils import cached_property, get_annotation_to_kwargs
from unotools.utils import set_kwargs


class LocalContext(unohelper.Base):

    def __init__(self):
        self.context = uno.getComponentContext()

    @cached_property
    def service_manager(self) -> XMultiComponentFactory:
        return self.context.getServiceManager()

    @cached_property
    def desktop(self) -> XDesktop:
        service_name = 'com.sun.star.frame.Desktop'
        return self.create_instance_with_context(service_name)

    @cached_property
    def resolver(self) -> XUnoUrlResolver:
        service_name = 'com.sun.star.bridge.UnoUrlResolver'
        return self.create_instance_with_context(service_name)

    @cached_property
    def document(self) -> XComponent:
        return self.desktop.getCurrentComponent()

    @cached_property
    def core_reflection(self) -> XIdlReflection:
        return self.create_instance('com.sun.star.reflection.CoreReflection')

    def load_component_from_url(self, url: str, target_frame_name: str,
                                search_flags: int, arguments: tuple=()
                                ) -> XComponent:
        """
        url:
            'private:factory/swriter', 'private:factory/scalc', etc ...

        target_frame_name:
            '_blank': always creates a new frame
            '_default': special UI functionality
                (e.g. detecting of already loaded documents, using of
                      empty frames of creating of new top frames as fallback)
            '_self', ''(!):  means frame himself
            '_parent': address direct parent of frame
            '_top': indicates top frame of current path in tree
            '_beamer': means special sub frame

        search_flag:
            0: Auto, 1: Parent, 2: Self, 4: Children, 8: Create,
            16: Siblings, 32: Tasks, 23: All, 55: Global
        """
        return self.desktop.loadComponentFromURL(url, target_frame_name,
                                                 search_flags, arguments)

    def create_instance(self, name: str) -> XInterface:
        return self.service_manager.createInstance(name)

    def create_instance_with_context(self, name: str) -> XInterface:
        args = (name, self.context)
        return self.service_manager.createInstanceWithContext(*args)

    def create_struct(self, type_name: str) -> XInterface:
        rv, struct = self.core_reflection.forName(type_name).createObject(None)
        return struct

    def make_struct_data(self, type_name: str, **kwargs) -> XInterface:
        struct = self.create_struct(type_name)
        set_kwargs(struct, kwargs)
        return struct

    def make_point(self, x: int, y: int) -> XInterface:
        kwargs = self._get_kwargs('make_point', locals())
        return self.make_struct_data('com.sun.star.awt.Point', **kwargs)

    def make_property_value(self, name: str=None, value: str=None,
                            handle: int=None, state: object=None
                            ) -> XInterface:
        type_name = 'com.sun.star.beans.PropertyValue'
        kwargs = self._get_kwargs('make_property_value', locals())
        return self.make_struct_data(type_name, **kwargs)

    def make_rectangle(self, x: int, y: int, width: int, height: int
                       ) -> XInterface:
        kwargs = self._get_kwargs('make_rectangle', locals())
        return self.make_struct_data('com.sun.star.awt.Rectangle', **kwargs)

    def make_size(self, width: int, height: int) -> XInterface:
        kwargs = self._get_kwargs('make_size', locals())
        return self.make_struct_data('com.sun.star.awt.Size', **kwargs)

    def _get_kwargs(self, func_name: str, values: dict) -> dict:
        return get_annotation_to_kwargs(self.__class__, func_name, values)


class ScriptContext(LocalContext, XScriptContext):

    def __init__(self, context):
        self.context = context
