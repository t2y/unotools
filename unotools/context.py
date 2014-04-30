# -*- coding: utf-8 -*-
import uno
import unohelper
from com.sun.star.script.provider import XScriptContext


class LocalContext(unohelper.Base):

    def __init__(self):
        self.context = uno.getComponentContext()

    @property
    def resolver(self):
        return self.context.getServiceManager().createInstanceWithContext(
                'com.sun.star.bridge.UnoUrlResolver', self.context)


class ScriptContext(unohelper.Base, XScriptContext):

    def __init__(self, context):
        self.context = context

    @property
    def desktop(self):
        return self.context.getServiceManager().createInstanceWithContext(
                'com.sun.star.frame.Desktop', self.context)

    @property
    def document(self):
        return self.desktop.getCurrentComponent()
