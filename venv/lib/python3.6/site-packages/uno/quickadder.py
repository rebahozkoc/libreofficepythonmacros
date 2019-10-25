# -*- coding: utf-8 -*-

import copy
import uno.css_constants as c




class CssQuickAdder(object):

    def __init__(self, parent_obj):
        self.element = parent_obj

    def _add(self, thing):
        self.element._add_css(thing)

    def add(self, thing):
        self._add(thing)

    def append_css(self, name, info):
        self.element.__dict__[name] += info

    def ng_model(self, info):
        self._add({c.NGM: info})

    def name(self, info):
        self._add({c.NAME: info})

    def _type(self, info):
        self._add({c.TYPE: info})

    def TYPE(self, info):
        self._type(info)

    def _class(self, info):
        self._add({c.CLASS: info})

    def klass(self, info):
        self._class(info)

    def CLASS(self, info):
        self._class(info)

    def class_(self, info):
        self._class(info)

    def equal_to(self, info):
        self._add({c.EQ : info})

    def for_(self, info):
        self.FOR(info)

    def _for(self, info):
        self.FOR(info)

    def FOR(self, info):
        self._add({c.FOR: info})

    def required(self, info):
        self._add({c.REQUIRED, info})






class GroupQuickAdder(object):

    def element(self, name, feature, info):
        for ele in self.group.elements:
            if ele.name == name:
                action = getattr(ele.feature, feature)
                action(info)

    def verifier(self, ele_name):
        current_index = 0
        for ele_obj in self.group.elements:
            if ele.name == ele_name:
                ver = copy.copy(ele_obj)
                ver.add_css({c.NAME : self.ele_obj.css_dict[c.NAME] + c._CONFIRM })
                ver.add_css({c.EQ : ele_obj.css_dict[c.NGM]})
                ele_obj.add_css({c.EQ : ver.css_dict[c.NAME]})
                self.group.insert_element(current_index+1, ver)
            current_index += 1


    def __init__(self, parent_obj):
        self.group = parent_obj



