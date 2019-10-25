# -*- coding: utf-8 -*-

from collections import Iterable, OrderedDict

from uno import (helpers, constants)

from uno.constants import (PAYLOAD, CSS, STATIC_TAGS, NORMAL_TAGS, 
                           ABNORMAL_TAGS, SELF_CLOSING_TAGS,
                           RESERVED_WORDS_UPPER)

PAYLOAD_TAGS = helpers.minus(NORMAL_TAGS, ABNORMAL_TAGS)

from uno.quickadder import CssQuickAdder, GroupQuickAdder

import copy

from itertools import chain


class UnoBase(object):

    def __repr__(self):
        return self._render

    def __str__(self):
        return self._render

    def __call__(self):
        return self #._render

    def __unicode__(self):
        return unicode(self._render)

    def __html__(self):
        return self._render

    def __add__(self, other):
        try:
            y = other._render
        except:
            y = str(other)
        return self._render + y

    def __setattr__(self, name, value):
        if not name.startswith('_'):
            if isinstance(value, UnoBase):
                self._features[name] = value
        object.__setattr__(self, name, value)

class UnoBaseFeature(UnoBase):
    members = []

    def _auto_add_features(self):
        for key in self.__class__.__dict__.keys():
            if not key.startswith('_'):
                obj = self.__class__.__dict__[key]
                self._features[key] = obj

    def _render_features(self):
        
        text = ''
        for key in self._features:
            value = self._features[key]
            text += value._render
        return text

    @property
    def _render(self):
        self._render = self._text + self._render_features()
        return self.__render

    @_render.setter
    def _render(self, value):
        self.__render = value

    def __init__(self, *args, **kwargs):
        self._features = OrderedDict()
        self._auto_add_features()
        #self._features      = []
        #self._features_dict = {}
        self._is_type       = ('feature', 'base')
        self._text          = ''
        self._payload       = ''
        self.__class__.members.append(self)
        self._uno_id = self.__class__.members.index(self)


    def _register(self, obj):
        self._parent = obj

    def _add_feature(self, feature):
        self._features[feature._name] = feature

    def _all_features(self):
        x = self._features
        for feat in self._features:
            x = helpers.combine_dicts(x, self._features[feat]._all_features())
        return x


class Payload(UnoBaseFeature):

    def __init__(self, name, text, **kwargs):
        super(Payload, self).__init__(self, **kwargs)
        self._is_type    = ('payload',)
        self._name       = name
        self._text      = text

    @property
    def _render(self):
        x = self._text + self._render_features()
        self._render = x
        return self.__render

    @_render.setter
    def _render(self, value):
        self.__render = value



class Css(UnoBaseFeature):

    @property
    def _value(self):
        return self.__value

    @_value.setter
    def _value(self, value):
        print 'new css', self._attr, 'value:', value
        self.__value = value
    
    @property
    def _attr(self):
        return self.__attr
    @_attr.setter
    def _attr(self, value):
        print 'new css attr:', value
        self.__attr = value

    def __init__(self, attr, value, **kwargs):
        super(Css, self).__init__(self, **kwargs)
        self._attr       = attr 
        self._value      = value
        self._is_type   = ('css',)
        self._text      = ' ' + self._reservered_word_check(attr)\
                        + '="{}"'.format(value)
        

    def _reservered_word_check(self, word):
        if word in RESERVED_WORDS_UPPER:
            return word.lower()
        else:
            return word

    def _extend(self, value):
        self.value += ' ' + value

    @property
    def _render(self):
        self._render = self._text + self._render_features()
        return self.__render
    @_render.setter
    def _render(self, value):
        self.__render = value





class Element(UnoBaseFeature):

    def __init__(self, name, tag, *args, **kwargs):
        self._quick = CssQuickAdder(self)
        self._tag = tag
        self._postcss_tag = '>'
        self._closing_tag = '</' + self._tag + '>'
        self._precss_tag = '<' + self._tag
        self._static_tag_check(tag)
        self._self_closing_check(tag)
        super(Element, self).__init__(self, name, **kwargs)
        self._name = name
        self._is_type = ('element',)

    def _static_tag_check(self, tag):
        for stat in STATIC_TAGS:
            if tag == stat[0]:
                self._precss_tag = stat[1]
                self._closing_tag = ''
                self._postcss_tag = ''
                
    def _add_css(self, css_dict):
        print 'css_dict:', css_dict
        for key in css_dict:
            attr_key = key.replace('-', '_')
            setattr(self, attr_key, Css(key, css_dict[key]) )




    def _self_closing_check(self, tag):
        if tag in SELF_CLOSING_TAGS:
            self._closing_tag = ''
            self._postcss_tag = '/>'

    def _render_group(self, is_type):
        text = ''
        first = True
        for name in self._features:
            obj = self._features[name]
            if is_type in obj._is_type:
                if first and is_type in ['payload','element']:
                    first = False
                    text += ' '
                text += obj._render
        return text

    def _render_css(self):
        is_type = 'css'
        text    = ''
        first   = True
        for name in self._features:
            obj = self._features[name]
            if is_type in obj._is_type:
                text += obj._render
        return text

    def _render_payload(self):
        x = self._render_group('payload')
        if x != '':
            x += '\n'
        return x 

    def _render_elements(self):
        x = self._render_group('element')
        return x


    @property
    def _render(self):
        r =  self._precss_tag
        #print 'precss: ', r
        e = self._render_css()
        #print 'css: ', e
        n = self._postcss_tag
        #print 'postcss: ', n
        d = self._render_payload()
        #print 'payload: ', d
        er = self._render_elements()
        #print 'elements: ', er
        pls = self._closing_tag
        #print 'closing tag: ', pls
        self.__render = r+e+n+d+er+pls + '\n'
        return self.__render.replace('\n', '').replace('\\', '')

class UnoBaseField(UnoBaseFeature):

    def __init__(self, name, **kwargs):
        super(UnoBaseField, self).__init__(self, **kwargs)
        self._is_type = ('field', 'base')
        self._data = ''
        self._name = ''
        self._values = []
        self._title = self._name.replace('_', ' ').title()
        self._form_name = kwargs.get('parent_name', '')


    def _add_value(self, value):
        self._values.append(value)

    def _add_values(self, values):
        self._values += values

    def _new_values(self, values):
        self._values = values

    def _add_name(self, name):
        self._name = name


class UnoBaseForm(UnoBaseFeature):

    def __init__(self, *args, **kwargs):
        super(UnoBaseForm, self).__init__(self, *args, **kwargs)
        self._is_type = ('form', 'base')
        self._fields = []

    @property
    def _fields(self):
        self._update_fields()
        return self.__fields
    @_fields.setter
    def _fields(self, value):
        self.__fields = value
    

    def _update_fields(self):
        for attr in self._features.keys():
            if 'field' in self._features[attr]._is_type:
                self._fields.append(self._features[attr])

    def _populate_obj(self, obj):
        for field in self._fields:
            print field._name
