


from uno.base import  Element


class BaseInput(object):
    """docstring for BaseInput"""
    def __init__(self, name, **kwargs):
        super(BaseInput, self).__init__(name, 'input', **kwargs)
        self._is_type = ('element', 'input', 'base')
        self._name = name
        self._value = kwargs.get('value', '')
        self._type  = kwargs.get('_type', 'text')
        self._type  = kwargs.get('TYPE', self._type)

        self.TYPE = Css('type', self._type)
        self.name = Css('name', name)
        self.value = Css('value', self._value)


class TextInput(BaseInput):
    pass


#ng-modeled text input 

class AngularTextInput(BaseInput):
    def __init__(self, name, **kwargs):
        super(AngularTextInput, self).__init__(name, 'input', **kwargs)
        self.ng_model = Css('ng_model', 'ng_'+ name)

