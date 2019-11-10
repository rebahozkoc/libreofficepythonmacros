#coding:utf8
import re
class Base():
    def fenge(self,content):
        return [content] + re.split(r'[。\n]\s*', content)