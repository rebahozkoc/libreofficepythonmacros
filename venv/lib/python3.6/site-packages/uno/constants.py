# -*- coding: utf-8 -*-

PAYLOAD = '{PAYLOAD}'
CSS     = '{CSS}'

#Html tag classification attributed to htmldog.com. It's pretty nice. 

FORM_TAGS = ('input', 'form', 'textarea', 'select', 'option', 'optgroup', 
            'button', 'label', 'fieldset', 'legend',)

STRUCTURE_TAGS = ('html', 'head', 'body', 'div', 'span',)

META_TAGS = ('DOCTYPE', 'title', 'link', 'meta', 'style',)

TEXT_TAGS = ('p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'strong', 'em', 
            'abbr', 'acronym', 'address', 'bdo', 'blockquote', 'cite', 
            'q', 'code', 'ins', 'del', 'dfn', 'kbd', 'pre', 'samp', 'var',
            'br',)

LINK_TAGS = ('a', 'base',)

IMAGE_AND_OBJECTS_TAGS = ('img', 'area', 'map', 'object', 'param',)

LIST_TAGS = ('ul','ol','li','dl','dt','dd',)

TABLE_TAGS = ('table', 'tr', 'td', 'th', 'tbody', 'thead', 'tfoot', 'col', 'colgroup', 'caption',)

SCRIPTING_TAGS = ('script', 'noscript',)

PRESENTATION_TAGS = ('b', 'i', 'tt', 'sub', 'sup', 'big', 'small', 'hr',)

SPECIAL_PAYLOAD_TAGS = [('comment', '<!-- {PAYLOAD} -->')]

STATIC_TAGS = [('DOCTYPE', '<!DOCTYPE>'),('doctype', '<!DOCTYPE>')]

SELF_CLOSING_TAGS = ('area', 'base', 'br', 'col', 'command', 'embed', 'hr',
        'img', 'input', 'keygen', 'link', 'meta', 'param', 'source', 
        'track', 'wbr')




NORMAL_TAGS = ['div', 'nav', 'html', 'head',
                'label', 'form', 'button', 'textarea', 'select', 'option', 
                'optgroup', 'ul', 'ol', 'li', 'a', ] +  list(TEXT_TAGS)


ALL_TAGS = list(FORM_TAGS) + list(STRUCTURE_TAGS) + list(META_TAGS) + list(TEXT_TAGS) + list(LINK_TAGS) + list(IMAGE_AND_OBJECTS_TAGS) + list(LIST_TAGS) + list(TABLE_TAGS) + list(SCRIPTING_TAGS) + list(PRESENTATION_TAGS)

ABNORMAL_TAGS = list(SPECIAL_PAYLOAD_TAGS) + list(STATIC_TAGS) + list(SELF_CLOSING_TAGS)



#PAYLOAD_TAGS = helpers.minus(ALL_TAGS, SELF_CLOSING_TAGS)
#PAYLOAD_TAGS = helpers.minus(PAYLOAD_TAGS, STATIC_TAGS)

#html tags
#TEXTAREA    = 'textarea'

RESERVED_WORDS_LOWER = ('for', 'class', 'type',)
RESERVED_WORDS_UPPER = ('FOR', 'CLASS', 'TYPE',)


