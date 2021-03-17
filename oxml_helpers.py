
from docx.text.paragraph import Paragraph
from docx.oxml.text.paragraph import CT_P
from docx.oxml.simpletypes import ST_String
from docx.oxml import *
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
    OneAndOnlyOne,
)

class CT_Tag(BaseOxmlElement):
    """
    Used for ``<w:b>``, ``<w:i>`` elements and others, containing a bool-ish
    string in its ``val`` attribute, xsd:boolean plus 'on' and 'off'.
    """
    val = OptionalAttribute('w:val', ST_String, default=True)

class CT_Alias(BaseOxmlElement):
    """
    Used for ``<w:b>``, ``<w:i>`` elements and others, containing a bool-ish
    string in its ``val`` attribute, xsd:boolean plus 'on' and 'off'.
    """
    val = OptionalAttribute('w:val', ST_String, default=True)

class CT_BlockProperties(BaseOxmlElement):
    """
    Used for ``<w:b>``, ``<w:i>`` elements and others, containing a bool-ish
    string in its ``val`` attribute, xsd:boolean plus 'on' and 'off'.
    """
    # val = OptionalAttribute('w:val', ST_String, default=True)
    tag = ZeroOrOne('w:tag')
    alias = ZeroOrOne('w:alias')

class CT_Block(BaseOxmlElement):
    """
    Used for ``<w:b>``, ``<w:i>`` elements and others, containing a bool-ish
    string in its ``val`` attribute, xsd:boolean plus 'on' and 'off'.
    """
    # val = OptionalAttribute('w:val', ST_String, default=True)
    properties = ZeroOrOne('w:sdtPr', successors=())
    content = ZeroOrOne('w:sdtContent', successors=())

    def hasTag(self, tag):
        if self.properties != None:
            if self.properties.tag != None:
                if self.properties.tag.val == tag:
                    return True
        return False
    
    def getContent(self):
        return self.content


class CT_BlockContent(BaseOxmlElement):
    """
    Used for ``<w:b>``, ``<w:i>`` elements and others, containing a bool-ish
    string in its ``val`` attribute, xsd:boolean plus 'on' and 'off'.
    """
    p = ZeroOrMore('w:p')
    sdt = ZeroOrMore('w:sdt')
    r = ZeroOrMore('w:r')

    def hasTag(self, tag):
        for s in self.sdt_lst:
            if s.hasTag(tag) == True:
                return True
                # if s.getContent().hasField(tag):
                #     return True

        for p in self.p_lst:
            for s in p.sdt_lst:
                if s.hasTag(tag) == True:
                    return True
                    # if s.getContent().hasField(tag):
                    #     return True
        return False

    def getTag(self, tag):
        for s in self.sdt_lst:
            if s.hasTag(tag) == True:
                return s.getContent()

        for p in self.p_lst:
            for s in p.sdt_lst:
                if s.hasTag(tag) == True:
                    return s.getContent()
                # if s.getContent().hasField(tag):
                #     return s.getContent().getContent()
        return None

    def hasField(self, field):
        if self.hasTag('fields'):
            c = self.getTag('fields')
            if c.hasTag(field):
                return True
        return False

    def getField(self, field):
        if self.hasTag('fields'):
            c = self.getTag('fields')
            if c.hasTag(field):
                return c.getTag(field)
        return None


    @property
    def all_text(self):
        res = ''
        for r in self.r_lst:
            for t in r.t_lst:
                res += t.text
        return res

class CT_P_Custom(CT_P):
    """
    ``<w:p>`` element, containing the properties and text for a paragraph.
    """
    sdt = ZeroOrMore('w:sdt')

    def getFields(self):
        for s in self.sdt_lst:
            if s.properties != None:
                if s.properties.tag != None:
                    if s.properties.tag.val == 'fields':
                        yield s.content

    def hasFieldInSdt(self):
        for s in self.sdt_lst:
            if s.properties != None:
                if s.properties.tag != None:
                    if s.properties.tag.val == 'fields':
                        print('yay!')
                        return True

        return False


# register_element_cls()
register_element_cls('w:alias', CT_Alias)
register_element_cls('w:tag', CT_Tag)
register_element_cls('w:sdt', CT_Block)
register_element_cls('w:sdtPr', CT_BlockProperties)
register_element_cls('w:sdtContent', CT_BlockContent)
register_element_cls('w:p', CT_P_Custom)