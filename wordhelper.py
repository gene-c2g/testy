##################################################
## testy.py
##################################################
## Author: clenahan@cloud2gnd.com
## Copyright: Copyright 2024
## Version: 1.0
##################################################

try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML

WORD_NAMESPACE = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TEXT = WORD_NAMESPACE + "t"

def getAcceptedText(p):
    """Return text of a paragraph after accepting all changes"""
    xml = p._p.xml
    if "w:del" in xml or "w:ins" in xml:
        tree = XML(xml)
        runs = (node.text for node in tree.iter(TEXT) if node.text)
        # Note: on older versions it is `tree.getiterator` instead of `tree.iter`
        return "".join(runs)
    else:
        ret = ""
        for run in p.runs:
            if not (run.font.strike or run.style.font.strike):
                ret += run.text
        return ret

def Paras2Text(paragraphs):
    text = ""
    bFirst = True
    for para in paragraphs:
        text = text + "\n" if not bFirst else "" + getAcceptedText(para)
        bFirst = False
    return text

def addNewline(doc):
    p=doc.add_paragraph()
    run = p.add_run()
    run.add_break()