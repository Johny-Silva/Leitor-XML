# utils/xml.py  (novo)
from lxml import etree
import threading

_thread_local = threading.local()

def get_xml_parser():
    p = getattr(_thread_local, "xml_parser", None)
    if p is None:
        p = etree.XMLParser(
            encoding="utf-8",
            recover=True,              # tolera pequenos problemas no XML
            remove_blank_text=True,    # menos nós
            remove_comments=True,
            remove_pis=True,
            huge_tree=False,           # true só se pegar XMLs gigantescos (deixe false p/ ganhar velocidade)
        )
        _thread_local.xml_parser = p
    return p
