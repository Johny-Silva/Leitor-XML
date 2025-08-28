from abc import ABC, abstractmethod
from typing import Dict, Any
from lxml import etree

class XMLParser(ABC):
    name: str = "base"

    @abstractmethod
    def matches(self, root: etree._Element) -> bool:
        """Retorna True se o XML pertencer a este tipo de nota."""
        raise NotImplementedError

    @abstractmethod
    def parse_header(self, root: etree._Element) -> Dict[str, Any]:
        """Extrai os campos principais da nota."""
        raise NotImplementedError
