from typing import Dict, Any, Optional
from lxml import etree
from .base import XMLParser

def _t(el: Optional[etree._Element]) -> Optional[str]:
    return el.text.strip() if el is not None and el.text else None

class NFSeABRASFParser(XMLParser):
    name = "NFS-e (ABRASF)"

    def matches(self, root: etree._Element) -> bool:
        # Heurística: presença de <CompNfse>, <Nfse>, <InfNfse> com nomes típicos ABRASF
        lname = etree.QName(root).localname
        if lname in {"CompNfse", "Nfse"}:
            return True
        return root.find(".//*[local-name()='InfNfse']") is not None

    def parse_header(self, root: etree._Element) -> Dict[str, Any]:
        # Tenta localizar InfNfse independentemente de namespace
        inf = root.find(".//*[local-name()='InfNfse']")
        ide = root.find(".//*[local-name()='IdentificacaoRps']") or root
        emit = root.find(".//*[local-name()='PrestadorServico']") or root
        dest = root.find(".//*[local-name()='TomadorServico']") or root
        valores = root.find(".//*[local-name()='Valores']")

        def find_text(xpath: str) -> Optional[str]:
            el = inf.find(xpath)
            return _t(el) if el is not None else None

        def any_text(node, tag) -> Optional[str]:
            if node is None:
                return None
            el = node.find(f".//*[local-name()='{tag}']")
            return _t(el)

        return {
            "tipo": "NFSe",
            "numero": any_text(ide, "Numero"),
            "serie": any_text(ide, "Serie"),
            "emissao": any_text(inf, "DataEmissao"),
            "emit_CNPJ": any_text(emit, "Cnpj") or any_text(emit, "Cpf"),
            "emit_xNome": any_text(emit, "RazaoSocial"),
            "dest_CNPJ": any_text(dest, "Cnpj") or any_text(dest, "Cpf"),
            "dest_xNome": any_text(dest, "RazaoSocial") or any_text(dest, "Nome"),
            "vNF": any_text(valores, "ValorServicos") or any_text(inf, "OutrasInformacoes"),
            "municipio": any_text(inf, "CodigoMunicipio"),
        }
