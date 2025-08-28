from typing import Dict, Any, Optional
from lxml import etree
from .base import XMLParser

NS = "http://www.portalfiscal.inf.br/cte"

def _t(el: Optional[etree._Element]) -> Optional[str]:
    return el.text.strip() if el is not None and el.text else None

def _g(node: Optional[etree._Element], tag: str) -> Optional[str]:
    if node is None:
        return None
    el = node.find(f"{{{NS}}}{tag}")
    return _t(el)

class CTeParser(XMLParser):
    name = "CT-e"

    def matches(self, root: etree._Element) -> bool:
        try:
            cte = root.find(f".//{{{NS}}}CTe")
            if cte is None:
                cte = root
            ide = cte.find(f".//{{{NS}}}ide")
            mod = _g(ide, "mod") if ide is not None else None
            return mod == "57"
        except Exception:
            return False

    def parse_header(self, root: etree._Element) -> Dict[str, Any]:
        cte = root.find(f".//{{{NS}}}CTe")
        if cte is None:
            cte = root
        inf = cte.find(f".//{{{NS}}}infCte")

        ide      = inf.find(f".//{{{NS}}}ide")      if inf is not None else None
        emit     = inf.find(f".//{{{NS}}}emit")     if inf is not None else None
        rem      = inf.find(f".//{{{NS}}}rem")      if inf is not None else None
        dest     = inf.find(f".//{{{NS}}}dest")     if inf is not None else None
        vPrest   = inf.find(f".//{{{NS}}}vPrest")   if inf is not None else None
        infCarga = inf.find(f".//{{{NS}}}infCarga") if inf is not None else None
        infProt  = root.find(f".//{{{NS}}}protCTe/{{{NS}}}infProt")

        chave = (inf.get("Id") or "").replace("CTe", "") if inf is not None else None

        return {
            "tipo": "CT-e",
            "chave": chave,
            "nCT": _g(ide, "nCT"),
            "serie": _g(ide, "serie"),
            "emissao": _g(ide, "dhEmi"),
            "tpCTe": _g(ide, "tpCTe"),
            "CFOP": _g(ide, "CFOP"),
            "natOp": _g(ide, "natOp"),
            "emit_CNPJ": _g(emit, "CNPJ"),
            "emit_xNome": _g(emit, "xNome"),
            "rem_CNPJ": _g(rem, "CNPJ"),
            "rem_xNome": _g(rem, "xNome"),
            "dest_CNPJ": _g(dest, "CNPJ"),
            "dest_xNome": _g(dest, "xNome"),
            "vTPrest": _g(vPrest, "vTPrest"),
            "vRec": _g(vPrest, "vRec"),
            "vCarga": _g(infCarga, "vCarga"),
            "status": _g(infProt, "cStat"),
            "autorizacao": _g(infProt, "xMotivo"),
        }
