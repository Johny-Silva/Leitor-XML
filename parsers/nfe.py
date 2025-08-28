# parsers/nfe.py
from lxml import etree
from typing import Optional, Dict

NFE_NS = "http://www.portalfiscal.inf.br/nfe"

def _txt(node: Optional[etree._Element], tag: str, ns: str = NFE_NS) -> Optional[str]:
    if node is None:
        return None
    el = node.find(f"{{{ns}}}{tag}")
    return el.text.strip() if (el is not None and el.text) else None

def _to_number(x: Optional[str]) -> Optional[float]:
    if x is None:
        return None
    s = str(x).strip()
    if s == "":
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None

class NFeParser:
    """Parser para NF-e (modelo 55). Extrai cabeçalho, emit/dest, totais e CFOPs por item."""
    name = "NF-e"

    def matches(self, root: etree._Element) -> bool:
        # aceita <NFe> ou <procNFe> com <NFe> dentro
        nfe = root.find(f".//{{{NFE_NS}}}NFe")
        if nfe is None:
            # pode ser que o root JÁ seja NFe
            try:
                if etree.QName(root).localname == "NFe":
                    nfe = root
            except Exception:
                pass
        if nfe is None:
            return False

        # confere o modelo (55) quando possível
        ide = nfe.find(f".//{{{NFE_NS}}}ide")
        mod = _txt(ide, "mod")
        return (mod == "55") or (mod is None)  # se não tiver, ainda assim é muito provavelmente NF-e

    def parse_header(self, root: etree._Element) -> Dict:
        # normaliza ponto de entrada (NFe mesmo quando vier procNFe)
        nfe = root.find(f".//{{{NFE_NS}}}NFe")
        if nfe is None:
            nfe = root

        inf = nfe.find(f".//{{{NFE_NS}}}infNFe")
        ide = nfe.find(f".//{{{NFE_NS}}}ide")
        emit = nfe.find(f".//{{{NFE_NS}}}emit")
        dest = nfe.find(f".//{{{NFE_NS}}}dest")
        total = nfe.find(f".//{{{NFE_NS}}}total/{{{NFE_NS}}}ICMSTot")

        # chave (Id começa com "NFe")
        chave = (inf.get("Id") if inf is not None else None) or ""
        chave = chave.replace("NFe", "") or None

        # ide
        nNF = _txt(ide, "nNF")
        serie = _txt(ide, "serie")
        tpNF = _txt(ide, "tpNF")  # 0=Entrada, 1=Saída
        emissao = _txt(ide, "dhEmi") or _txt(ide, "dEmi")
        modelo = _txt(ide, "mod") or "55"

        # emit/dest
        emit_cnpj = _txt(emit, "CNPJ") or _txt(emit, "CPF")
        emit_nome = _txt(emit, "xNome")
        dest_cnpj = _txt(dest, "CNPJ") or _txt(dest, "CPF")
        dest_nome = _txt(dest, "xNome")

        # totais
        vNF = _to_number(_txt(total, "vNF"))

        # ---------- CFOP por item ----------
        cfops = []
        # também calcular CFOP predominante (pelo somatório de vProd)
        soma_por_cfop = {}
        for det in nfe.findall(f".//{{{NFE_NS}}}det"):
            prod = det.find(f"./{{{NFE_NS}}}prod")
            if prod is None:
                continue
            cfop = _txt(prod, "CFOP")
            if cfop:
                cfops.append(cfop)

            vprod = _to_number(_txt(prod, "vProd"))
            if cfop:
                soma_por_cfop[cfop] = soma_por_cfop.get(cfop, 0.0) + (vprod or 0.0)

        cfops_unicos = "; ".join(sorted(set(cfops))) if cfops else None
        cfop_pred = None
        if soma_por_cfop:
            cfop_pred = max(soma_por_cfop.items(), key=lambda kv: kv[1])[0]

        return {
            "chave": chave,
            "nNF": nNF,
            "serie": serie,
            "tpNF": tpNF,
            "emissao": emissao,
            "emit_CNPJ": emit_cnpj,
            "emit_xNome": emit_nome,
            "dest_CNPJ": dest_cnpj,
            "dest_xNome": dest_nome,
            "vNF": vNF,
            "modelo": modelo,
            # 🆕 CFOPs
            "CFOPs_itens": cfops_unicos,         # ex.: "5102; 5405"
            "CFOP_predominante": cfop_pred,      # ex.: "5102"
        }
