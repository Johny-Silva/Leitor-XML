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

        # ===== Totais ICMSTot (primeira tentativa) =====
        vBC_tot  = _to_number(_txt(total, "vBC"))
        vICMS_tot = _to_number(_txt(total, "vICMS"))
        vBCST_tot = _to_number(_txt(total, "vBCST"))
        vST_tot   = _to_number(_txt(total, "vST"))

        # ===== Fallback por itens, se necessário =====
        # Soma por ICMS "próprio"
        if (vBC_tot is None or vBC_tot == 0.0) or (vICMS_tot is None or vICMS_tot == 0.0):
            s_vBC_it = 0.0
            s_vICMS_it = 0.0
            achou_icms_item = False
            for det in nfe.findall(f".//{{{NFE_NS}}}det"):
                icms = det.find(f"./{{{NFE_NS}}}imposto/{{{NFE_NS}}}ICMS")
                if icms is None:
                    continue
                # pega o primeiro grupo ICMS* existente
                grp = next((child for child in icms if isinstance(child.tag, str)), None)
                if grp is None:
                    continue
                vbc_i = _to_number(_txt(grp, "vBC"))
                vicms_i = _to_number(_txt(grp, "vICMS"))
                if vbc_i is not None:
                    s_vBC_it += vbc_i
                    achou_icms_item = True
                if vicms_i is not None:
                    s_vICMS_it += vicms_i
                    achou_icms_item = True
            if achou_icms_item:
                # só substitui se estava vazio/zero nos totais
                if (vBC_tot is None or vBC_tot == 0.0):
                    vBC_tot = s_vBC_it
                if (vICMS_tot is None or vICMS_tot == 0.0):
                    vICMS_tot = s_vICMS_it

        # Soma por ST (normal e retida)
        if (vBCST_tot is None or vBCST_tot == 0.0) or (vST_tot is None or vST_tot == 0.0):
            s_vBCST_it = 0.0
            s_vICMSST_it = 0.0
            achou_st_item = False
            for det in nfe.findall(f".//{{{NFE_NS}}}det"):
                icms = det.find(f"./{{{NFE_NS}}}imposto/{{{NFE_NS}}}ICMS")
                if icms is None:
                    continue
                grp = next((child for child in icms if isinstance(child.tag, str)), None)
                if grp is None:
                    continue
                # ST "normal"
                vbcst_i = _to_number(_txt(grp, "vBCST"))
                vicmsst_i = _to_number(_txt(grp, "vICMSST"))
                # ST retida (CST 60)
                vbcst_ret_i = _to_number(_txt(grp, "vBCSTRet"))
                vicmsst_ret_i = _to_number(_txt(grp, "vICMSSTRet"))
                # alguns emissores trazem também vICMSSubstituto
                vicms_subst_i = _to_number(_txt(grp, "vICMSSubstituto"))

                if vbcst_i is not None:
                    s_vBCST_it += vbcst_i
                    achou_st_item = True
                if vicmsst_i is not None:
                    s_vICMSST_it += vicmsst_i
                    achou_st_item = True

                if vbcst_ret_i is not None:
                    s_vBCST_it += vbcst_ret_i
                    achou_st_item = True
                if vicmsst_ret_i is not None:
                    s_vICMSST_it += vicmsst_ret_i
                    achou_st_item = True

                if vicms_subst_i is not None:
                    s_vICMSST_it += vicms_subst_i
                    achou_st_item = True

            if achou_st_item:
                if (vBCST_tot is None or vBCST_tot == 0.0):
                    vBCST_tot = s_vBCST_it
                if (vST_tot is None or vST_tot == 0.0):
                    vST_tot = s_vICMSST_it

        # Inclui no dict de retorno
        return {
            # ... (todos os campos que você já retorna)
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
            "CFOPs_itens": cfops_unicos,
            "CFOP_predominante": cfop_pred,
            # 🆕 Totais ICMS / ST (nota)
            "vBC_ICMS": vBC_tot,
            "vICMS": vICMS_tot,
            "vBC_ST": vBCST_tot,
            "vICMS_ST": vST_tot,
        }