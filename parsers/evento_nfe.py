from typing import Dict, Any, Optional
from lxml import etree
from .base import XMLParser

NFE_NS = "http://www.portalfiscal.inf.br/nfe"

def _txt(node, tag, ns=NFE_NS):
    if node is None:
        return None
    el = node.find(f"{{{ns}}}{tag}")
    return el.text.strip() if (el is not None and el.text) else None

class NFeEventParser:
    """Parser para procEventoNFe (eventos de NF-e: cancelamento 110111, CCe 110110, etc.)."""
    name = "Evento NF-e"

    def matches(self, root: etree._Element) -> bool:
        """
        Reconhece:
          - raiz <procEventoNFe> ou <evento>
          - ou quando o arquivo tem esses nós em qualquer nível
        """
        try:
            lname = etree.QName(root).localname
        except Exception:
            lname = (root.tag or "").split("}")[-1]
        if lname in {"procEventoNFe", "evento"}:
            return True
        if root.find(f".//{{{NFE_NS}}}procEventoNFe") is not None:
            return True
        if root.find(f".//{{{NFE_NS}}}evento") is not None:
            return True
        return False

    def parse_header(self, root: etree._Element) -> dict:
        """
        Extrai campos padronizados usados pelo app:
          - chNFe, tpEvento, descEvento, dhEvento, nProt_retEvento, emit_CNPJ
        Aceita tanto <procEventoNFe> (com <retEvento>) quanto <evento> “cru”.
        """
        # tenta encontrar o nó <evento>
        evento = root.find(f".//{{{NFE_NS}}}evento")
        if evento is None:
            # às vezes o root já é <evento>
            try:
                if etree.QName(root).localname == "evento":
                    evento = root
            except Exception:
                pass
        if evento is None:
            # fallback: se tudo falhar, usa root mesmo (evita crash)
            evento = root

        inf = evento.find(f".//{{{NFE_NS}}}infEvento")
        det = evento.find(f".//{{{NFE_NS}}}detEvento")

        # quando for procEventoNFe, existe o <retEvento> com <nProt>
        ret_evento = root.find(f".//{{{NFE_NS}}}retEvento")

        chNFe  = _txt(inf, "chNFe")
        tpEv   = _txt(inf, "tpEvento")
        dhEv   = _txt(inf, "dhEvento")
        descEv = _txt(det, "descEvento")

        # Protocolo pode vir no retEvento (mais comum) ou dentro de detEvento, dependendo da UF/versão
        nProt  = _txt(ret_evento, "nProt") or _txt(det, "nProt")

        # Autor do evento pode ser CNPJ OU CPF
        autor  = _txt(inf, "CNPJ") or _txt(inf, "CPF")

        return {
            "chNFe": chNFe,
            "tpEvento": tpEv,
            "descEvento": descEv,
            "dhEvento": dhEv,
            "nProt_retEvento": nProt,
            "emit_CNPJ": autor,   # ajuda a enriquecer linha sintética no app
        }