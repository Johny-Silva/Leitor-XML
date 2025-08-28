from typing import Dict, Any, Optional
from lxml import etree
from .base import XMLParser

ABRASF_NS = "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd"

def _t(el: Optional[etree._Element]) -> Optional[str]:
    return el.text.strip() if el is not None and el.text else None

def _find(node: Optional[etree._Element], tag: str) -> Optional[etree._Element]:
    if node is None:
        return None
    return node.find(f"{{{ABRASF_NS}}}{tag}")

def _g(node: Optional[etree._Element], tag: str) -> Optional[str]:
    return _t(_find(node, tag)) if node is not None else None

class NFSERNPrestadoParser(XMLParser):
    name = "NFSe RN (Prestado)"

    def matches(self, root: etree._Element) -> bool:
        try:
            inf = root.find(f".//{{{ABRASF_NS}}}InfNfse")
            if inf is None:
                return False
            org = _find(inf, "OrgaoGerador")
            uf  = _g(org, "Uf")
            if uf != "RN":
                return False
            tom = _find(inf, "TomadorServico")
            cpf = tom.find(f".//{{{ABRASF_NS}}}Cpf") if tom is not None else None
            return cpf is not None
        except Exception:
            return False

    def parse_header(self, root: etree._Element) -> Dict[str, Any]:
        inf = root.find(f".//{{{ABRASF_NS}}}InfNfse")

        numero         = _g(inf, "Numero")
        cod_ver        = _g(inf, "CodigoVerificacao")
        emissao        = _g(inf, "DataEmissao")
        competencia    = _g(inf, "Competencia")

        servico        = _find(inf, "Servico")
        valores        = _find(servico, "Valores") if servico is not None else None
        v_servicos     = _g(valores, "ValorServicos")
        v_iss          = _g(valores, "ValorIss")
        aliquota       = _g(valores, "Aliquota")
        iss_retido     = _g(valores, "IssRetido")

        item_lista     = _g(servico, "ItemListaServico") if servico is not None else None
        cnae           = _g(servico, "CodigoCnae") if servico is not None else None
        discriminacao  = _g(servico, "Discriminacao") if servico is not None else None
        cod_municipio  = _g(servico, "CodigoMunicipio") if servico is not None else None

        prest          = _find(inf, "PrestadorServico")
        prest_id       = _find(prest, "IdentificacaoPrestador")
        prest_cnpj     = _g(prest_id, "Cnpj")
        prest_im       = _g(prest_id, "InscricaoMunicipal")
        prest_razao    = _g(prest, "RazaoSocial")

        tomador        = _find(inf, "TomadorServico")
        tom_cnpj       = _t(tomador.find(f".//{{{ABRASF_NS}}}Cnpj")) if tomador is not None else None
        tom_cpf        = _t(tomador.find(f".//{{{ABRASF_NS}}}Cpf")) if tomador is not None else None
        tom_razao      = _g(tomador, "RazaoSocial")

        org            = _find(inf, "OrgaoGerador")
        org_mun        = _g(org, "CodigoMunicipio")
        org_uf         = _g(org, "Uf")

        return {
            "tipo": "NFSe",
            "modelo_nfse": "RN",
            "sentido_nfse": "Prestado",
            "numero": numero,
            "codigoVerificacao": cod_ver,
            "emissao": emissao,
            "competencia": competencia,
            "emit_CNPJ": prest_cnpj,
            "emit_IM": prest_im,
            "emit_xNome": prest_razao,
            "dest_CNPJ": tom_cnpj or tom_cpf,
            "dest_xNome": tom_razao,
            "vNF": v_servicos,
            "valor_iss": v_iss,
            "aliquota": aliquota,
            "iss_retido": iss_retido,
            "itemListaServico": item_lista,
            "codigoCNAE": cnae,
            "discriminacao": discriminacao,
            "codigoMunicipioServico": cod_municipio,
            "orgaoGeradorCodigo": org_mun,
            "orgaoGeradorUF": org_uf,
        }
