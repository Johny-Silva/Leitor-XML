from .nfe import NFeParser
from .nfce import NFCeParser
from .nfse_abrasf import NFSeABRASFParser
from .evento_nfe import NFeEventParser
from .nfse_rn_prestado import NFSERNPrestadoParser
from .nfse_rn_tomado import NFSERNTomadoParser
from .cte import CTeParser

ALL_PARSERS = [
    NFeParser(),
    NFCeParser(),
    NFSeABRASFParser(),
    NFeEventParser(),
    NFSERNPrestadoParser(),  # novo
    NFSERNTomadoParser(), 
    CTeParser()   # novo
]

def get_parser_by_name(name: str):
    for p in ALL_PARSERS:
        if p.name == name:
            return p
    raise ValueError(f"Tipo de nota não suportado: {name}")
