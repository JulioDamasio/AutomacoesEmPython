from siafi.auth.security import WSSESecurity
from siafi.auth.bilhetagem import CabecalhoSIAFI
from siafi.base.client import SIAFIClient


class RetiradaDetalhamentoService:
    def __init__(self, client, cabecalho, security):
        self.client = client
        self.cabecalho = cabecalho
        self.security = security
        
        