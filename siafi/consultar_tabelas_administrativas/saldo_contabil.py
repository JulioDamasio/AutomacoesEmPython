from siafi.auth.security import WSSESecurity
from siafi.auth.bilhetagem import CabecalhoSIAFI
from siafi.base.client import SIAFIClient


class SaldoContabilService:
    def __init__(self, client, cabecalho, security):
        self.client = client
        self.cabecalho = cabecalho
        self.security = security

    def consultar(
        self,
        cod_ug,
        conta_contabil,
        mes_ref,
        cod_fonte,
        cod_ptres
    ):
        soap_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Header>
        {self.cabecalho.render()}
        {self.security.render()}
    </soap:Header>

    <soap:Body>
        <ns2:tabConsultarSaldoContabil xmlns:ns2="http://www.tesouro.gov.br/siafi/services/tabelas/consultarTabelasAdministrativas">
            <tabConsultarSaldo>
                <codUG>{cod_ug}</codUG>
                <contaContabil>{conta_contabil}</contaContabil>
                <mesRefSaldo>{mes_ref}</mesRefSaldo>
                <outrosParamSaldoContabil>
                    <codFonteRec>{cod_fonte}</codFonteRec>
                    <codPtres>{cod_ptres}</codPtres>
                </outrosParamSaldoContabil>
            </tabConsultarSaldo>
        </ns2:tabConsultarSaldoContabil>
    </soap:Body>
</soap:Envelope>
"""
        return self.client.post(soap_xml)