class CabecalhoSIAFI:
    def __init__(self, ug: str, nonce: str):
        self.ug = ug
        self.nonce = nonce

    def render(self) -> str:
        return f"""
        <ns2:cabecalhoSIAFI xmlns:ns2="http://www.tesouro.gov.br/siafi/services/tabelas/consultarTabelasAdministrativas">
            <ug>{self.ug}</ug>
            <bilhetador>
                <nonce>{self.nonce}</nonce>
            </bilhetador>
        </ns2:cabecalhoSIAFI>
        """