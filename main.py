from siafi.base.client import SIAFIClient
from siafi.auth.security import WSSESecurity
from siafi.auth.bilhetagem import CabecalhoSIAFI
from siafi.consultar_tabelas_administrativas.saldo_contabil import SaldoContabilService

client = SIAFIClient(
    url="https://servicos-siafi.tesouro.gov.br/siafi2026/services/tabelas/consultarTabelasAdministrativas",
    cert=r"W:\B - TED\7 - AUTOMAÇÃO\Scripts\siafi\certificado\wssimecpfmecgovbr.pem",
    ca=r"W:\B - TED\7 - AUTOMAÇÃO\Scripts\siafi\certificado\cadeiaCA.pem"
)

security = WSSESecurity(
    cpf="10628518625",
    senha_env="SIAFI_SENHA"
)

cabecalho = CabecalhoSIAFI(
    ug="152734",
    nonce="144966"
)

service = SaldoContabilService(client, cabecalho, security)

response = service.consultar(
    cod_ug="152734",
    conta_contabil="622110000",
    mes_ref="JAN",
    cod_fonte="1000A0008U",
    cod_ptres="229566"
)

print(response.status_code)
print(response.text)