import requests
import os

CPF = "10628518625"
SENHA = os.getenv("SIAFI_SENHA")  # ideal

# Endpoint
url = "https://servicos-siafi.tesouro.gov.br/siafi2026/services/tabelas/consultarTabelasAdministrativas"

# Caminhos dos certificados
CERT_CLIENTE = r"W:\B - TED\7 - AUTOMAÇÃO\Scripts\siafi\certificado\wssimecpfmecgovbr.pem"
CA_SIAFI     = r"W:\B - TED\7 - AUTOMAÇÃO\Scripts\siafi\certificado\cadeiaCA.pem"

# Headers
headers = {
    "Content-Type": "text/xml; charset=utf-8",
    "SOAPAction": ""
}

# Envelope SOAP
soap_xml = soap_xml = soap_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
   <soap:Header>
      <ns2:cabecalhoSIAFI xmlns:ns2="http://www.tesouro.gov.br/siafi/services/tabelas/consultarTabelasAdministrativas">
         <ug>152734</ug>
         <bilhetador>
            <nonce>144966</nonce>
         </bilhetador>
      </ns2:cabecalhoSIAFI>

      <wsse:Security xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
         <wsse:UsernameToken>
            <wsse:Username>{CPF}</wsse:Username>
            <wsse:Password Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText">
               {SENHA}
            </wsse:Password>
         </wsse:UsernameToken>
      </wsse:Security>
   </soap:Header>

   <soap:Body>
      <ns2:tabConsultarSaldoContabil xmlns:ns2="http://www.tesouro.gov.br/siafi/services/tabelas/consultarTabelasAdministrativas">
         <tabConsultarSaldo>
            <codUG>152734</codUG>
            <contaContabil>622110000</contaContabil>
            <mesRefSaldo>JAN</mesRefSaldo>
            <outrosParamSaldoContabil>
               <codFonteRec>1000A0008U</codFonteRec>
               <codPtres>229566</codPtres>
            </outrosParamSaldoContabil>   
         </tabConsultarSaldo>
      </ns2:tabConsultarSaldoContabil>
   </soap:Body>
</soap:Envelope>
"""

# Requisição SOAP com mTLS
response = requests.post(
    url,
    data=soap_xml.encode("utf-8"),
    headers=headers,
    cert=CERT_CLIENTE,   # <<< certificado cliente (mTLS)
    verify=CA_SIAFI,     # <<< CA do SIAFI 
    timeout=60
)

print("Status HTTP:", response.status_code)
print("Resposta SOAP:")
print(response.text)