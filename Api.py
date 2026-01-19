import requests

# Endpoint de homologação
url = "https://homextservicos-siafi.tesouro.gov.br/siafi2026he/services/tabelas/consultarTabelasAdministrativas?wsdl"

response = requests.get(
    url,
    cert=r"C:\Users\juliodamasio\Documents\meu_certificado.pfx",
    timeout=30
)

print(response.status_code)

# Certificado (arquivo .pem e .key)
cert = (
    
)

# Headers HTTP
headers = {
    "Content-Type": "text/xml; charset=utf-8",
    "SOAPAction": ""
}

# Envelope SOAP (SEU XML, sem alterações estruturais)
soap_xml = """<?xml version="1.0" encoding="UTF-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
   <soap:Header>
      <ns2:cabecalhoSIAFI xmlns:ns2="http://www.tesouro.gov.br/siafi/services/tabelas/consultarTabelasAdministrativas">
         <ug>173057</ug>
         <bilhetador>
            <nonce>144966</nonce>
         </bilhetador>
      </ns2:cabecalhoSIAFI>

      <wsse:Security xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
         <wsse:UsernameToken>
            <wsse:Username>10628518625</wsse:Username>
            <wsse:Password Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText">
            
            </wsse:Password>
         </wsse:UsernameToken>
      </wsse:Security>
   </soap:Header>

   <soap:Body>
      <ns2:tabConsultarSaldoContabil xmlns:ns2="http://www.tesouro.gov.br/siafi/services/tabelas/consultarTabelasAdministrativas">
         <tabConsultarSaldo>
            <codUG></codUG>
            <contaContabil></contaContabil>
            <contaCorrente></contaCorrente>
         </tabConsultarSaldo>
      </ns2:tabConsultarSaldoContabil>
   </soap:Body>
</soap:Envelope>
"""

# Requisição
response = requests.post(
    url,
    data=soap_xml.encode("utf-8"),
    headers=headers,
    cert=cert,
    verify=True,  # valida SSL
    timeout=30
)

# Resultado
print("Status HTTP:", response.status_code)
print("Resposta SOAP:")
print(response.text)