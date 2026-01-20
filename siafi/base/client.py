import requests

class SIAFIClient:
    def __init__(self, url, cert, ca):
        self.url = url
        self.cert = cert
        self.ca = ca

    def post(self, soap_xml: str):
        headers = {
            "Content-Type": "text/xml; charset=utf-8",
            "SOAPAction": ""
        }

        response = requests.post(
            self.url,
            data=soap_xml.encode("utf-8"),
            headers=headers,
            cert=self.cert,
            verify=self.ca,
            timeout=60
        )

        return response