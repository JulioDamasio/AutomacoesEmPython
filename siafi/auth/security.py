import os

class WSSESecurity:
    def __init__(self, cpf: str, senha_env: str):
        self.cpf = cpf
        self.senha = os.getenv(senha_env)

        if not self.senha:
            raise ValueError("Senha do SIAFI não encontrada na variável de ambiente.")

    def render(self) -> str:
        return f"""
        <wsse:Security xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
            <wsse:UsernameToken>
                <wsse:Username>{self.cpf}</wsse:Username>
                <wsse:Password Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText">
                    {self.senha}
                </wsse:Password>
            </wsse:UsernameToken>
        </wsse:Security>
        """