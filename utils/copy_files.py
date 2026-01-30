from pathlib import Path
import shutil
from typing import Iterable, List


class FilePreparer:
    def __init__(
        self,
        destino_base: Path,
        prefixo: str = "COPIA_",
        sobrescrever: bool = False
    ):
        self.destino_base = destino_base
        self.prefixo = prefixo
        self.sobrescrever = sobrescrever

        self.destino_base.mkdir(parents=True, exist_ok=True)

    def copiar_arquivo(self, origem: Path) -> Path:
        destino = self.destino_base / f"{self.prefixo}{origem.name}"

        if destino.exists() and not self.sobrescrever:
            raise FileExistsError(f"Arquivo já existe: {destino}")

        shutil.copy(origem, destino)
        return destino

    def copiar_varios(self, arquivos: Iterable[Path]) -> List[Path]:
        copiados = []

        for arquivo in arquivos:
            copiado = self.copiar_arquivo(arquivo)
            copiados.append(copiado)

        return copiados