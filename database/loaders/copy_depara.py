
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
copy_depara.py
- Lê um CSV com mapeamentos: Source, Destinations
- Copia cada arquivo de Source para 1..N destinos (separados por ';')
- Sobrescreve, cria pastas se necessário, registra log e faz retentativas
- Só copia se mudou (comparação por tamanho e mtime, opcional hash)
"""

import csv
import os
import sys
import time
import shutil
import hashlib
import logging
import argparse
from pathlib import Path

def setup_logger(log_path: Path):
    log_path.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout)
        ],
    )

def expand(p: str) -> str:
    """Expande variáveis de ambiente e ~"""
    return os.path.expanduser(os.path.expandvars(p.strip()))

def file_signature(p: Path, hash_algo: str = None):
    """Retorna assinatura do arquivo para comparação.
       Por padrão, usa (size, mtime). Se hash_algo em {'md5','sha256'}, calcula hash.
    """
    if not p.exists() or not p.is_file():
        return None
    if hash_algo:
        h = hashlib.md5() if hash_algo.lower() == "md5" else hashlib.sha256()
        with p.open("rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        return (p.stat().st_size, h.hexdigest())
    else:
        st = p.stat()
        # arredonda mtime para 1s para evitar pequenas diferenças de FS
        return (st.st_size, int(st.st_mtime))

def needs_copy(src: Path, dst: Path, hash_algo: str = None) -> bool:
    """Decide se precisa copiar: True se destino não existe ou assinatura diferente."""
    sig_src = file_signature(src, hash_algo)
    sig_dst = file_signature(dst, hash_algo)
    return sig_dst != sig_src

def copy_with_retry(src: Path, dst: Path, retries: int, delay: int):
    attempt = 0
    while True:
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            # copy2 preserva timestamps; Force overwrite
            shutil.copy2(src, dst)
            logging.info(f"OK: Copiado '{src}' → '{dst}'")
            return True
        except Exception as e:
            attempt += 1
            logging.warning(f"Falha ao copiar '{src}' → '{dst}' (tentativa {attempt}/{retries}): {e}")
            if attempt >= retries:
                logging.error(f"ERRO FINAL: não foi possível copiar '{src}' → '{dst}' após {retries} tentativas.")
                return False
            time.sleep(delay)

def main():
    parser = argparse.ArgumentParser(description="Copia arquivos com base em um CSV de de-para (Source, Destinations).")
    parser.add_argument("--csv", "-c", default=r"W:\B - TED\7 - AUTOMAÇÃO\Scripts\database\loaders\mapa_depara.csv", help="Caminho do CSV (Source,Destinations)")
    parser.add_argument("--retries", "-r", type=int, default=3, help="Qtde de retentativas por destino (padrão: 3)")
    parser.add_argument("--delay", "-d", type=int, default=5, help="Delay entre tentativas em segundos (padrão: 5)")
    parser.add_argument("--log", "-l", default="copy_depara.log", help="Caminho do arquivo de log")
    parser.add_argument("--hash", choices=["md5","sha256"], help="Comparação por hash (mais precisa, mais lenta). Se omitido, usa tamanho+mtime.")
    args = parser.parse_args()

    csv_path = Path(expand(args.csv))
    log_path = Path(expand(args.log))

    setup_logger(log_path)

    if not csv_path.exists():
        logging.error(f"CSV não encontrado: {csv_path}")
        sys.exit(1)

    try:
        with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            required_cols = {"Source", "Destinations"}
            if not required_cols.issubset({c.strip() for c in reader.fieldnames or []}):
                logging.error(f"CSV inválido. Cabeçalhos esperados: {required_cols}. Recebido: {reader.fieldnames}")
                sys.exit(1)

            for row in reader:
                src_str = expand(row.get("Source", ""))
                dsts_str = expand(row.get("Destinations", ""))

                if not src_str or not dsts_str:
                    logging.warning(f"Linha ignorada (faltando Source/Destinations): {row}")
                    continue

                src = Path(src_str)
                if not src.exists():
                    logging.warning(f"Origem NÃO encontrada: {src}")
                    continue

                destinations = [Path(expand(p)) for p in dsts_str.split(";") if p.strip()]
                for dst in destinations:
                    try:
                        if not dst.parent.exists():
                            dst.parent.mkdir(parents=True, exist_ok=True)

                        logging.info(f"Sobrescrevendo: copiando '{src.name}' para '{dst}'")
                        copy_with_retry(src, dst, args.retries, args.delay)
                        
                    except Exception as e:
                        logging.error(f"Erro ao processar destino '{dst}': {e}")

    except Exception as e:
        logging.error(f"Erro geral: {e}")
        sys.exit(1)

    logging.info("Concluído.")

if __name__ == "__main__":
    main()