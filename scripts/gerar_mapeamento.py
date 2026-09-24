"""Gera o esqueleto do mapeamento Nome=email a partir da escala vigente.

O sistema só notifica quem está em Configurações → Mapeamento de E-mails. Sem
essa tabela preenchida nenhuma notificação sai — e, antes da correção, isso
acontecia em silêncio.

Este script lista os funcionários da escala base atual, já no formato aceito
pela interface, preservando os e-mails que já estiverem cadastrados. Os
endereços NÃO são inventados: cada linha nova sai com o campo vazio para você
preencher.

    python scripts/gerar_mapeamento.py                  # imprime na tela
    python scripts/gerar_mapeamento.py -o mapa.txt      # grava em arquivo
    python scripts/gerar_mapeamento.py --dominio g.globo --sugerir
"""

import argparse
import os
import sys
import unicodedata

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.database import DatabaseManager  # noqa: E402


def sugerir(nome: str, dominio: str) -> str:
    """primeiro.ultimo@dominio — apenas um palpite, sempre revise."""
    partes = [
        "".join(
            ch for ch in unicodedata.normalize("NFKD", p.lower())
            if not unicodedata.combining(ch) and ch.isalnum()
        )
        for p in str(nome).split()
    ]
    partes = [p for p in partes if p and p not in ("de", "da", "do", "dos", "das", "e")]
    if not partes:
        return ""
    local = partes[0] if len(partes) == 1 else f"{partes[0]}.{partes[-1]}"
    return f"{local}@{dominio}"


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("-o", "--saida", help="arquivo de destino (padrão: imprime na tela)")
    ap.add_argument("--dominio", default="", help="domínio para sugerir endereços")
    ap.add_argument("--sugerir", action="store_true",
                    help="preenche com primeiro.ultimo@dominio (REVISE antes de usar)")
    args = ap.parse_args()

    db = DatabaseManager()
    df = db.get_last_grade_df()
    if df.empty or "funcionario" not in df.columns:
        print("[ERRO] Nenhuma escala base carregada. Use 'Carregar Escala Base Inicial'.")
        return 1

    nomes = sorted({
        str(n).strip() for n in df["funcionario"] if str(n).strip()
    })
    ja_mapeados = {k.strip().lower(): v for k, v in db.get_employee_mapping().items()}

    linhas, faltando = [], 0
    for nome in nomes:
        email = ja_mapeados.get(nome.lower(), "")
        if not email:
            faltando += 1
            if args.sugerir and args.dominio:
                email = sugerir(nome, args.dominio)
        linhas.append(f"{nome}={email}")

    texto = "\n".join(linhas)
    if args.saida:
        with open(args.saida, "w", encoding="utf-8") as f:
            f.write(texto + "\n")
        print(f"{len(linhas)} funcionário(s) gravados em {args.saida}")
    else:
        print(texto)

    print(f"\n-- {len(nomes)} na escala | {len(nomes) - faltando} já mapeados | "
          f"{faltando} sem e-mail --", file=sys.stderr)
    if args.sugerir:
        print("-- ATENÇÃO: endereços sugeridos são palpites. Confira um a um antes "
              "de desligar o Modo Teste. --", file=sys.stderr)
    print("-- Cole o conteudo em Configuracoes > Mapeamento de E-mails --", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
