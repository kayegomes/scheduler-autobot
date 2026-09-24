"""Reverte grades no-op que soterraram a escala base.

Antes da correção, um cruzamento sem efeito com grade de TV salvava uma cópia
idêntica da base como se fosse uma grade nova. Cada cópia virava a nova base de
comparação — inclusive por cima de uma escala carregada manualmente.

Este script identifica grades cujo conteúdo é byte a byte igual ao de uma grade
anterior e as marca como 'ignored', devolvendo a base para a última escala real.
Nada é apagado: reverter é trocar o status de volta para 'done'.

    python scripts/corrigir_baseline.py            # simulação (não grava)
    python scripts/corrigir_baseline.py --apply    # aplica
"""

import hashlib
import os
import shutil
import sqlite3
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import DB_PATH  # noqa: E402


def assinatura(con, grade_id):
    h = hashlib.sha256()
    for (raw,) in con.execute(
        "SELECT raw_data FROM grade_items WHERE grade_id = ? ORDER BY id", (grade_id,)
    ):
        h.update((raw or "").encode("utf-8", "replace"))
    return h.hexdigest()


def resumo(con, grade_id):
    row = con.execute(
        "SELECT filename, import_date, status FROM grades WHERE id = ?", (grade_id,)
    ).fetchone()
    n = con.execute(
        "SELECT COUNT(*) FROM grade_items WHERE grade_id = ?", (grade_id,)
    ).fetchone()[0]
    datas = [
        d[0] for d in con.execute(
            "SELECT DISTINCT data FROM grade_items WHERE grade_id = ? AND data <> ''", (grade_id,)
        )
    ]
    periodo = "-"
    if datas:
        ordenadas = sorted(
            datas, key=lambda s: (s[6:10], s[3:5], s[0:2]) if len(s) == 10 else ("", "", "")
        )
        periodo = f"{ordenadas[0]} a {ordenadas[-1]}"
    return f"{row[2]:<8} {row[1]} {n:>5} linhas  {periodo:<24} {row[0][:46]}"


def main():
    aplicar = "--apply" in sys.argv

    if not os.path.exists(DB_PATH):
        print(f"[ERRO] Banco não encontrado: {DB_PATH}")
        return 1

    con = sqlite3.connect(DB_PATH)
    grades = [r[0] for r in con.execute("SELECT id FROM grades ORDER BY id")]

    vistos = {}
    duplicadas = []
    for gid in grades:
        status = con.execute("SELECT status FROM grades WHERE id = ?", (gid,)).fetchone()[0]
        sig = assinatura(con, gid)
        if not sig or sig == hashlib.sha256().hexdigest():
            continue  # grade sem itens (ex.: já ignorada)
        if sig in vistos:
            if status == 'done':
                duplicadas.append((gid, vistos[sig]))
        else:
            vistos[sig] = gid

    base_antes = con.execute(
        "SELECT id FROM grades WHERE status='done' ORDER BY id DESC LIMIT 1"
    ).fetchone()
    print(f"Banco: {DB_PATH}")
    print(f"\nBase de comparação ATUAL:\n  {base_antes[0]:>3} {resumo(con, base_antes[0])}")

    if not duplicadas:
        print("\nNenhuma grade duplicada encontrada. Nada a fazer.")
        con.close()
        return 0

    print(f"\n{len(duplicadas)} grade(s) que são cópia de uma anterior:")
    for gid, origem in duplicadas:
        print(f"  {gid:>3} {resumo(con, gid)}  (cópia da {origem})")

    ids = [g for g, _ in duplicadas]
    marcadores = ",".join("?" * len(ids))
    nova_base = con.execute(
        f"SELECT id FROM grades WHERE status='done' AND id NOT IN ({marcadores}) "
        f"ORDER BY id DESC LIMIT 1", ids
    ).fetchone()

    if not nova_base:
        print("\n[ABORTADO] Marcar essas grades deixaria o sistema sem base. Nada foi alterado.")
        con.close()
        return 1

    print(f"\nBase de comparação DEPOIS:\n  {nova_base[0]:>3} {resumo(con, nova_base[0])}")

    if not aplicar:
        print("\n--- SIMULAÇÃO: nada foi gravado. Rode com --apply para aplicar. ---")
        con.close()
        return 0

    backup = f"{DB_PATH}.antes_corrigir_baseline.{datetime.now():%Y%m%d_%H%M%S}.bak"
    con.close()
    shutil.copy2(DB_PATH, backup)
    print(f"\nBackup: {backup}")

    con = sqlite3.connect(DB_PATH)
    con.executemany("UPDATE grades SET status='ignored' WHERE id = ?", [(g,) for g in ids])
    con.commit()
    base_depois = con.execute(
        "SELECT id FROM grades WHERE status='done' ORDER BY id DESC LIMIT 1"
    ).fetchone()[0]
    con.close()

    print(f"{len(ids)} grade(s) marcadas como 'ignored'.")
    print(f"Base de comparação agora: {base_depois}")
    print("\nPara reverter:")
    print(f"  UPDATE grades SET status='done' WHERE id IN ({','.join(map(str, ids))});")
    return 0


if __name__ == "__main__":
    sys.exit(main())
