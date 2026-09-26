# -*- coding: utf-8 -*-
"""Диагностика нулевого матчинга граф-рёбер.

Берёт рёбра из graph_rag.db (приоритет — NULL source_chunk), подтягивает тексты
чанков окна из smk_docs и показывает: сущности триплета vs реальный текст окна.
Запуск: python scripts/diag_graph_matching.py [N]  (N = сколько окон показать, по умолч. 5)
"""
import os
import re
import sys
import sqlite3
import chromadb

DB_DIR = r"D:\Plaksunov VB - Uniscan\Programm\AI Agent QMS\smk_vector_db"
WORD = re.compile(r"[0-9a-zA-Zа-яА-ЯёЁ]")

def norm_entity(name):
    return re.sub(r"\s+", " ", (name or "")).strip().lower()

def entity_in_text(ne, nt):
    if not ne:
        return False
    pos = nt.find(ne)
    while pos != -1:
        before = nt[pos - 1] if pos > 0 else ""
        after = nt[pos + len(ne)] if pos + len(ne) < len(nt) else ""
        if ((not before) or not WORD.match(before)) and ((not after) or not WORD.match(after)):
            return True
        pos = nt.find(ne, pos + 1)
    return False

def main():
    n_show = int(sys.argv[1]) if len(sys.argv) > 1 else 5
    gdb = os.path.join(DB_DIR, "graph_rag.db")
    conn = sqlite3.connect(gdb)
    conn.row_factory = sqlite3.Row
    total = conn.execute("SELECT COUNT(*) FROM relations").fetchone()[0]
    nulls = conn.execute("SELECT COUNT(*) FROM relations WHERE source_chunk IS NULL").fetchone()[0]
    print(f"Всего рёбер: {total}, с NULL source_chunk: {nulls} ({100.0*nulls/max(total,1):.1f}%)")
    rows = conn.execute(
        "SELECT source, relation, target, chunk_id FROM relations "
        "WHERE source_chunk IS NULL ORDER BY rowid DESC LIMIT ?", (n_show * 3,)).fetchall()
    if not rows:
        print("NULL-якорей нет — промахи не накапливаются в NULL (или база пуста).")
        return
    client = chromadb.PersistentClient(path=DB_DIR)
    coll = client.get_or_create_collection(name="smk_docs", embedding_function=None)
    shown = 0
    for r in rows:
        if shown >= n_show:
            break
        m = re.match(r"^(.+)_chunk_(\d+)$", r["chunk_id"] or "")
        if not m:
            print(f"\n[скип GraphML-ребро] {r['source']} -> {r['target']}")
            continue
        fp, idx = m.group(1), int(m.group(2))
        win_ids = [f"{fp}_chunk_{k}" for k in range(idx, idx + 6)]
        try:
            recs = coll.get(ids=win_ids, include=["documents"])
        except Exception as e:
            print(f"\n[скип: smk_docs недоступен: {e}]")
            return
        docs = {i: d for i, d in zip(recs.get("ids", []), recs.get("documents", []))}
        texts = [(i, d) for i, d in docs.items() if d]
        if not texts:
            print(f"\n[скип: чанки окна не найдены в smk_docs] {r['chunk_id']}")
            continue
        shown += 1
        ents = [r["source"], r["target"]]
        norms = [norm_entity(e) for e in ents if e and e.strip()]
        found = {n: any(entity_in_text(n, norm_entity(t)) for _, t in texts) for n in norms}
        print(f"\n=== Окно {r['chunk_id']} (файл: {os.path.basename(fp)}) ===")
        print(f"Триплет: «{r['source']}» -> [{r['relation']}] -> «{r['target']}»")
        for n, f in found.items():
            print(f"  сущность «{n}»: {'НАЙДЕНА в тексте окна' if f else 'НЕ найдена'}")
        for i, t in sorted(texts, key=lambda p: int(p[0].rsplit('_chunk_', 1)[1])):
            print(f"  --- {os.path.basename(i)} ({len(t)} симв.) ---")
            print("  " + t[:350].replace("\n", " ") + ("…" if len(t) > 350 else ""))
    conn.close()

if __name__ == "__main__":
    main()
