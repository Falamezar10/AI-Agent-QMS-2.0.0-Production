# -*- coding: utf-8 -*-
"""Этап 2 плана D-2026-09-25-03: одноразовый ручной сброс рёбер + очереди
(processed_chunks) по ударным файлам с высокой долей NULL-якорей.

Пере-извлечение выполнит штатный паук в новых раундах НОВЫМ 5-полевым промптом
(FIFO, бэкофф, poison-guard). Дубликатов нет: рёбра удалены, узлы дедупятся по
векторному кэшу (node_embeddings / smk_graph_nodes) — пересбор узлов НЕ нужен.

Преусловия (все обязательны):
  1) запуск ТОЛЬКО при закрытом приложении (guard BEGIN IMMEDIATE, busy_timeout=3000);
  2) локальная база свежая (недавний admin-sync);
  3) по умолчанию --dry-run: только счётчики, ничего не удаляется;
  4) запись подтверждается флагом --yes.

Аргументы файлов: полные пути-префиксы, как в chunk_id (путь тот, что в рёбрах ИМЕННО
текущей базы; chunk_id машинно-зависимый), ИЛИ --top N (файлы с максимальной NULL-долей
по текущей базе — рекомендуется).
Удаление БЕЗ LIKE (underscore в пути = любой символ шаблона LIKE): фильтр
chunk_id.startswith(fp + "_chunk_") в Python — эталон расчёта _file_chunks_info автолечения.
Лимит: максимум --max-files файлов за прогон (защита от шторма).
smk_graph_nodes / node_embeddings / .cache не трогаются.

Примеры:
  py -3 scripts/reextract_null_edges.py --top 10
  py -3 scripts/reextract_null_edges.py --top 10 --yes
  py -3 scripts/reextract_null_edges.py "D:\\Docs\\СП 52.13330.2016.rtf" --yes
"""
import argparse
import os
import re
import sqlite3
import sys

LA = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
PROFILES = ["SMK_Agent_6c4d08", "SMK_Agent_948257", "SMK_Agent_ad218e"]


def find_graph_db():
    """Ищет graph_rag.db по известным профилям (первый найденный)."""
    for name in PROFILES:
        gr = os.path.join(LA, name, "local_vector_db", "graph_rag.db")
        if os.path.exists(gr):
            return gr
    return None


def chunk_file_prefix(chunk_id):
    """'{файл}_chunk_{N}' -> файл; эталон — расчёт в _file_chunks_info автолечения main.py."""
    m = re.match(r"^(.+)_chunk_(\d+)$", chunk_id or "")
    return m.group(1) if m else None


def main():
    ap = argparse.ArgumentParser(
        description="Сброс рёбер+очереди по ударным файлам (этап 2 D-2026-09-25-03)")
    ap.add_argument("files", nargs="*", help="полные пути-префиксы файлов (как в chunk_id текущей базы)")
    ap.add_argument("--top", type=int, metavar="N",
                    help="файлы с максимальной NULL-долей по текущей базе (рекомендуется)")
    ap.add_argument("--max-files", type=int, default=20, metavar="N",
                    help="лимит файлов за прогон (по умолчанию 20)")
    ap.add_argument("--yes", action="store_true", help="реально удалить (иначе только dry-run)")
    args = ap.parse_args()

    if not args.top and not args.files:
        ap.error("укажите --top N или список путей-префиксов")

    gr_db = find_graph_db()
    if not gr_db:
        print("ОШИБКА: graph_rag.db не найден в известных профилях:", PROFILES)
        sys.exit(1)
    print(f"База: {gr_db}")
    print(f"Режим: {'ЗАПИСЬ (--yes)' if args.yes else 'DRY-RUN (без изменений; запись подтвердите --yes)'}")

    # Преусловие (1): приложение закрыто — эксклюзивная блокировка записи
    conn = sqlite3.connect(gr_db, timeout=3)
    try:
        conn.execute("PRAGMA busy_timeout=3000")
        try:
            conn.execute("BEGIN IMMEDIATE")
        except sqlite3.OperationalError as e:
            print(f"ОШИБКА: база занята ({e}) — закройте приложение SMK Agent и повторите.")
            sys.exit(2)

        cols = {r[1] for r in conn.execute("PRAGMA table_info(relations)")}
        if "source_verbatim" not in cols or "source_chunk" not in cols:
            print("ОШИБКА: в relations нет колонок источника (source_chunk/source_verbatim) — "
                  "миграции init_graph_db не выполнены (старая база/старый exe). Запуск невозможен.")
            sys.exit(3)
        rows = conn.execute("SELECT id, chunk_id, source_chunk FROM relations").fetchall()
        proc_ids_all = [r[0] for r in conn.execute("SELECT chunk_id FROM processed_chunks")]

        # Группировка по файлу: фильтр в Python, БЕЗ LIKE (underscore = любой символ в шаблоне)
        per_file = {}
        for rid, cid, sc in rows:
            fp = chunk_file_prefix(cid)
            if fp is None:
                continue  # GraphML-рёбра (chunk_id = имя схемы): не сбрасываются
            d = per_file.setdefault(fp, {"null_n": 0, "all_n": 0})
            d["all_n"] += 1
            if sc is None:
                d["null_n"] += 1
        proc_by_file = {}
        for pid in proc_ids_all:
            fp = chunk_file_prefix(pid)
            if fp is not None:
                proc_by_file.setdefault(fp, []).append(pid)

        if not per_file:
            print("Рёбер со строковыми chunk_id нет — удалять нечего.")
            return

        # Выбор целевых файлов: заданные префиксы или топ по NULL-доле (при равенстве — по числу NULL)
        if args.top:
            scored = sorted(per_file.items(),
                            key=lambda kv: (-kv[1]["null_n"] / max(1, kv[1]["all_n"]),
                                            -kv[1]["null_n"], kv[0]))[:args.top]
            targets = [fp for fp, _ in scored]
        else:
            known = set(per_file)
            targets = []
            for fp in args.files:
                if fp in known:
                    targets.append(fp)
                else:
                    print(f"ПРЕДУПРЕЖДЕНИЕ: префикс не найден в рёбрах текущей базы (пропущен): {fp}")
        if not targets:
            print("Не осталось целевых файлов — выход.")
            return
        if len(targets) > args.max_files:
            print(f"Лимит --max-files={args.max_files} превышен ({len(targets)} файлов) — "
                  f"обрабатываются первые {args.max_files}; остальное — следующим прогоном.")
            targets = targets[:args.max_files]

        # План по каждому файлу (собирается ОДИН раз и переиспользуется в записи —
        # фильтр план/запись не расходятся): удалить ВСЕ рёбра файла (как уровень 2
        # автолечения — в т.ч. случайно валидные) и сбросить в очередь ВСЕ его чанки
        plan = {}
        total_edges = total_chunks = 0
        print()
        for fp in targets:
            d = per_file[fp]
            edges = [(rid,) for rid, cid, _sc in rows if chunk_file_prefix(cid) == fp]
            chunks = proc_by_file.get(fp, [])
            plan[fp] = (edges, chunks)
            total_edges += len(edges)
            total_chunks += len(chunks)
            print(f"Файл: {fp}")
            print(f"  рёбер всего: {d['all_n']}, из них NULL-якорь: {d['null_n']} "
                  f"({100 * d['null_n'] / max(1, d['all_n']):.1f}%)")
            print(f"  план: удалить рёбер={len(edges)}, чанков в очередь={len(chunks)}")

        if args.yes:
            for fp, (edges, chunks) in plan.items():
                conn.executemany("DELETE FROM relations WHERE id = ?", edges)
                conn.executemany("DELETE FROM processed_chunks WHERE chunk_id = ?",
                                 [(c,) for c in chunks])
            conn.commit()
            print()
            print(f"Записано: удалено рёбер={total_edges}, чанков в очередь={total_chunks} "
                  f"(файлов: {len(targets)}). Паук до-обработает очередь штатными раундами; "
                  "контроль: доля NULL по этим файлам должна падать "
                  "(scripts/diag_graph_matching.py).")
        else:
            conn.rollback()
            print()
            print(f"Dry-run: ничего не удалено (файлов: {len(targets)}, "
                  f"рёбер к удалению: {total_edges}, чанков к очереди: {total_chunks}). "
                  "Для записи добавьте --yes.")
    finally:
        try:
            conn.rollback()  # no-op после commit; страховка при выходе по ошибке
        except Exception:
            pass
        conn.close()


if __name__ == "__main__":
    main()
