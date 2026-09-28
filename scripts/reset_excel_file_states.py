# -*- coding: utf-8 -*-
"""Одноразовый точный пересинк Excel-файлов для D-2026-09-25-01 (скрытые листы).

Правка extract_text_from_excel_for_rag (main.py) не меняет mtime файлов, поэтому
штатный синк сам не переиндексирует Excel — записи file_states.json сбрасываются
скриптом, и следующий синк приложения пере-эмбеддит затронутые файлы
(delete by file_path -> upsert, дублей нет) с self-heal графа по префиксам.

РЕЖИМЫ:
  по умолчанию — ТОЧНЫЙ: pre-scan каждого *.xlsx из file_states.json через
  openpyxl (read_only), кандидат — файл, у которого есть хоть один лист
  sheet_state != "visible"; ошибка чтения (пароль/битый) — тоже кандидат (безопасно);
  --all — запасной режим: сбрасываются ВСЕ Excel-записи (*.xlsx/*.xls).

ОТЧЁТ (и в dry-run, и перед записью): кандидаты с перечнем листов и состояний,
счётчики; справочно — .xls-записи в file_states (в RAG не индексируются) и
подсчёт *.xls в indexed_folders из global_settings.json; stale-записи (файл не
существует) — только в отчёт, приложение чистит их само как deleted.

ПРЕУСЛОВИЯ:
  1) запуск ТОЛЬКО при закрытых ВСЕХ экземплярах приложения (guard BEGIN IMMEDIATE
     на КАЖДОЙ найденной graph_rag.db из профилей SMK_Agent_* — прод-агентов
     НЕСКОЛЬКО на разных базах (AI Agent SMQ Gr, Ai Agent SMQ GOZ), у каждого свой
     профиль; занята хоть одна база -> abort);
  2) деплой НОВОГО exe должен быть ВЫПОЛНЕН ДО сброса (иначе автосинк старым exe
     вернёт записи со свежими mtime — понадобится повторный сброс);
  3) по умолчанию dry-run; запись подтверждается флагом --yes.

ЗАПИСЬ (--yes): backup file_states.json.bak_<timestamp>; атомарная запись
(tmp + os.replace); формат как save_file_states (utf-8, ensure_ascii=False,
indent=2). Идемпотентен: после успешного сброса повторный --yes — no-op.
Скрипт автономен (main.py не импортирует) и НЕ трогает smk_graph_nodes /
node_embeddings / graph_rag.db / .cache.

ВАЖНО: у каждого прод-агента СВОЙ base (свой file_states.json) — сброс выполняется
ОТДЕЛЬНО для каждой базы (--base для каждой папки).

Примеры:
  py -3 scripts/reset_excel_file_states.py --base "Z:\\...\\AI Agent SMQ Gr"
  py -3 scripts/reset_excel_file_states.py --base "Z:\\...\\Ai Agent SMQ GOZ" --yes
  py -3 scripts/reset_excel_file_states.py --all --yes
"""
import argparse
import os
import sqlite3
import sys
import time
import json

import openpyxl

LA = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
PROFILES = ["SMK_Agent_6c4d08", "SMK_Agent_948257", "SMK_Agent_ad218e"]

EXCEL_EXTS = (".xlsx", ".xls")
# Точно пресканируемые расширения: только .xlsx (openpyxl .xls не читает, а
# .xls-записи в состояниях не бывают — см. отчёт ниже); .xls сбрасывается в --all.
PRESCAN_EXTS = (".xlsx",)


def check_app_closed():
    """Guard «приложение закрыто»: BEGIN IMMEDIATE на КАЖДОЙ найденной graph_rag.db.

    Приложение хэширует base-путь в имя профиля SMK_Agent_* (несколько прод-агентов
    = несколько профилей на одной машине), поэтому проверяются ВСЕ найденные базы:
    занята хоть одна — abort (один из агентов запущен).
    Может быть отключён флагом --skip-guard (для синтетических тестов на
    временной базе, где профилей не существует).
    """
    found = [(name, os.path.join(LA, name, "local_vector_db", "graph_rag.db"))
             for name in PROFILES
             if os.path.exists(os.path.join(LA, name, "local_vector_db", "graph_rag.db"))]
    if not found:
        print("ПРЕДУПРЕЖДЕНИЕ: graph_rag.db не найден ни в одном из профилей "
              f"{PROFILES} — guard пропущен (приложение не проверялось).")
        return
    for name, gr_db in found:
        print(f"Локальная база (guard): {gr_db}")
        conn = sqlite3.connect(gr_db, timeout=3)
        try:
            conn.execute("PRAGMA busy_timeout=3000")
            try:
                conn.execute("BEGIN IMMEDIATE")
            except sqlite3.OperationalError as e:
                print(f"ОШИБКА: база профиля {name} занята ({e}) — закройте ВСЕ экземпляры "
                      "SMK Agent и повторите.")
                sys.exit(2)
        finally:
            try:
                conn.rollback()
            except Exception:
                pass
            conn.close()
    print(f"Guard пройден: все локальные графовые базы свободны ({len(found)} шт.).")


def load_json(path, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"ОШИБКА чтения {os.path.basename(path)}: {e}")
        sys.exit(3)


def prescan_xlsx(path):
    """Возвращает список [(имя листа, state), ...] или None при ошибке чтения."""
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        try:
            return [(ws.title, ws.sheet_state) for ws in wb.worksheets]
        finally:
            wb.close()
    except Exception:
        return None


def main():
    ap = argparse.ArgumentParser(
        description="Одноразовый точный сброс Excel-записей в file_states.json "
                    "(D-2026-09-25-01: скрытые листы)")
    ap.add_argument("--base", metavar="PATH",
                    help="папка с file_states.json/global_settings.json "
                         "(по умолчанию — корень репозитория)")
    ap.add_argument("--all", action="store_true",
                    help="сброс ВСЕХ Excel-записей (запасной режим; по умолчанию — "
                         "точный режим только по кандидатам pre-scan)")
    ap.add_argument("--skip-guard", action="store_true",
                    help="не проверять локальную graph_rag.db (для тестов)")
    ap.add_argument("--yes", action="store_true",
                    help="реально записать (иначе только dry-run)")
    args = ap.parse_args()

    base_dir = os.path.abspath(args.base) if args.base else os.getcwd()
    states_path = os.path.join(base_dir, "file_states.json")
    settings_path = os.path.join(base_dir, "global_settings.json")
    print(f"base: {base_dir}")

    states = load_json(states_path, {})
    if not states:
        print("file_states.json пуст или отсутствует — пересинк не нужен.")
        return
    excel_keys = [k for k in states if k.lower().endswith(EXCEL_EXTS)]
    xlsx_keys = [k for k in excel_keys if k.lower().endswith(".xlsx")]
    xls_keys = [k for k in excel_keys if k.lower().endswith(".xls")]
    stale = [k for k in excel_keys if not os.path.exists(k)]

    # --- отчётная часть ---
    if xls_keys:
        print(f"\nСПРАВКА: .xls-записей в file_states.json: {len(xls_keys)} "
              "(в RAG не индексируются; здесь только список):")
        for k in xls_keys:
            print(f"  .xls: {k}")

    indexed_folders = []
    settings = load_json(settings_path, {})
    if isinstance(settings, dict):
        raw_folders = settings.get("indexed_folders")
        if isinstance(raw_folders, list):
            indexed_folders = [
                f if os.path.isabs(f) else os.path.normpath(os.path.join(base_dir, f))
                for f in raw_folders
            ]
    xls_in_folders = [f for f in indexed_folders
                      if os.path.isdir(f) and any(
                          fn.lower().endswith(".xls")
                          for fn in os.listdir(f))]
    if xls_in_folders:
        print(f"\nСПРАВКА: папок indexed_folders с *.xls: {len(xls_in_folders)} "
              "(для будущего решения по идее .xls):")
        for f in xls_in_folders:
            print(f"  папка: {f}")

    if stale:
        print(f"\nSTALE-записи (файл не существует; синк чистит их сам как deleted): "
              f"{len(stale)}")
        for k in stale:
            print(f"  stale: {k}")

    # --- выбор кандидатов ---
    if args.all:
        reset_keys = sorted(set(xlsx_keys) | set(xls_keys) | set(stale))
        if reset_keys:
            print(f"\nРежим --all: сбрасываются ВСЕ Excel-записи ({len(reset_keys)}).")
    else:
        reset_keys = []
        print("\nPre-scan xlsx (точный режим):")
        for k in sorted(xlsx_keys):
            exists = os.path.exists(k)
            if not exists:
                continue  # stale — только в отчёт
            sheets = prescan_xlsx(k)
            if sheets is None:
                reset_keys.append(k)
                print(f"  КАНДИДАТ (ошибка чтения — безопасно): {k}")
            else:
                hidden = [(nm, st) for nm, st in sheets if st != "visible"]
                if hidden:
                    reset_keys.append(k)
                    lst = ", ".join(f"{nm}[{st}]" for nm, st in hidden)
                    print(f"  КАНДИДАТ: {k}\n    скрытые листы: {lst}")
        if not reset_keys:
            print("  кандидатов не найдено (пересинк не нужен).")

    if not reset_keys:
        print("\nИтог: ничего сбрасывать (сброс уже выполнялся или кандидатов нет).")
        return

    print(f"\nИтог: кандидатов на сброс: {len(reset_keys)}")
    if args.yes:
        if not args.skip_guard:
            check_app_closed()
        ts = time.strftime("%Y%m%d_%H-%M-%S")
        bak_path = states_path + f".bak_{ts}"
        try:
            with open(states_path, "r", encoding="utf-8") as f:
                raw = f.read()
            with open(bak_path, "w", encoding="utf-8") as f:
                f.write(raw)
            print(f"Backup: {bak_path}")
        except Exception as e:
            print(f"ОШИБКА: не удалось сделать backup: {e}")
            sys.exit(4)

        new_states = {k: v for k, v in states.items() if k not in set(reset_keys)}
        tmp_path = states_path + ".tmp"
        try:
            with open(tmp_path, "w", encoding="utf-8") as f:
                json.dump(new_states, f, ensure_ascii=False, indent=2)
            os.replace(tmp_path, states_path)
        except Exception as e:
            print(f"ОШИБКА: атомарная запись file_states.json не удалась: {e}")
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            sys.exit(4)
        print(f"Записано: удалено записей={len(reset_keys)} "
              f"(осталось всего записей: {len(new_states)})")
        print("Следующий шаг: открыть приложение — автосинк пере-эмбеддит затронутые "
              "файлы (deletes by file_path, self-heal графа) и выгрузит базу на "
              "сервер. Контроль: лог синка без «Индексатор: пропущен», "
              "count'ы Chroma/graph до/после.")
    else:
        print("Dry-run: ничего не записано. Для записи добавьте --yes.")


if __name__ == "__main__":
    main()
