# -*- coding: utf-8 -*-
"""Soak-тест chromadb на КОПИИ прод-базы (план D-2026-09-25-02, шаг 9).

Цель: гейт перед деплоем обновлённой chromadb. Гоняет real-продовые вызовы
(count / query_texts / upsert / get / where_document) на копии smk_vector_db,
при фоновой daemon-нагрузке (GIL-timing класс крашей chromadb_rust_bindings —
issue #6979: SIGSEGV случается, когда главный поток блокирован в syscall).

ЧТО ЭТО ДЕЛАЕТ:
  - открывает PersistentClient ТОЛЬКО на --db (копия; живая реплика не трогается);
  - read-only работа с smk_docs и smk_graph_nodes (count + query_texts);
  - запись — в одноразовую scratch-коллекцию `soak_scratch_<pid>` (батчи upsert,
    потом delete колекции) — smk_docs / smk_graph_nodes / temp_chat_memory НЕ пишутся;
  - эмбеддинги реальные: get_cloud_ef() из main.py (qwen через OpenRouter),
    запускать ИЗ КОРНЯ ПРОЕКТА (ключ secrets.vault репо-корня, прокси из профиля
    локальных настроек этой машины);
  - faulthandler пишет стек всех потоков в crash.log рядом со скриптом при нативном AV;
  - код возврата: 0 — гейт пройден, 1 — ошибки Python. Нативный краш (c0000005) = смерть
    процесса БЕЗ кода 1/без трейсбека → проверять crash.log и Event Viewer (EventID 1000).

ПРИМЕРЫ:
  py -3 scripts/soak_chromadb.py --db "backup_smk_vector_db_2026-09-26\\server_smk_vector_db" --mode read --rounds 5
  py -3 scripts/soak_chromadb.py --db "<копия>" --mode write --rounds 20
  py -3 scripts/soak_chromadb.py --db "<копия>" --mode counts          (эталон до миграции)
  py -3 scripts/soak_chromadb.py --db "<копия>" --list-collections
"""
import argparse
import faulthandler
import os
import random
import shutil
import sys
import tempfile
import threading
import time

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
CRASH_LOG = os.path.join(SCRIPT_DIR, "soak_crash.log")

_scratch_marker = "soak_scratch_"

SMK_QUERIES = [
    "документированная информация СМК",
    "процесс управления документированной информацией",
    "внутренние аудиты системы менеджмента качества",
    "корректирующие действия несоответствия",
    "риски и возможности организации",
    "цикл PDCA планирование поддержка мероприятия",
    "записи о подтверждении прослеживаемости измерений",
    "управление несоответствующей продукцией изоляция",
    "мониторинг удовлетворённости потребителейClaims",
    "компетентность персонала испытательного стенда",
    "обеспечение сохранности поверки средствами измерений",
    "анализ и оценка деятельности со стороны руководства",
    "приобретение внешних поставщиков оценка закупки",
    "цели в области качества показатели процессов",
    "рассмотрение жалоб претензий потребителей",
    "идентификация и статус инспекций контроля",
    "управление изменениями документации актуализация",
    "обеспечение сохранности продукции при хранении",
    "планирование разработка проектирование и разработка",
    "менеджмент риска в СМК",
]

# Цель for pipe: thumb-size verification  of the data after migration
ETALON_COLLECTIONS = ["smk_docs", "smk_graph_nodes", "temp_chat_memory"]


def log(msg):
    print(time.strftime("%H:%M:%S") + " " + str(msg), flush=True)


def make_scratch_collection(client, ef, tag):
    name = f"{_scratch_marker}{tag}_{os.getpid()}_{random.randint(1000, 9999)}"
    coll = client.get_or_create_collection(
        name=name, embedding_function=ef, metadata={"hnsw:batch_size": 100000, "hnsw:sync_threshold": 1000}
    )
    return coll


def clear_scratch(client):
    """Удаляет до последней scratch-коллекции на этой копии (самолечение от прошлых прогонов)."""
    try:
        names = [c.strip() for c in client.list_collections()]
        real_names = []
        for c in client.list_collections():
            nm = getattr(c, "name", None)
            real_names.append(nm)
        for nm in real_names:
            if nm and nm.startswith(_scratch_marker):
                try:
                    client.delete_collection(nm)
                    log(f"scratch-очистка: удалена {nm}")
                except Exception as e:
                    log(f"scratch-очистка FAIL {nm}: {e}")
    except Exception as e:
        log(f"scratch-очистка: ошибка списка коллекций: {e}")


def do_counts(client, ef=None):
    report = {}
    for cname in ETALON_COLLECTIONS:
        try:
            kw = {"embedding_function": ef} if ef else {}
            coll = client.get_collection(name=cname, **kw)
            n = coll.count()
            report[cname] = n
            log(f"count({cname}) = {n}")
        except Exception as e:
            report[cname] = f"ERROR: {e}"
            log(f"count({cname}) FAILED: {e}")
    return report


def run_read_round(client, ef, round_idx, queries):
    # ВАЖНО: ef передаётся в get_collection — без него chromadb для query_texts
    # подставит дефолтную 384-dim MiniLM, что не соответствует продовой qwen (4096).
    smk = client.get_collection(name="smk_docs", embedding_function=ef)
    q = queries[round_idx % len(queries)]
    t0 = time.time()
    res = smk.query(query_texts=[q], n_results=5)
    docs = res.get("documents", [[]])[0]
    log(f"round {round_idx} [smk_docs.query] q={q[:40]!r} -> {len(docs)} документов, {time.time()-t0:.2f} с")

    try:
        nodes = client.get_collection(name="smk_graph_nodes", embedding_function=ef)
        t0 = time.time()
        gres = nodes.query(query_texts=[q], n_results=5)
        gnames = gres.get("documents", [[]])[0]
        log(f"round {round_idx} [graph.query] -> {len(gnames)} узлов, {time.time()-t0:.2f} с")
    except Exception as e:
        log(f"round {round_idx} [graph.query] FAILED: {e}")

    t0 = time.time()
    got = smk.get(limit=10, include=["documents", "metadatas"])
    log(f"round {round_idx} [get(limit=10)] -> {len(got.get('ids') or [])} id, {time.time()-t0:.2f} с")

    doc_marker = (got.get("documents") or [""])[0][:20].strip()
    if doc_marker:
        r2 = smk.get(where_document={"$contains": doc_marker}, limit=5)
        log(f"round {round_idx} [where_document $contains] -> {len(r2.get('ids') or [])}")


def run_write_round(client, ef, round_idx):
    scratch = make_scratch_collection(client, ef, f"r{round_idx}")
    try:
        batch_total = 0
        t0 = time.time()
        # 3 батча по 100 документов, как паук (имитация upsert-нагрузки)
        for b in range(3):
            ids = [f"soak_{round_idx}_{b}_{i}" for i in range(100)]
            docs = [
                f"Soak-документ round{round_idx} batch{b} item{i}: "
                f"документированная информация СМК, аудит процесса_pdca, риск {random.randint(0, 999)}"
                for i in range(100)
            ]
            metas = [{"round": round_idx, "batch": b, "file_path": f"soak://batch{b}/item{i}"} for i in range(100)]
            scratch.upsert(ids=ids, documents=docs, metadatas=metas)
            batch_total += 100
        log(f"round {round_idx} [write] upsert {batch_total} доков ({time.time()-t0:.2f} с)")

        t0 = time.time()
        res = scratch.query(query_texts=["риск процесса аудит СМК"], n_results=5)
        log(f"round {round_idx} [write] query scratch -> {len(res.get('documents', [[]])[0])}, {time.time()-t0:.2f} с")

        t0 = time.time()
        for i in range(0, 300, 50):
            scratch.delete(ids=[f"soak_{round_idx}_0_{k}" for k in range(i, i + 50)])
        log(f"round {round_idx} [write] delete 150 ids ({time.time()-t0:.2f} с)")

        n = scratch.count()
        log(f"round {round_idx} [write] scratch.count после удаления = {n}")
    finally:
        try:
            client.delete_collection(scratch.name)
        except Exception as e:
            log(f"round {round_idx} [write] удаление scratch FAILED: {e}")


def main():
    ap = argparse.ArgumentParser(description="Soak-тест chromadb на копии прод-базы (D-2026-09-25-02)")
    ap.add_argument("--db", required=True, help="путь к КОПИИ smk_vector_db (папка с chroma.sqlite3)")
    ap.add_argument("--mode", choices=["read", "write", "counts"], default="read",
                    help="read = count+query; write = +upsert/delete scratch; counts = только эталон")
    ap.add_argument("--rounds", type=int, default=10, help="сколько раундов гонять")
    ap.add_argument("--sleep", type=float, default=0.5, help="пауза между раундами, с")
    ap.add_argument("--list-collections", action="store_true", help="список коллекций и выход")
    args = ap.parse_args()

    db_path = os.path.abspath(args.db)
    if not os.path.exists(os.path.join(db_path, "chroma.sqlite3")):
        print(f"FATAL: {db_path} не содержит chroma.sqlite3", flush=True)
        return 2

    # faulthandler: стеки всех потоков при нативном AV -> soak_crash.log рядом со скриптом
    try:
        fh = open(CRASH_LOG, "a", encoding="utf-8")
        faulthandler.enable(fh)
        faulthandler.dump_traceback_later(120, repeat=True, file=fh)
        log(f"faulthandler -> {CRASH_LOG}")
    except Exception as e:
        log(f"faulthandler не включён: {e}")

    log(f"chromadb import: {__import__('chromadb')}")
    log(f"chromadb version: {getattr(__import__('chromadb'), '__version__', 'UNKNOWN')}")
    log(f"db = {db_path}")
    log(f"mode = {args.mode}, rounds = {args.rounds}")

    # Фоновая daemon-активность: syscall-блокировка главного потока — условие repro #6979.
    stop_flag = threading.Event()

    def background_banner():
        while not stop_flag.is_set():
            try:
                time.sleep(1.5)
            except Exception:
                pass

    bg = threading.Thread(target=background_banner, daemon=True)
    bg.start()

    import chromadb

    log(f"chromadb {chromadb.__version__}, PersistentClient на {db_path}")
    client = chromadb.PersistentClient(path=db_path)

    if args.list_collections:
        for c in client.list_collections():
            log(f"collection: {c.name}")
        stop_flag.set()
        return 0

    # Эмбеддинги — реальные продовые (get_cloud_ef из main.py; запуск из корня проекта)
    sys.path.insert(0, PROJECT_ROOT)
    try:
        import main as smk_main
        ef = smk_main.get_cloud_ef()
        log(f"EF = {type(ef).__name__} (модель {smk_main.load_global_settings().get('embedding_model', '?')})")
        # прогрев: один реальный вызов, чтобы сетевые проблемы упали здесь, а не посреди раунда
        warm = ef(["soak warmup проверка связи"])
        log(f"EF warmup OK, dims={len(warm[0]) if warm else 0}")
    except Exception as e:
        log(f"FATAL: EF не строится / сеть недоступна: {e}")
        print("ГЕП НЕ ПРОЙДЕН: soak без реальных эмбеддингов невалиден (см. риски плана).", flush=True)
        stop_flag.set()
        return 1

    counts = do_counts(client, ef=ef)
    log("---")

    if args.mode == "counts":
        stop_flag.set()
        print("COUNTS-REPORT " + repr(counts), flush=True)
        return 0

    ok_rounds = 0
    t_start = time.time()
    try:
        for i in range(args.rounds):
            run_read_round(client, ef, i, SMK_QUERIES)
            if args.mode == "write":
                run_write_round(client, ef, i)
            time.sleep(args.sleep)
            if (i + 1) % 10 == 0:
                log(f"=== прогресс {i+1}/{args.rounds} раундов, {time.time()-t_start:.1f} с ===")
    except KeyboardInterrupt:
        log("Прервано пользователем")
        stop_flag.set()
        return 130
    except Exception as e:
        log(f"FATAL в цикле: {e.__class__.__name__}: {e}")
        stop_flag.set()
        return 1

    total_t = time.time() - t_start
    log(f"ЗАВЕРШЕНО: {args.rounds} раундов за {total_t:.1f} с, mode={args.mode}")
    counts_end = do_counts(client, ef=ef)
    same = all(str(counts[c]) == str(counts_end[c]) for c in ETALON_COLLECTIONS)
    log(f"count до/после совпадают: {same}")
    stop_flag.set()

    if not same:
        print("FATAL: count изменился в течение soak — данные дрогнули", flush=True)
        return 1
    print("RESULTS_OK", flush=True)
    return 0


if __name__ == "__main__":
    rc = 83
    try:
        rc = main()
    except SystemExit as e:
        rc = e.code if isinstance(e.code, int) else 0
    except Exception:
        import traceback
        traceback.print_exc()
        rc = 1
    print(f"exit code = {rc}", flush=True)
    sys.exit(rc)
