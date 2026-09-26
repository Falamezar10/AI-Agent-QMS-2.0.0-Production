# Одноразовый ремонт: пересборка коллекции smk_graph_nodes из graph_rag.db.
# Приложение должно быть ЗАКРЫТО. Запуск: python scripts\rebuild_graph_nodes.py "<профиль>"
import os
import sys
import re
import array
import sqlite3
import hashlib

PROFILE = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "SMK_Agent_6c4d08")
VDB = os.path.join(PROFILE, "local_vector_db")
GR = os.path.join(VDB, "graph_rag.db")
BATCH = 128
print("Профиль:", PROFILE)
print("Векторная база:", VDB)

if not os.path.exists(GR):
    raise SystemExit("graph_rag.db не найден: " + GR)

import chromadb
print("chromadb:", chromadb.__version__)

client = chromadb.PersistentClient(path=VDB)
try:
    client.delete_collection("smk_graph_nodes")
    print("Старая коллекция удалена (вместе с испорченным HNSW-сегментом)")
except Exception as e:
    print("Коллекции не было (ок):", e)

# Облачный embedding нужен только для доливки graphml-узлов, которых нет в кэше; сбой не критичен
ef = None
try:
    import main as m
    m.get_local_path = lambda: PROFILE
    m.get_base_path = lambda: PROFILE
    ef = m.get_cloud_ef()
    print("Облачный embedding подключён")
except Exception as e:
    print("Облачный embedding недоступен (не критично):", e)

created = False
for meta in ({"hnsw:num_threads": 1}, None):
    try:
        coll = client.get_or_create_collection(
            name="smk_graph_nodes", embedding_function=ef, metadata=meta)
        created = True
        print("Коллекция создана, metadata:", meta)
        break
    except Exception as e:
        print("Создание с metadata", meta, "не удалось:", e)
if not created:
    raise SystemExit("Не удалось создать коллекцию")

conn = sqlite3.connect(GR, timeout=30)
rows = conn.execute("SELECT node_id, document, embedding FROM node_embeddings").fetchall()
print("Строк в кэше node_embeddings:", len(rows))


def norm(s):
    return re.sub(r"\s+", " ", (s or "")).strip().lower()


ids, docs, vecs = [], [], []
written_ids = set()
dim = None
skipped = 0
total = 0


def flush():
    global ids, docs, vecs, total
    if not ids:
        return
    coll.upsert(ids=ids, embeddings=vecs, documents=docs,
                metadatas=[{"entity": d} for d in docs])
    total += len(ids)
    print(f"  upsert: {total}/{len(rows)}")
    ids, docs, vecs = [], [], []


for node_id, document, blob in rows:
    v = array.array("f")
    v.frombytes(blob)
    if dim is None:
        dim = len(v)
        print("Размерность векторов узлов:", dim)
    if len(v) != dim:
        skipped += 1
        continue
    ids.append(node_id)
    docs.append(document)
    vecs.append(v.tolist())
    written_ids.add(node_id)
    if len(ids) >= BATCH:
        flush()
flush()
print("Из кэша записано узлов:", total, "пропущено по размерности:", skipped)

# Метки из рёбер, отсутствующие в кэше (узлы graphml) — доливаем через ef
need = sorted({r[0].strip() for r in conn.execute(
    "SELECT source FROM relations UNION SELECT target FROM relations")
    if r[0] and r[0].strip()})
missing = [s for s in need
           if "gn_" + hashlib.md5(norm(s).encode("utf-8")).hexdigest() not in written_ids]
print("Метки из рёбер без вектора:", len(missing))
if missing:
    if ef is None:
        print("ef недоступен — пропускаю (на паука не влияет)")
    else:
        try:
            for i in range(0, len(missing), 64):
                part = missing[i:i + 64]
                vs = [list(map(float, v)) for v in ef(part)]
                vid = ["gn_" + hashlib.md5(norm(s).encode("utf-8")).hexdigest() for s in part]
                coll.upsert(ids=vid, embeddings=vs, documents=part,
                            metadatas=[{"entity": d} for d in part])
                print(f"  долито из ef: {min(i + 64, len(missing))}/{len(missing)}")
        except Exception as e:
            print("Доливка через ef не удалась (не критично):", e)

# Пробный запрос первым вектором из кэша — проверка живости нового HNSW
try:
    pr = array.array("f")
    pr.frombytes(rows[0][2])
    res = coll.query(query_embeddings=[pr.tolist()], n_results=1, include=["distances"])
    print("Пробный запрос ok:", res["ids"][0])
except Exception as e:
    print("Пробный запрос не удался:", e)

conn.close()
print("Итог: узлов в коллекции:", coll.count())
