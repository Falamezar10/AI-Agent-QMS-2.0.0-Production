# Доливка 129 graphml-меток, которых нет в кэше node_embeddings (продолжение rebuild)
import os
import sys
import re
import sqlite3
import hashlib

PROFILE = r"C:\Users\plaksunov\AppData\Local\SMK_Agent_6c4d08"
VDB = os.path.join(PROFILE, "local_vector_db")
GR = os.path.join(VDB, "graph_rag.db")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import chromadb
client = chromadb.PersistentClient(path=VDB)
coll = client.get_or_create_collection(name="smk_graph_nodes")

import main as m
m.get_local_path = lambda: PROFILE
m.get_base_path = lambda: r"Z:\Производство\Проекты производства\1603 СМК Полюс-СТ\AI Agent SMQ Gr"
ef = m.get_cloud_ef()
print("ef готов")


def norm(s):
    return re.sub(r"\s+", " ", (s or "")).strip().lower()


conn = sqlite3.connect(GR, timeout=30)
need = sorted({r[0].strip() for r in conn.execute(
    "SELECT source FROM relations UNION SELECT target FROM relations")
    if r[0] and r[0].strip()})
have = {r[0] for r in conn.execute("SELECT node_id FROM node_embeddings")}
missing = [s for s in need
           if "gn_" + hashlib.md5(norm(s).encode("utf-8")).hexdigest() not in have]
conn.close()
print("Метки без вектора:", len(missing))

done = 0
for i in range(0, len(missing), 64):
    part = missing[i:i + 64]
    vs = [list(map(float, v)) for v in ef(part)]
    vid = ["gn_" + hashlib.md5(norm(s).encode("utf-8")).hexdigest() for s in part]
    coll.upsert(ids=vid, embeddings=vs, documents=part,
                metadatas=[{"entity": d} for d in part])
    done = min(i + 64, len(missing))
    print(f"долито: {done}/{len(missing)}")

print("Итог: узлов в коллекции:", coll.count())
