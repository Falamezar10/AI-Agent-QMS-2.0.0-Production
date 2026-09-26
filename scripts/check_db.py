# Только чтение: диагностика состояния баз SMK Agent (chroma.sqlite3 + graph_rag.db)
import os
import sqlite3
import hashlib

LA = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
PROFILES = ["SMK_Agent_6c4d08", "SMK_Agent_948257", "SMK_Agent_ad218e"]

# Сопоставление: какой base-путь даёт какой хэш профиля
CANDIDATES = [
    r"C:\Users\plaksunov\AppData\Local\SMK_Agent_Client_AI Agent SMQ Gr",
    r"C:\Users\plaksunov\AppData\Local\SMK_Agent_Client_Ai Agent SMQ GOZ",
]
for c in CANDIDATES:
    print("base:", c, "-> профиль SMK_Agent_" + hashlib.md5(c.encode("utf-8")).hexdigest()[:6])


def nid(norm_key: str) -> str:
    return "gn_" + hashlib.md5(norm_key.encode("utf-8")).hexdigest()


for name in PROFILES:
    vdb = os.path.join(LA, name, "local_vector_db")
    if not os.path.isdir(vdb):
        print(f"--- {name}: нет local_vector_db")
        continue
    print("=" * 70)
    print(name)
    gr = os.path.join(vdb, "graph_rag.db")
    if os.path.exists(gr):
        conn = sqlite3.connect(f"file:{gr}?mode=ro", uri=True)
        try:
            ic = conn.execute("PRAGMA integrity_check").fetchone()[0]
            proc = conn.execute("SELECT COUNT(*) FROM processed_chunks").fetchone()[0]
            rel = conn.execute("SELECT COUNT(*) FROM relations").fetchone()[0]
            ne = conn.execute("SELECT COUNT(*) FROM node_embeddings").fetchone()[0]
            meta = dict(conn.execute("SELECT key, value FROM cache_meta").fetchall())
            print(f"  graph_rag.db: {os.path.getsize(gr)/1048576:.0f} МБ, integrity={ic}")
            print(f"  processed_chunks={proc}, relations={rel}, node_embeddings={ne}, cache_meta={meta}")
            need = {r[0] for r in conn.execute(
                "SELECT DISTINCT lower(source) FROM relations "
                "UNION SELECT DISTINCT lower(target) FROM relations")}
            have = {r[0] for r in conn.execute("SELECT node_id FROM node_embeddings")}
            missing = [s for s in need if nid(s) not in have]
            print(f"  сущностей в рёбрах: {len(need)}, из них без вектора в кэше: {len(missing)}")
            if missing:
                print("  примеры без вектора:", missing[:5])
        except Exception as e:
            print("  graph_rag.db: ошибка чтения:", e)
        finally:
            conn.close()
    ch = os.path.join(vdb, "chroma.sqlite3")
    if os.path.exists(ch):
        conn = sqlite3.connect(f"file:{ch}?mode=ro", uri=True)
        try:
            tables = [r[0] for r in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")]
            print(f"  chroma.sqlite3: {os.path.getsize(ch)/1048576:.0f} МБ")
            print(f"  таблицы: {tables}")
            if "collections" in tables:
                for cid, cname in conn.execute("SELECT id, name FROM collections").fetchall():
                    print(f"  коллекция: {cname} (id={cid[:8]}...)")
                    if "segments" in tables:
                        for sid, stype in conn.execute(
                                "SELECT id, scope FROM segments WHERE collection=?", (cid,)).fetchall():
                            print(f"    сегмент {sid[:8]}... scope={stype}")
        except Exception as e:
            print("  chroma: ошибка чтения:", e)
        finally:
            conn.close()
