# Хэш профиля для произвольного base-пути: python scripts\hash_base.py "<путь>"
import sys
import hashlib
print("SMK_Agent_" + hashlib.md5(sys.argv[1].encode("utf-8")).hexdigest()[:6])
