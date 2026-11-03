# -*- coding: utf-8 -*-
"""تشفير data.json بكلمة سر (PBKDF2 + AES-GCM) ليفكّه المتصفح بنفس الكلمة."""
import json, base64, hashlib, hmac
ITER = 600000
def encrypt(obj, password, salt_label=b"hom-dashboard-salt"):
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    salt = hashlib.sha256(salt_label).digest()[:16]      # ثابت: الملف ميتغيرش إلا لو البيانات اتغيرت
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(password.encode("utf-8"))
    plain = json.dumps(obj, ensure_ascii=False, sort_keys=True).encode("utf-8")
    iv = hmac.new(key, plain, hashlib.sha256).digest()[:12]         # IV مشتق من المحتوى
    ct = AESGCM(key).encrypt(iv, plain, None)
    b = lambda x: base64.b64encode(x).decode()
    return dict(v=1, iter=ITER, salt=b(salt), iv=b(iv), ct=b(ct))
