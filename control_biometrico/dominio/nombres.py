import re
from unicodedata import normalize
def clave(v):
    s=normalize("NFKD",str(v or "").strip().upper()).encode("ascii","ignore").decode()
    return " ".join(re.findall(r"[A-Z0-9]+",s))
