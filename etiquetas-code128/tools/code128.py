"""Copia en Python de vba/ModCode128.bas (misma logica), para pruebas y para
precalcular los valores que se guardan en el libro."""

CODE_C, CODE_B, START_B, START_C, STOP = 99, 100, 104, 105, 106


def _digit_run(s, i):
    n = 0
    while i + n < len(s) and s[i + n] in "0123456789":
        n += 1
    return n


def _use_c(i, d, L):
    # i es 0-based aqui; en VBA es 1-based.
    if d < 2:
        return False
    if i == 0 and d == L:
        return d == 2 or d >= 4
    if i == 0 or i + d == L:
        return d >= 4
    return d >= 6


def values(text):
    """Valores de simbolo (start, datos, verificador, stop) o None si no se puede."""
    s = str(text).strip()
    L = len(s)
    if L == 0 or any(not 32 <= ord(ch) <= 126 for ch in s):
        return None
    v = []
    subset = None
    i = 0
    while i < L:
        d = _digit_run(s, i)
        if subset != "C" and _use_c(i, d, L):
            if d % 2 == 1 and i + d == L:
                if subset is None:
                    v.append(START_B)
                subset = "B"
                v.append(ord(s[i]) - 32)
                i += 1
                d -= 1
            v.append(START_C if subset is None else CODE_C)
            subset = "C"
            for _ in range(d // 2):
                v.append(int(s[i:i + 2]))
                i += 2
        else:
            if subset is None:
                v.append(START_B)
            elif subset == "C":
                v.append(CODE_B)
            subset = "B"
            v.append(ord(s[i]) - 32)
            i += 1
    check = (v[0] + sum(k * x for k, x in enumerate(v) if k > 0)) % 103
    return v + [check, STOP]


def to_char(value, zebu=False):
    if zebu:
        return chr(value + 32) if value < 95 else chr(value + 105)
    if value == 0:
        return chr(194)
    return chr(value + 32) if value < 95 else chr(value + 100)


def encode(text, mapeo=""):
    v = values(text)
    if v is None:
        return ""
    zebu = "ZEBU" in str(mapeo).upper()
    return "".join(to_char(x, zebu) for x in v)


def check_digit(text):
    v = values(text)
    return "" if v is None else v[-2]


def modules(text):
    """Ancho total en modulos (sin zonas de silencio)."""
    v = values(text)
    return 0 if v is None else (len(v) - 1) * 11 + 13

