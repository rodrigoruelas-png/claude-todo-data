"""Genera un vbaProject.bin (MS-OVBA) a partir de archivos .bas.

El proyecto se escribe solo con codigo fuente comprimido y con la version de
_VBA_PROJECT en 0xFFFF, asi Excel recompila el codigo al abrir el libro (no
hace falta p-code). Incluye modulos de documento para ThisWorkbook y cada hoja
para que coincidan con los codeName del libro.

Referencias: [MS-OVBA] secciones 2.3 (estructura), 2.4.1 (compresion),
2.4.3 (cifrado de CMG/DPB/GC) y [MS-CFB] (contenedor OLE).
"""

import os
import random
import struct
import uuid

CODEPAGE = 1252
LCID = 0x0409

BASE_WORKBOOK = "0{00020819-0000-0000-C000-000000000046}"
BASE_WORKSHEET = "0{00020820-0000-0000-C000-000000000046}"

REFERENCES = [
    ("stdole",
     "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
     "C:\\Windows\\System32\\stdole2.tlb#OLE Automation"),
    ("Office",
     "*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#"
     "C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL#"
     "Microsoft Office 16.0 Object Library"),
]


# ---------------------------------------------------------------------------
# Compresion MS-OVBA 2.4.1
# ---------------------------------------------------------------------------
def compress(data: bytes) -> bytes:
    out = bytearray([0x01])
    for chunk_start in range(0, len(data), 4096):
        chunk_end = min(chunk_start + 4096, len(data))
        body = _compress_chunk(data, chunk_start, chunk_end)
        if len(body) <= 4096:
            out += struct.pack("<H", 0xB000 | (len(body) + 2 - 3))
            out += body
        elif chunk_end - chunk_start == 4096:
            out += struct.pack("<H", 0x3000 | 0x0FFF)
            out += data[chunk_start:chunk_end]
        else:
            raise ValueError("bloque final no comprimible")
    return bytes(out)


def _compress_chunk(data, start, end):
    body = bytearray()
    index = {}  # prefijo de 3 bytes -> posiciones dentro del bloque
    pos = start

    def remember(p):
        if p + 3 <= end:
            index.setdefault(data[p:p + 3], []).append(p)

    while pos < end:
        flag_at = len(body)
        body.append(0)
        flags = 0
        for bit in range(8):
            if pos >= end:
                break
            best_len, best_pos = 0, 0
            for cand in reversed(index.get(data[pos:pos + 3], ())[-512:]):
                n = 0
                while pos + n < end and data[cand + n] == data[pos + n]:
                    n += 1
                if n > best_len:
                    best_len, best_pos = n, cand
            if best_len >= 3:
                bit_count = max((pos - start - 1).bit_length(), 4)
                length = min(best_len, (0xFFFF >> bit_count) + 3)
                offset = pos - best_pos
                token = ((offset - 1) << (16 - bit_count)) | (length - 3)
                body += struct.pack("<H", token)
                flags |= 1 << bit
                for p in range(pos, pos + length):
                    remember(p)
                pos += length
            else:
                body.append(data[pos])
                remember(pos)
                pos += 1
        body[flag_at] = flags
    return body


# ---------------------------------------------------------------------------
# Cifrado de datos MS-OVBA 2.4.3 (CMG, DPB, GC del stream PROJECT)
# ---------------------------------------------------------------------------
def encrypt(data: bytes, project_id: str, rng: random.Random) -> str:
    seed = rng.randrange(256)
    version = 2
    proj_key = sum(project_id.encode("ascii")) & 0xFF
    out = [seed, seed ^ version, seed ^ proj_key]
    unenc1, enc1, enc2 = proj_key, out[2], out[1]

    def push(b):
        nonlocal unenc1, enc1, enc2
        e = b ^ ((enc2 + unenc1) & 0xFF)
        out.append(e)
        enc2, enc1, unenc1 = enc1, e, b

    for _ in range((seed & 6) // 2):
        push(rng.randrange(256))
    for b in struct.pack("<I", len(data)) + data:
        push(b)
    return bytes(out).hex().upper()


def decrypt(hexstr: str) -> bytes:
    b = bytes.fromhex(hexstr)
    seed, version_enc, proj_key_enc = b[0], b[1], b[2]
    unenc1, enc1, enc2 = seed ^ proj_key_enc, proj_key_enc, version_enc
    plain = []
    for e in b[3:]:
        p = e ^ ((enc2 + unenc1) & 0xFF)
        plain.append(p)
        enc2, enc1, unenc1 = enc1, e, p
    ignored = (seed & 6) // 2
    length = struct.unpack("<I", bytes(plain[ignored:ignored + 4]))[0]
    return bytes(plain[ignored + 4:ignored + 4 + length])


# ---------------------------------------------------------------------------
# Streams del proyecto
# ---------------------------------------------------------------------------
def _rec(rid, payload=b""):
    return struct.pack("<HI", rid, len(payload)) + payload


def _mbcs(s):
    return s.encode("cp1252")


def _utf16(s):
    return s.encode("utf-16-le")


def document_source(name, base):
    return "\r\n".join([
        'Attribute VB_Name = "%s"' % name,
        'Attribute VB_Base = "%s"' % base,
        "Attribute VB_GlobalNameSpace = False",
        "Attribute VB_Creatable = False",
        "Attribute VB_PredeclaredId = True",
        "Attribute VB_Exposed = True",
        "Attribute VB_TemplateDerived = False",
        "Attribute VB_Customizable = True",
    ]) + "\r\n"


def normalize_source(text):
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if not text.endswith("\n"):
        text += "\n"
    return text.replace("\n", "\r\n")


def _dir_stream(modules):
    d = bytearray()
    d += _rec(0x0001, struct.pack("<I", 1))            # SYSKIND win32
    d += _rec(0x0002, struct.pack("<I", LCID))
    d += _rec(0x0014, struct.pack("<I", LCID))
    d += _rec(0x0003, struct.pack("<H", CODEPAGE))
    d += _rec(0x0004, b"VBAProject")
    d += _rec(0x0005) + _rec(0x0040)                   # DOCSTRING
    d += _rec(0x0006) + _rec(0x003D)                   # HELPFILEPATH
    d += _rec(0x0007, struct.pack("<I", 0))            # HELPCONTEXT
    d += _rec(0x0008, struct.pack("<I", 0))            # LIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 0x6B5E4A31, 1)  # VERSION
    d += _rec(0x000C) + _rec(0x003C)                   # CONSTANTS
    for name, libid in REFERENCES:
        d += _rec(0x0016, _mbcs(name)) + _rec(0x003E, _utf16(name))
        lib = libid.encode("cp1252")
        d += _rec(0x000D, struct.pack("<I", len(lib)) + lib + struct.pack("<IH", 0, 0))
    d += _rec(0x000F, struct.pack("<H", len(modules)))
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))       # PROJECTCOOKIE
    for m in modules:
        d += _rec(0x0019, _mbcs(m["name"])) + _rec(0x0047, _utf16(m["name"]))
        d += _rec(0x001A, _mbcs(m["name"])) + _rec(0x0032, _utf16(m["name"]))
        d += _rec(0x001C) + _rec(0x0048)               # DOCSTRING
        d += _rec(0x0031, struct.pack("<I", 0))        # OFFSET del codigo
        d += _rec(0x001E, struct.pack("<I", 0))        # HELPCONTEXT
        d += _rec(0x002C, struct.pack("<H", 0xFFFF))   # COOKIE
        d += _rec(0x0022 if m["document"] else 0x0021)
        d += _rec(0x002B)                              # fin de modulo
    d += _rec(0x0010)                                  # fin de dir
    return bytes(d)


def _project_stream(modules, project_id, rng):
    lines = ['ID="%s"' % project_id]
    for m in modules:
        if m["document"]:
            lines.append("Document=%s/&H00000000" % m["name"])
        else:
            lines.append("Module=%s" % m["name"])
    lines += [
        'Name="VBAProject"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG="%s"' % encrypt(b"\x00\x00\x00\x00", project_id, rng),
        'DPB="%s"' % encrypt(b"\x00", project_id, rng),
        'GC="%s"' % encrypt(b"\xFF", project_id, rng),
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    for m in modules:
        lines.append("%s=0, 0, 0, 0, C" % m["name"] if m["document"]
                     else "%s=26, 26, 1100, 640, " % m["name"])
    return ("\r\n".join(lines) + "\r\n").encode("cp1252")


def _projectwm_stream(modules):
    out = bytearray()
    for m in modules:
        out += _mbcs(m["name"]) + b"\x00" + _utf16(m["name"]) + b"\x00\x00"
    return bytes(out + b"\x00\x00")


# ---------------------------------------------------------------------------
# Contenedor MS-CFB version 3 (sectores de 512 bytes)
# ---------------------------------------------------------------------------
FREESECT, ENDOFCHAIN, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF


def _cfb_key(name):
    return (len(name), name.upper())


def write_cfb(tree):
    """tree: {"name": str, "children": [...]} o {"name": str, "data": bytes}."""
    entries = []

    def add(node, kind):
        node = dict(node)
        node["type"] = kind
        node["id"] = len(entries)
        entries.append(node)
        if kind in (1, 5):
            node["kids"] = [add(c, 1 if "children" in c else 2) for c in node["children"]]
        return node["id"]

    add(tree, 5)

    # Arbol rojo-negro balanceado de cada almacenamiento.
    for e in entries:
        e.update(left=NOSTREAM, right=NOSTREAM, child=NOSTREAM, color=1)
    for e in entries:
        if e["type"] not in (1, 5) or not e["kids"]:
            continue
        kids = sorted(e["kids"], key=lambda i: _cfb_key(entries[i]["name"]))
        depth = {}

        def build(lo, hi, level):
            if lo > hi:
                return NOSTREAM
            mid = (lo + hi) // 2
            node = entries[kids[mid]]
            depth[node["id"]] = level
            node["left"] = build(lo, mid - 1, level + 1)
            node["right"] = build(mid + 1, hi, level + 1)
            return node["id"]

        e["child"] = build(0, len(kids) - 1, 0)
        max_level = max(depth.values())
        full = len(kids) == 2 ** (max_level + 1) - 1
        if not full and max_level > 0:
            for i, lvl in depth.items():
                if lvl == max_level:
                    entries[i]["color"] = 0

    # Mini stream para streams < 4096 bytes.
    mini = bytearray()
    minifat = []
    big = []
    for e in entries:
        if e["type"] != 2:
            continue
        data = e["data"]
        e["size"] = len(data)
        if len(data) == 0:
            e["start"] = ENDOFCHAIN
        elif len(data) < 4096:
            first = len(mini) // 64
            count = (len(data) + 63) // 64
            minifat += [first + k + 1 for k in range(count - 1)] + [ENDOFCHAIN]
            e["start"] = first
            mini += data + b"\x00" * (count * 64 - len(data))
        else:
            big.append(e)

    def nsec(nbytes):
        return (nbytes + 511) // 512

    n_dir = nsec(len(entries) * 128)
    n_minifat = nsec(len(minifat) * 4)
    n_mini = nsec(len(mini))
    n_big = sum(nsec(len(e["data"])) for e in big)
    n_fat = 1
    while n_fat * 128 < n_fat + n_dir + n_minifat + n_mini + n_big:
        n_fat += 1
    if n_fat > 109:
        raise ValueError("archivo demasiado grande para este escritor")

    fat = []
    sectors = []

    def chain(count):
        start = len(fat)
        fat.extend(start + k + 1 for k in range(count - 1))
        fat.append(ENDOFCHAIN)
        return start

    fat.extend([FATSECT] * n_fat)
    dir_start = chain(n_dir)
    minifat_start = chain(n_minifat) if n_minifat else ENDOFCHAIN
    mini_start = chain(n_mini) if n_mini else ENDOFCHAIN
    for e in big:
        e["start"] = chain(nsec(len(e["data"])))
    fat += [FREESECT] * (n_fat * 128 - len(fat))

    root = entries[0]
    root["start"], root["size"] = mini_start, len(mini)

    def pad(b):
        return bytes(b) + b"\x00" * (-len(b) % 512)

    dir_bytes = bytearray()
    for e in entries:
        name = _utf16(e["name"])
        if len(name) > 62:
            raise ValueError("nombre demasiado largo: " + e["name"])
        start = e.get("start", 0)
        size = e.get("size", 0)
        dir_bytes += (name + b"\x00" * (64 - len(name)))
        dir_bytes += struct.pack("<HBBIII", len(name) + 2, e["type"], e["color"],
                                 e["left"], e["right"], e["child"])
        dir_bytes += b"\x00" * 16 + struct.pack("<I", 0) + b"\x00" * 16
        dir_bytes += struct.pack("<IQ", start, size)
    while len(dir_bytes) % 512:
        dir_bytes += b"\x00" * 64 + struct.pack("<HBBIII", 0, 0, 0, NOSTREAM, NOSTREAM, NOSTREAM)
        dir_bytes += b"\x00" * 36 + struct.pack("<IQ", 0, 0)

    minifat_bytes = struct.pack("<%dI" % len(minifat), *minifat)
    minifat_bytes += struct.pack("<I", FREESECT) * (-(len(minifat)) % 128)

    header = bytearray(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" + b"\x00" * 16)
    header += struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6) + b"\x00" * 6
    header += struct.pack("<IIIIIIIII", 0, n_fat, dir_start, 0, 4096,
                          minifat_start, n_minifat, ENDOFCHAIN, 0)
    difat = list(range(n_fat)) + [FREESECT] * (109 - n_fat)
    header += struct.pack("<109I", *difat)
    assert len(header) == 512

    body = bytearray(header)
    body += struct.pack("<%dI" % len(fat), *fat)
    body += dir_bytes
    if n_minifat:
        body += minifat_bytes
    if n_mini:
        body += pad(mini)
    for e in big:
        body += pad(e["data"])
    return bytes(body)


# ---------------------------------------------------------------------------
def build_vba_project(standard_modules, workbook_codename, sheet_codenames,
                      workbook_code="", seed=1):
    """standard_modules: lista de (nombre, texto fuente con Attribute VB_Name).
    workbook_code: codigo opcional para el modulo ThisWorkbook (eventos)."""
    rng = random.Random(seed)
    project_id = "{%s}" % str(uuid.UUID(int=rng.getrandbits(128), version=4)).upper()

    modules = [{"name": workbook_codename, "document": True,
                "source": document_source(workbook_codename, BASE_WORKBOOK)
                + (normalize_source(workbook_code) if workbook_code else "")}]
    for cn in sheet_codenames:
        modules.append({"name": cn, "document": True,
                        "source": document_source(cn, BASE_WORKSHEET)})
    for name, text in standard_modules:
        src = normalize_source(text)
        if not src.startswith('Attribute VB_Name = "%s"' % name):
            raise ValueError("el modulo %s debe iniciar con Attribute VB_Name" % name)
        modules.append({"name": name, "document": False, "source": src})

    vba_children = [{"name": "_VBA_PROJECT", "data": b"\xCC\x61\xFF\xFF\x00\x00\x00"},
                    {"name": "dir", "data": compress(_dir_stream(modules))}]
    for m in modules:
        vba_children.append({"name": m["name"],
                             "data": compress(m["source"].encode("cp1252"))})
    tree = {"name": "Root Entry", "children": [
        {"name": "VBA", "children": vba_children},
        {"name": "PROJECT", "data": _project_stream(modules, project_id, rng)},
        {"name": "PROJECTwm", "data": _projectwm_stream(modules)},
    ]}
    return write_cfb(tree)


def load_bas(path):
    with open(path, encoding="utf-8") as f:
        text = f.read()
    name = os.path.splitext(os.path.basename(path))[0]
    return name, text
