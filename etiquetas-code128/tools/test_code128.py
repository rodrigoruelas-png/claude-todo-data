"""Pruebas del codificador Code 128 (tools/code128.py, igual a ModCode128.bas).

    python3 tools/test_code128.py [ruta/LibreBarcode128-Regular.ttf]

1. Verificador mod 103 contra ejemplos calculados a mano.
2. Si se da la fuente y estan instalados Pillow y zxing-cpp: dibuja cada texto
   con la fuente (3 px por modulo, como la Zebra de 203 dpi) y lo decodifica.
   Los textos cubren los 107 valores de simbolo (todo ASCII 32-126 en el
   subconjunto B y los pares 00-99 en el subconjunto C).
"""

import os
import string
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import code128  # noqa: E402

failures = 0


def check(cond, msg):
    global failures
    if not cond:
        failures += 1
        print("FALLA:", msg)


# 1. Valores y verificador calculados a mano --------------------------------
# 50426529 -> Start C, 50 42 65 29; (105 + 1*50 + 2*42 + 3*65 + 4*29) % 103 = 35
check(code128.values("50426529") == [105, 50, 42, 65, 29, 35, 106], "50426529")
# 120 -> Start B, '1' '2' '0'; (104 + 17 + 2*18 + 3*16) % 103 = 102
check(code128.values("120") == [104, 17, 18, 16, 102, 106], "120")
# AB12345: corrida impar al final -> '1' en B y luego C 23 45
check(code128.values("AB12345")[:7] == [104, 33, 34, 17, 99, 23, 45], "AB12345")
# 12345AB: corrida impar al inicio -> C 12 34, Code B, '5' 'A' 'B'
check(code128.values("12345AB")[:7] == [105, 12, 34, 100, 21, 33, 34], "12345AB")
check(code128.encode("50426529") == "ÍRJa=CÎ", "mapeo LIBRE")
check(code128.encode("50426529", "GRANDZEBU") == "ÒRJa=CÓ", "mapeo GRANDZEBU")
check(code128.encode("00") == "ÍÂ\"Î", "valor 0 -> Chr 194")
check(code128.encode("Ñ123") == "" and code128.encode("") == "", "caracteres no validos")

# 2. Dibujar con la fuente y decodificar ------------------------------------
font_path = sys.argv[1] if len(sys.argv) > 1 else None
if font_path:
    from PIL import Image, ImageDraw, ImageFont
    import zxingcpp

    font = ImageFont.truetype(font_path, 100)   # 0.03 em por modulo = 3 px

    def roundtrip(text):
        enc = code128.encode(text)
        w = int(font.getlength(enc)) + 120
        img = Image.new("L", (w, 80), 255)
        ImageDraw.Draw(img).text((60, 10), enc, font=font, fill=0)
        found = zxingcpp.read_barcodes(img, formats=zxingcpp.BarcodeFormat.Code128)
        return found[0].text if found else None

    ascii_b = "".join(chr(c) for c in range(32, 127))
    samples = [ascii_b[i:i + 12].strip() or "A A" for i in range(0, len(ascii_b), 12)]
    samples += ["".join("%02d" % n for n in range(k, min(k + 10, 100))) for k in range(0, 100, 10)]
    samples += ["50426529", "120", "44-3263-1083-1058", "XX-XXXX-XXX-XXX", "AB12345",
                "12345AB", "A123456B", "1234567", "0", "12", "123", "PN 12 34",
                string.ascii_uppercase, "OB-0001234"]
    seen = set()
    for t in samples:
        v = code128.values(t)
        seen.update(v)
        got = roundtrip(t)
        check(got == t, "decodificado %r -> %r" % (t, got))
    check(seen >= set(range(107)) - {101, 103}, "faltan valores: %s"
          % sorted(set(range(107)) - {101, 103} - seen))
    print("decodificados:", len(samples), "textos; valores de simbolo cubiertos:", len(seen))

print("OK" if failures == 0 else "%d fallas" % failures)
sys.exit(1 if failures else 0)
