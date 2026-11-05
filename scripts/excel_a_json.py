"""Convierte datos/productos.xlsx en productos.json para la web Ofertas del Norte.

Uso:  python scripts/excel_a_json.py [ruta_excel] [ruta_salida]
Lo ejecuta GitHub Actions cada vez que subes un Excel nuevo.
"""
import json
import os
import re
import sys
import unicodedata

import openpyxl

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL = sys.argv[1] if len(sys.argv) > 1 else os.path.join(RAIZ, "datos", "productos.xlsx")
SALIDA = sys.argv[2] if len(sys.argv) > 2 else os.path.join(RAIZ, "productos.json")

COLUMNAS = {
    "mostrar": "mostrar",
    "producto": "nombre",
    "tipo": "tipo",
    "categoria": "tipo",
    "marca": "marca",
    "detalle": "detalle",
    "precio hoy": "precio",
    "precio antes": "antes",
    "descuento": "descuento",
    "cantidad": "cantidad",
    "unidad": "unidad",
    "link afiliado": "link",
    "codigo ml": "codigo",
    "stock": "stock",
    "imagen": "imagen",
    "destacado": "destacado",
}


def norm(txt):
    txt = unicodedata.normalize("NFD", str(txt or "")).encode("ascii", "ignore").decode()
    txt = re.sub(r"\(.*?\)", "", txt.lower())
    return re.sub(r"\s+", " ", txt).strip()


def si(v):
    return norm(v) in ("si", "s", "x", "1", "true", "yes")


def numero(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).replace("$", "").replace(" ", "").replace("%", "")
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")
    elif t.count(".") >= 1 and len(t.split(".")[-1]) == 3:
        t = t.replace(".", "")
    else:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def slug(txt):
    return re.sub(r"[^a-z0-9]+", "-", norm(txt)).strip("-")


def main():
    if not os.path.exists(EXCEL):
        sys.exit(f"No encuentro el Excel: {EXCEL}")
    wb = openpyxl.load_workbook(EXCEL, data_only=True)
    if "Productos" not in wb.sheetnames:
        sys.exit("El Excel debe tener una hoja llamada 'Productos'.")
    ws = wb["Productos"]

    fila_titulos, mapa = None, {}
    for r in range(1, 16):
        valores = [norm(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if "producto" in valores:
            fila_titulos = r
            for c, v in enumerate(valores, 1):
                if v in COLUMNAS:
                    mapa[COLUMNAS[v]] = c
            break
    if not fila_titulos:
        sys.exit("No encontré la fila de títulos (debe tener una columna 'Producto').")

    avisos, productos, vistos = [], [], set()
    for r in range(fila_titulos + 1, ws.max_row + 1):
        dato = {k: ws.cell(r, c).value for k, c in mapa.items()}
        nombre = str(dato.get("nombre") or "").strip()
        if not nombre:
            continue
        if "mostrar" in mapa and not si(dato.get("mostrar")):
            continue
        precio = numero(dato.get("precio"))
        link = str(dato.get("link") or "").strip()
        if not precio or not link:
            avisos.append(f"Fila {r} ({nombre}): falta precio o link, no se muestra.")
            continue
        if not link.startswith("http"):
            link = "https://" + link
        antes = numero(dato.get("antes"))
        desc = numero(dato.get("descuento"))
        if desc is not None and desc > 1:
            desc = desc / 100
        if antes and antes <= precio:
            antes = None
        if not desc and antes:
            # Si no escribiste el %, se calcula como lo muestra Mercado Libre (sin redondear hacia arriba)
            desc = int((1 - precio / antes) * 100 + 1e-9) / 100
        cantidad = numero(dato.get("cantidad"))
        unidad = norm(dato.get("unidad")) or None
        imagen = str(dato.get("imagen") or "").strip()
        es_url = imagen.startswith("http")
        if imagen and not es_url and not os.path.exists(os.path.join(RAIZ, "img", imagen)):
            avisos.append(f"Fila {r} ({nombre}): no existe img/{imagen}.")
            imagen = ""
        tipo = str(dato.get("tipo") or "").strip() or "Otros"
        tipo = tipo[0].upper() + tipo[1:]
        stock = str(dato.get("stock") or "Disponible").strip()
        base = slug(nombre) or f"p{r}"
        uid, n = base, 2
        while uid in vistos:
            uid, n = f"{base}-{n}", n + 1
        vistos.add(uid)
        productos.append({
            "id": uid,
            "nombre": nombre,
            "tipo": tipo,
            "marca": str(dato.get("marca") or "").strip(),
            "detalle": str(dato.get("detalle") or "").strip(),
            "precio": round(precio),
            "antes": round(antes) if antes else None,
            "descuento": round(desc * 100) if desc else None,
            "cantidad": cantidad,
            "unidad": unidad if unidad in ("kg", "un") else None,
            "precio_por": round(precio / cantidad) if cantidad else None,
            "link": link,
            "codigo": str(dato.get("codigo") or "").strip().strip("—-") or None,
            "stock": stock,
            "imagen": (imagen if es_url else "img/" + imagen) if imagen else None,
            "destacado": si(dato.get("destacado")),
        })

    config = {}
    if "Configuración" in wb.sheetnames:
        wc = wb["Configuración"]
        for r in range(1, wc.max_row + 1):
            k, v = wc.cell(r, 1).value, wc.cell(r, 2).value
            if k and norm(k) not in ("dato", "configuracion de la web"):
                config[slug(k).replace("-", "_")] = str(v).strip() if v is not None else ""
    if config.get("whatsapp"):
        config["whatsapp"] = re.sub(r"\D", "", config["whatsapp"])

    if not productos:
        sys.exit("No quedó ningún producto para mostrar. Revisa la hoja Productos.")

    with open(SALIDA, "w", encoding="utf-8") as f:
        json.dump({"config": config, "productos": productos}, f, ensure_ascii=False, indent=1)
    print(f"OK: {len(productos)} productos -> {SALIDA}")
    for a in avisos:
        print("AVISO:", a)


if __name__ == "__main__":
    main()
