import os
os.environ["OPENBLAS_NUM_THREADS"] = "1"

import sqlite3
from datetime import datetime
import io
import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageChops
import streamlit as st
import urllib.parse
import urllib.request
import urllib.error
import html
import json
import zipfile
import tempfile
import base64

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# -------------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA
# -------------------------------------------------------------
logo_icon = None
for posible_logo in ["logo.jpg", "logo.png", "logo.jpeg", "21237.jpg"]:
    if os.path.exists(posible_logo):
        try:
            logo_icon = Image.open(posible_logo)
            break
        except Exception:
            pass

st.set_page_config(
    page_title="Minici Store",
    page_icon=logo_icon if logo_icon else "🛍️",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# -------------------------------------------------------------
# 1B. CONFIGURACIÓN SEGURA Y RESPALDOS EN GITHUB
# -------------------------------------------------------------
DB_FILE = "minici_store.db"
GH_RUTA_ZIP = "respaldos/respaldo_completo.zip"
GH_RUTA_DB = "respaldos/minici_store.db"


def _secret(nombre, default=None):
    """Lee un valor de los 'Secrets' de Streamlit sin romper si no existen."""
    try:
        valor = st.secrets[nombre]
        return valor if valor not in (None, "") else default
    except Exception:
        return default


def _gh_config():
    token, repo = _secret("GITHUB_TOKEN"), _secret("GITHUB_REPO")
    if not token or not repo:
        return None
    return {"token": str(token).strip(),
            "repo": str(repo).strip().strip("/"),
            "branch": str(_secret("GITHUB_BRANCH", "main")).strip()}


def _gh_request(method, url, token, payload=None, raw=False):
    cuerpo = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=cuerpo, method=method)
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Accept", "application/vnd.github.raw+json" if raw else "application/vnd.github+json")
    req.add_header("X-GitHub-Api-Version", "2022-11-28")
    req.add_header("User-Agent", "minici-store-respaldo")
    if cuerpo is not None:
        req.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


def _gh_mensaje_error(e):
    if isinstance(e, urllib.error.HTTPError):
        if e.code == 401:
            return "GitHub rechazó el token (401). Revisa que GITHUB_TOKEN sea correcto y no haya vencido."
        if e.code == 403:
            return "GitHub negó el permiso (403). El token necesita 'Contents: Read and write' sobre el repositorio."
        if e.code == 404:
            return "GitHub no encontró el repositorio o la rama (404). Revisa GITHUB_REPO (usuario/repositorio) y GITHUB_BRANCH."
        return f"GitHub respondió con el error {e.code}."
    return f"No se pudo conectar con GitHub: {e}"


def github_subir(ruta_repo, contenido, mensaje):
    """Crea o actualiza un archivo en el repositorio de respaldos. Devuelve (ok, texto)."""
    cfg = _gh_config()
    if not cfg:
        return False, "GitHub no está configurado."
    url = f"https://api.github.com/repos/{cfg['repo']}/contents/{urllib.parse.quote(ruta_repo)}"
    try:
        sha = None
        try:
            info = json.loads(_gh_request("GET", f"{url}?ref={urllib.parse.quote(cfg['branch'])}", cfg["token"]))
            sha = info.get("sha")
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise
        payload = {"message": mensaje, "branch": cfg["branch"],
                   "content": base64.b64encode(contenido).decode("ascii")}
        if sha:
            payload["sha"] = sha
        _gh_request("PUT", url, cfg["token"], payload)
        return True, "ok"
    except Exception as e:
        return False, _gh_mensaje_error(e)


def github_leer(ruta_repo):
    cfg = _gh_config()
    url = f"https://api.github.com/repos/{cfg['repo']}/contents/{urllib.parse.quote(ruta_repo)}?ref={urllib.parse.quote(cfg['branch'])}"
    return _gh_request("GET", url, cfg["token"], raw=True)


def _extraer_fotos_zip(z):
    """Extrae solo las fotos de productos de un respaldo (sin rutas peligrosas)."""
    for nombre in z.namelist():
        if nombre.startswith("fotos_productos/") and not nombre.endswith("/") and ".." not in nombre:
            destino = os.path.join("fotos_productos", os.path.basename(nombre))
            os.makedirs("fotos_productos", exist_ok=True)
            with open(destino, "wb") as f:
                f.write(z.read(nombre))


def _restaurar_al_iniciar():
    """Si el servidor se reinició y la base de datos desapareció, la recupera del último respaldo en GitHub."""
    if os.path.exists(DB_FILE) or not _gh_config():
        return False
    try:
        datos = github_leer(GH_RUTA_ZIP)
        with zipfile.ZipFile(io.BytesIO(datos)) as z:
            if DB_FILE in z.namelist():
                with open(DB_FILE, "wb") as f:
                    f.write(z.read(DB_FILE))
                _extraer_fotos_zip(z)
                return True
    except Exception:
        pass
    try:
        datos = github_leer(GH_RUTA_DB)
        with open(DB_FILE, "wb") as f:
            f.write(datos)
        return True
    except Exception:
        return False


if _restaurar_al_iniciar():
    st.session_state["_restaurado_gh"] = True

# -------------------------------------------------------------
# 2. BASE DE DATOS SQLITE Y MIGRACIONES AUTOMÁTICAS
# -------------------------------------------------------------
conn = sqlite3.connect(DB_FILE, check_same_thread=False)
c = conn.cursor()

c.execute("""CREATE TABLE IF NOT EXISTS clientes (
               id_cliente TEXT PRIMARY KEY,
               nombre TEXT NOT NULL,
               telefono TEXT,
               correo TEXT)""")

c.execute("""CREATE TABLE IF NOT EXISTS cajas_emprendedores (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               id_cliente TEXT UNIQUE,
               nombre_caja TEXT,
               precio_caja REAL DEFAULT 0.0,
               FOREIGN KEY(id_cliente) REFERENCES clientes(id_cliente))""")

c.execute("""CREATE TABLE IF NOT EXISTS productos (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               id_cliente TEXT,
               tienda TEXT,
               categoria TEXT,
               descripcion TEXT,
               precio REAL DEFAULT 0.0,
               moneda TEXT DEFAULT 'CRC',
               cantidad INTEGER DEFAULT 1,
               estado TEXT,
               observaciones TEXT,
               foto_path TEXT,
               codigo_barras TEXT,
               FOREIGN KEY(id_cliente) REFERENCES clientes(id_cliente))""")

c.execute("""CREATE TABLE IF NOT EXISTS ventas_rapidas (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               fecha TEXT,
               id_cliente TEXT,
               producto TEXT,
               tipo TEXT,
               cantidad INTEGER,
               precio_u REAL,
               descuento REAL,
               total REAL)""")

columnas_necesarias_productos = [
    ("categoria", "TEXT"),
    ("precio", "REAL DEFAULT 0.0"),
    ("moneda", "TEXT DEFAULT 'CRC'"),
    ("cantidad", "INTEGER DEFAULT 1"),
    ("observaciones", "TEXT"),
    ("foto_path", "TEXT"),
    ("codigo_barras", "TEXT")
]

for col_name, col_type in columnas_necesarias_productos:
    try:
        c.execute(f"ALTER TABLE productos ADD COLUMN {col_name} {col_type}")
        conn.commit()
    except sqlite3.OperationalError:
        pass

c.execute("""CREATE TABLE IF NOT EXISTS abonos (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               id_cliente TEXT,
               monto_crc REAL,
               fecha TEXT)""")

c.execute("""CREATE TABLE IF NOT EXISTS gastos (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               concepto TEXT NOT NULL,
               categoria TEXT,
               monto_crc REAL NOT NULL,
               fecha TEXT NOT NULL,
               observaciones TEXT)""")

c.execute("""CREATE TABLE IF NOT EXISTS notificaciones (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               id_cliente TEXT,
               titulo TEXT,
               mensaje TEXT,
               leida INTEGER DEFAULT 0,
               fecha TEXT)""")
conn.commit()

if not os.path.exists("fotos_productos"):
    os.makedirs("fotos_productos")
if not os.path.exists("facturas_generadas"):
    os.makedirs("facturas_generadas")

# -------------------------------------------------------------
# UTILIDADES: TEXTO, MENSAJES Y RESPALDO / RESTAURACIÓN
# -------------------------------------------------------------
def cap1(texto):
    """Pone en mayúscula solo la primera letra (sin borrar mayúsculas como en 'Nike Air' o 'MAC')."""
    t = str(texto or "").strip()
    return t[:1].upper() + t[1:]


def flash(tipo, mensaje):
    """Guarda un mensaje para mostrarlo después de st.rerun() (si no, se pierde)."""
    st.session_state["_flash"] = (tipo, mensaje)


def mostrar_flash():
    f = st.session_state.pop("_flash", None)
    if f:
        getattr(st, f[0], st.info)(f[1])


def crear_respaldo_zip():
    """Zip con la base de datos, todas las fotos y un Excel legible. Devuelve (zip, db)."""
    conn.commit()
    tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
    tmp.close()
    try:
        destino = sqlite3.connect(tmp.name)
        conn.backup(destino)
        destino.close()
        with open(tmp.name, "rb") as f:
            db_bytes = f.read()
    finally:
        os.unlink(tmp.name)

    xlsx_bytes = None
    try:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            for tabla, hoja in [("clientes", "Clientes"), ("cajas_emprendedores", "Cajas"),
                                ("productos", "Productos"), ("abonos", "Abonos"),
                                ("gastos", "Gastos"), ("ventas_rapidas", "Ventas POS"),
                                ("notificaciones", "Notificaciones")]:
                pd.read_sql(f"SELECT * FROM {tabla}", conn).to_excel(w, sheet_name=hoja, index=False)
        xlsx_bytes = buf.getvalue()
    except Exception:
        pass

    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(DB_FILE, db_bytes)
        if xlsx_bytes:
            z.writestr("datos_minici.xlsx", xlsx_bytes)
        if os.path.isdir("fotos_productos"):
            for nombre in os.listdir("fotos_productos"):
                ruta = os.path.join("fotos_productos", nombre)
                if os.path.isfile(ruta):
                    z.write(ruta, f"fotos_productos/{nombre}")
    return zbuf.getvalue(), db_bytes


def hacer_respaldo():
    zip_bytes, db_bytes = crear_respaldo_zip()
    ahora = datetime.now()
    res = {"zip": zip_bytes, "nombre": f"respaldo_minici_{ahora:%Y-%m-%d_%H%M}.zip",
           "hora": ahora.strftime("%d/%m/%Y %H:%M"), "github": None, "detalle": ""}
    if _gh_config():
        msg = f"Respaldo Minici Store {ahora:%Y-%m-%d %H:%M}"
        ok_db, m_db = github_subir(GH_RUTA_DB, db_bytes, msg)
        if len(zip_bytes) < 90 * 1024 * 1024:
            ok_zip, m_zip = github_subir(GH_RUTA_ZIP, zip_bytes, msg)
        else:
            ok_zip, m_zip = False, "El respaldo con fotos pesa más de 90 MB; solo se subió la base de datos."
        res["github"] = bool(ok_db and ok_zip)
        res["detalle"] = "" if res["github"] else (m_db if not ok_db else m_zip)
    return res


def restaurar_respaldo(datos):
    """Reemplaza los datos actuales con los de un respaldo (.zip o .db). Devuelve (ok, texto)."""
    conn.commit()
    zf = None
    tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
    tmp.close()
    try:
        if zipfile.is_zipfile(io.BytesIO(datos)):
            zf = zipfile.ZipFile(io.BytesIO(datos))
            if DB_FILE not in zf.namelist():
                return False, f"El .zip no contiene {DB_FILE}."
            contenido_db = zf.read(DB_FILE)
        else:
            contenido_db = datos
        with open(tmp.name, "wb") as f:
            f.write(contenido_db)
        origen = sqlite3.connect(tmp.name)
        try:
            tablas = [r[0] for r in origen.execute("SELECT name FROM sqlite_master WHERE type='table'")]
            if "clientes" not in tablas or "productos" not in tablas:
                return False, "Ese archivo no parece un respaldo de Minici Store."
            if origen.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
                return False, "El archivo de respaldo está dañado."
            origen.backup(conn)
        finally:
            origen.close()
        if zf:
            _extraer_fotos_zip(zf)
        return True, "Datos restaurados correctamente."
    except Exception as e:
        return False, f"No se pudo restaurar: {e}"
    finally:
        os.unlink(tmp.name)


# -------------------------------------------------------------
# FACTURAS: IMAGEN + MENSAJE DE WHATSAPP
# -------------------------------------------------------------
import functools

FACT_COLORS = {
    "raspberry": "#BE185D", "plum": "#831843", "brand": "#F3B2C9",
    "soft": "#FDF2F8", "zebra": "#FFF7FB", "line": "#F6D5E4",
    "text": "#1F2937", "muted": "#6B7280", "green": "#059669",
    "orange": "#F5A623", "white": "#FFFFFF",
}

_FUENTES_REG = [
    "arial.ttf", "Arial.ttf", "LiberationSans-Regular.ttf", "DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    "C:/Windows/Fonts/arial.ttf",
    "/Library/Fonts/Arial.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
]
_FUENTES_BOLD = [
    "arialbd.ttf", "Arial Bold.ttf", "LiberationSans-Bold.ttf", "DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
    "C:/Windows/Fonts/arialbd.ttf",
    "/Library/Fonts/Arial Bold.ttf",
    "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
]


def _tiene_colon(font):
    """True si la fuente sabe dibujar el símbolo ₡ (colón)."""
    try:
        return font.getmask("₡").tobytes() != font.getmask("\ue000").tobytes()
    except Exception:
        return True


@functools.lru_cache(maxsize=32)
def _fuente(bold, size):
    nombres = list(_FUENTES_BOLD if bold else _FUENTES_REG)
    try:  # DejaVu viene incluida con matplotlib, si está instalado
        import matplotlib
        mp = os.path.join(matplotlib.get_data_path(), "fonts", "ttf")
        nombres.append(os.path.join(mp, "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf"))
    except Exception:
        pass
    for n in nombres:
        try:
            f = ImageFont.truetype(n, size)
        except Exception:
            continue
        if _tiene_colon(f):
            return f
    try:
        return ImageFont.load_default(size)
    except TypeError:
        return ImageFont.load_default()


@functools.lru_cache(maxsize=1)
def _logo_factura():
    """Busca el logo igual que el resto de la app, lo recorta a su contenido
    y devuelve (imagen, color_de_fondo)."""
    for nombre in ["logo.jpg", "logo.png", "logo.jpeg", "21237.jpg"]:
        if os.path.exists(nombre):
            try:
                im = Image.open(nombre).convert("RGB")
                bg = im.getpixel((4, 4))
                diff = ImageChops.difference(im, Image.new("RGB", im.size, bg)).convert("L")
                bbox = diff.point(lambda p: 255 if p > 14 else 0).getbbox()
                if bbox:
                    pad = 14
                    bbox = (max(0, bbox[0] - pad), max(0, bbox[1] - pad),
                            min(im.width, bbox[2] + pad), min(im.height, bbox[3] + pad))
                    im = im.crop(bbox)
                return im, bg
            except Exception:
                continue
    return None, None


def _crc(v):
    return f"₡{(v or 0):,.0f}"


def _etiqueta_caja(nombre_caja):
    n = str(nombre_caja).strip()
    return n if "caja" in n.lower() else f"Caja {n}"


def _recortar_msg(txt, n):
    txt = str(txt).strip()
    return txt if len(txt) <= n else txt[: n - 1].rstrip() + "…"


def _mensaje_whatsapp(nombre_cliente, prods, nombre_caja, precio_caja,
                      total_compras, total_abonos, saldo_pendiente):
    linea = "━━━━━━━━━━━━━━"
    m = "🛍️ *MINICI STORE* ✨\n_Personal Shopper_\n" + linea + "\n\n"
    m += f"¡Hola *{nombre_cliente}*! 💖\n"
    m += f"Este es el resumen de tu cuenta al *{datetime.now().strftime('%d/%m/%Y')}*:\n\n"

    if prods:
        m += "📦 *Tus pedidos*\n"
        for p in prods[:6]:
            cant = p[2] if p[2] else 1
            sub = (p[1] or 0.0) * cant
            x = f" ×{cant}" if cant > 1 else ""
            m += f"• {_recortar_msg(p[0], 28)}{x} — {_crc(sub)}\n"
        if len(prods) > 6:
            m += f"• _y {len(prods) - 6} más en tu factura_\n"
        m += "\n"
    if nombre_caja and precio_caja > 0:
        m += f"💼 *{_etiqueta_caja(nombre_caja)}* — {_crc(precio_caja)}\n\n"

    m += linea + "\n"
    m += f"🧾 Total compras: {_crc(total_compras)}\n"
    m += f"✅ Total abonado: {_crc(total_abonos)}\n"
    if saldo_pendiente > 0:
        m += f"💰 *Saldo pendiente: {_crc(saldo_pendiente)}*\n"
    else:
        m += "🎉 *¡Tu cuenta está al día!*\n"
    m += linea + "\n\n"
    m += "Te comparto tu factura en imagen 🖼️\n¡Gracias por confiar en Minici Store! 🌸"
    return m


def generar_factura_imagen(id_cliente):
    c.execute("SELECT nombre, telefono FROM clientes WHERE id_cliente = ?", (id_cliente,))
    res_cli = c.fetchone()
    if not res_cli:
        return None, None

    nombre_cliente, telefono = res_cli

    c.execute("SELECT nombre_caja, precio_caja FROM cajas_emprendedores WHERE id_cliente = ?", (id_cliente,))
    res_caja = c.fetchone()
    nombre_caja, precio_caja = (res_caja[0], res_caja[1] or 0.0) if res_caja else (None, 0.0)

    c.execute("SELECT descripcion, precio, cantidad, estado FROM productos WHERE id_cliente = ?", (id_cliente,))
    prods = c.fetchall()

    c.execute("SELECT COALESCE(SUM(monto_crc), 0) FROM abonos WHERE id_cliente = ?", (id_cliente,))
    total_abonos = c.fetchone()[0] or 0.0

    total_articulos = sum([((p[1] or 0.0) * (p[2] if p[2] else 1)) for p in prods])
    total_compras = total_articulos + precio_caja
    saldo_pendiente = total_compras - total_abonos

    # ---------- Lienzo (se dibuja al doble y se reduce para bordes suaves) ----------
    K = FACT_COLORS
    S, W = 2, 1000
    logo, bg_logo = _logo_factura()
    header_bg = bg_logo if bg_logo else "#FEEFF6"

    filas = []
    if nombre_caja and precio_caja > 0:
        filas.append((f"{_etiqueta_caja(nombre_caja)} · cuota mensual", "1", "Mensual", precio_caja))
    for p in prods:
        cant = p[2] if p[2] else 1
        filas.append((str(p[0] or ""), str(cant), str(p[3] or ""), (p[1] or 0.0) * cant))

    header_h, row_h = 250, 58
    client_top, client_h = header_h + 36, 130
    table_top = client_top + client_h + 36
    rows_top = table_top + 54
    rows_h = max(len(filas), 1) * row_h
    tot_top = rows_top + rows_h + 40
    footer_top = tot_top + 46 * 2 + 96 + 56
    H = footer_top + 120

    img = Image.new("RGB", (W * S, H * S), K["white"])
    draw = ImageDraw.Draw(img)

    def F(bold, size):
        return _fuente(bold, size * S)

    def R(box, fill=None, outline=None, width=1, r=0):
        b = [v * S for v in box]
        if r:
            draw.rounded_rectangle(b, radius=r * S, fill=fill, outline=outline, width=width * S)
        else:
            draw.rectangle(b, fill=fill, outline=outline, width=width * S)

    def tw(txt, font):
        return draw.textlength(txt, font=font) / S

    def T(x, y, txt, font, fill, align="l"):
        w = tw(txt, font)
        if align == "r":
            x -= w
        elif align == "c":
            x -= w / 2
        draw.text((x * S, y * S), txt, font=font, fill=fill)

    def ajustar(txt, font, maxw):
        if tw(txt, font) <= maxw:
            return txt
        while txt and tw(txt + "…", font) > maxw:
            txt = txt[:-1]
        return txt.rstrip() + "…"

    # ---------- Encabezado con logo ----------
    R((0, 0, W, header_h), fill=header_bg)
    if logo:
        lh = 190
        lw = int(logo.width * lh / logo.height)
        if lw > 460:
            lw, lh = 460, int(logo.height * 460 / logo.width)
        logo_r = logo.resize((lw * S, lh * S), Image.LANCZOS)
        img.paste(logo_r, (44 * S, ((header_h - lh) // 2) * S))
    else:
        T(44, 88, "Minici Store", F(True, 46), K["raspberry"])
    T(W - 44, 74, "ESTADO DE CUENTA", F(True, 30), K["plum"], "r")
    T(W - 44, 122, datetime.now().strftime("%d/%m/%Y  ·  %H:%M"), F(False, 20), K["muted"], "r")
    T(W - 44, 152, f"Ref. {id_cliente}", F(True, 20), K["raspberry"], "r")
    R((0, header_h, W, header_h + 6), fill=K["raspberry"])
    R((0, header_h, 150, header_h + 6), fill=K["orange"])

    # ---------- Tarjeta de clienta ----------
    R((40, client_top, W - 40, client_top + client_h), fill=K["soft"], outline=K["line"], width=2, r=20)
    T(66, client_top + 20, "CLIENTA", F(True, 15), K["raspberry"])
    T(66, client_top + 44, ajustar(str(nombre_cliente), F(True, 32), 520), F(True, 32), K["text"])
    T(66, client_top + 90, f"Tel: {telefono if telefono else 'No registrado'}", F(False, 20), K["muted"])
    if nombre_caja:
        etq = _etiqueta_caja(nombre_caja)
        fp = F(True, 20)
        pw = tw(etq, fp) + 44
        R((W - 66 - pw, client_top + 43, W - 66, client_top + 87), fill=K["raspberry"], r=22)
        T(W - 66 - pw / 2, client_top + 53, etq, fp, K["white"], "c")

    # ---------- Tabla de detalle ----------
    R((40, table_top, W - 40, table_top + 52), fill=K["raspberry"], r=14)
    fh = F(True, 19)
    T(66, table_top + 14, "Descripción", fh, K["white"])
    T(600, table_top + 14, "Cant.", fh, K["white"], "c")
    T(670, table_top + 14, "Estado", fh, K["white"])
    T(W - 66, table_top + 14, "Total", fh, K["white"], "r")

    if not filas:
        T(W / 2, rows_top + 16, "Aún no hay artículos registrados", F(False, 21), K["muted"], "c")
    fd, fb = F(False, 21), F(True, 21)
    fe = F(False, 19)
    for i, (desc, cant, estado, subtot) in enumerate(filas):
        y = rows_top + i * row_h
        if i % 2 == 0:
            R((40, y, W - 40, y + row_h), fill=K["zebra"])
        T(66, y + 17, ajustar(desc, fd, 480), fd, K["text"])
        T(600, y + 17, cant, fd, K["text"], "c")
        T(670, y + 18, ajustar(estado, fe, 170), fe, K["muted"])
        T(W - 66, y + 17, _crc(subtot), fb, K["text"], "r")
    R((40, rows_top + rows_h, W - 40, rows_top + rows_h + 2), fill=K["brand"])

    # ---------- Totales ----------
    y = tot_top
    T(500, y + 6, "Total compras", F(False, 21), K["muted"])
    T(W - 66, y + 4, _crc(total_compras), F(True, 24), K["text"], "r")
    y += 46
    T(500, y + 6, "Total abonado", F(False, 21), K["muted"])
    T(W - 66, y + 4, _crc(total_abonos), F(True, 24), K["green"], "r")
    y += 54
    al_dia = saldo_pendiente <= 0
    R((470, y, W - 40, y + 96), fill=K["green"] if al_dia else K["raspberry"], r=20)
    T(500, y + 16, "CUENTA AL DÍA" if al_dia else "SALDO PENDIENTE", F(True, 17), K["white"])
    T(500, y + 44, _crc(max(0.0, saldo_pendiente)), F(True, 38), K["white"])

    # ---------- Pie ----------
    R((0, footer_top, W, H), fill=K["soft"])
    R((0, footer_top, W, footer_top + 5), fill=K["brand"])
    T(W / 2, footer_top + 28, "¡Gracias por tu preferencia en Minici Store!", F(True, 25), K["plum"], "c")
    T(W / 2, footer_top + 72, "P E R S O N A L   S H O P P E R", F(False, 15), K["muted"], "c")

    # ---------- Logo como marca de agua de fondo ----------
    if logo is not None:
        try:
            body_top, body_bot = (header_h + 6) * S, footer_top * S
            wm_w = int(W * 0.66 * S)
            wm_h = int(logo.height * wm_w / logo.width)
            max_h = int((body_bot - body_top) * 0.88)
            if wm_h > max_h:
                wm_h = max_h
                wm_w = int(logo.width * wm_h / logo.height)
            wm = logo.resize((wm_w, wm_h), Image.LANCZOS)
            diff = ImageChops.difference(wm, Image.new("RGB", wm.size, bg_logo)).convert("L")
            mask = diff.point(lambda v: int(min(255, v * 4) * 0.10))
            capa = Image.new("RGB", img.size, "#FFFFFF")
            capa.paste(wm, ((img.width - wm_w) // 2, body_top + ((body_bot - body_top) - wm_h) // 2), mask)
            img = ImageChops.multiply(img, capa)
        except Exception:
            pass

    img = img.resize((W, H), Image.LANCZOS)
    os.makedirs("facturas_generadas", exist_ok=True)
    filepath = os.path.join("facturas_generadas", f"factura_{id_cliente}.jpg")
    img.save(filepath, "JPEG", quality=92, optimize=True)

    telefono_clean = "".join(filter(str.isdigit, str(telefono))) if telefono else ""
    if len(telefono_clean) == 8:
        telefono_clean = "506" + telefono_clean

    mensaje = _mensaje_whatsapp(nombre_cliente, prods, nombre_caja, precio_caja,
                                total_compras, total_abonos, saldo_pendiente)
    link = f"https://wa.me/{telefono_clean}?text={urllib.parse.quote(mensaje)}" if telefono_clean else None

    return filepath, link

_HTML_ESCANER = r'''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  body { margin: 0; padding: 2px; font-family: 'DM Sans', system-ui, -apple-system, sans-serif; color: #4a2c3b; }
  .row { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
  button, label.btn {
    background: #BE185D; color: #fff; border: 1.5px solid #BE185D; border-radius: 999px;
    padding: 10px 16px; font-size: 14px; font-weight: 600; cursor: pointer; font-family: inherit;
  }
  button.sec, label.btn.sec { background: #fff; color: #BE185D; border-color: #F3B2C9; }
  #wrap { display: none; margin-top: 10px; }
  #vbox { position: relative; max-width: 460px; overflow: hidden; border-radius: 14px; background: #000; }
  video { width: 100%; display: block; }
  .guide { position: absolute; left: 10%; top: 28%; width: 80%; height: 44%;
           border: 2px solid #F3B2C9; border-radius: 10px; box-shadow: 0 0 0 999px rgba(0,0,0,.30);
           pointer-events: none; }
  #msg { font-size: 13px; margin-top: 8px; color: #7C6572; min-height: 18px; }
  #msg.ok { color: #059669; font-weight: 600; }
  #msg.err { color: #B91C1C; }
</style>
</head>
<body>
<div class="row">
  <button id="btnCam">📷 Abrir cámara y escanear</button>
  <label class="btn sec">🖼️ Tomar foto del código
    <input id="foto" type="file" accept="image/*" capture="environment" style="display:none">
  </label>
</div>
<div id="wrap">
  <div id="vbox"><video id="v" playsinline muted autoplay></video><div class="guide"></div></div>
  <div class="row" style="margin-top:8px"><button class="sec" id="btnStop">Cerrar cámara</button></div>
</div>
<div id="msg"></div>

<script>
const $ = (id) => document.getElementById(id);
const video = $("v"), wrap = $("wrap"), msg = $("msg");
let stream = null, running = false, timer = null, nativeDet = null;

function post(type, data) {
  window.parent.postMessage(Object.assign({ isStreamlitMessage: true, type: type }, data || {}), "*");
}
function fit() { post("streamlit:setFrameHeight", { height: Math.ceil(document.documentElement.scrollHeight) + 6 }); }
function emit(code) { post("streamlit:setComponentValue", { value: { code: String(code), ts: Date.now() }, dataType: "json" }); }
function say(t, cls) { msg.textContent = t; msg.className = cls || ""; fit(); }

window.addEventListener("message", () => {});
post("streamlit:componentReady", { apiVersion: 1 });
new ResizeObserver(fit).observe(document.body);
setTimeout(fit, 50);

const NATIVE_FORMATS = ["ean_13", "ean_8", "upc_a", "upc_e", "code_128", "code_39", "code_93", "itf", "codabar", "qr_code"];

async function getNative() {
  if (!("BarcodeDetector" in window)) return null;
  try {
    const sup = await BarcodeDetector.getSupportedFormats();
    const f = NATIVE_FORMATS.filter((x) => sup.includes(x));
    return f.length ? new BarcodeDetector({ formats: f }) : null;
  } catch (e) { return null; }
}

function loadScript(src) {
  return new Promise((res, rej) => {
    const s = document.createElement("script");
    s.src = src; s.onload = res; s.onerror = () => rej(new Error("no cargó " + src));
    document.head.appendChild(s);
  });
}
async function ensureZXing() {
  if (window.ZXing) return;
  try { await loadScript("https://cdn.jsdelivr.net/npm/@zxing/library@0.21.3/umd/index.min.js"); }
  catch (e) { await loadScript("https://unpkg.com/@zxing/library@0.21.3/umd/index.min.js"); }
}
function makeReader() {
  const Z = window.ZXing;
  const F = Z.BarcodeFormat;
  const hints = new Map();
  hints.set(Z.DecodeHintType.POSSIBLE_FORMATS,
    [F.EAN_13, F.EAN_8, F.UPC_A, F.UPC_E, F.CODE_128, F.CODE_39, F.CODE_93, F.ITF, F.CODABAR, F.QR_CODE]);
  hints.set(Z.DecodeHintType.TRY_HARDER, true);
  const r = new Z.MultiFormatReader();
  r.setHints(hints);
  return r;
}
function zxingDecodeCanvas(reader, canvas) {
  const Z = window.ZXing;
  const lum = new Z.HTMLCanvasElementLuminanceSource(canvas);
  const bmp = new Z.BinaryBitmap(new Z.HybridBinarizer(lum));
  return reader.decode(bmp).getText();
}

function stopCam() {
  running = false;
  if (timer) { clearTimeout(timer); timer = null; }
  if (stream) { stream.getTracks().forEach((t) => t.stop()); stream = null; }
  video.srcObject = null;
  wrap.style.display = "none";
  fit();
}

function found(code) {
  stopCam();
  try { if (navigator.vibrate) navigator.vibrate(150); } catch (e) {}
  say("✅ Código detectado: " + code, "ok");
  emit(code);
}

function loopNative() {
  const tick = async () => {
    if (!running) return;
    try {
      const r = await nativeDet.detect(video);
      if (r && r.length) { found(r[0].rawValue); return; }
    } catch (e) {}
    timer = setTimeout(tick, 120);
  };
  tick();
}

function loopZXing() {
  const reader = makeReader();
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  let n = 0;
  const tick = () => {
    if (!running) return;
    if (video.readyState >= 2 && video.videoWidth) {
      const vw = video.videoWidth, vh = video.videoHeight;
      // alterna entre la franja central y el cuadro completo
      const full = (n++ % 3 === 2);
      const cw = Math.floor(vw * (full ? 1 : 0.8)), ch = Math.floor(vh * (full ? 1 : 0.44));
      const sx = Math.floor((vw - cw) / 2), sy = Math.floor((vh - ch) / 2);
      canvas.width = cw; canvas.height = ch;
      ctx.drawImage(video, sx, sy, cw, ch, 0, 0, cw, ch);
      try { const t = zxingDecodeCanvas(reader, canvas); if (t) { found(t); return; } } catch (e) {}
    }
    timer = setTimeout(tick, 100);
  };
  tick();
}

async function startCam() {
  if (running) return;
  say("Abriendo cámara…");
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
    say("⚠️ Esta página no puede usar la cámara en vivo (se necesita conexión https). Usa «Tomar foto del código».", "err");
    return;
  }
  try {
    stream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: "environment" }, width: { ideal: 1280 }, height: { ideal: 720 } },
      audio: false
    });
  } catch (e) {
    say("⚠️ No se pudo abrir la cámara (" + (e.name || e) + "). Revisa el permiso del navegador o usa «Tomar foto del código».", "err");
    return;
  }
  try {
    const tr = stream.getVideoTracks()[0];
    const caps = tr.getCapabilities ? tr.getCapabilities() : {};
    if (caps.focusMode && caps.focusMode.includes("continuous")) {
      await tr.applyConstraints({ advanced: [{ focusMode: "continuous" }] });
    }
  } catch (e) {}
  video.srcObject = stream;
  try { await video.play(); } catch (e) {}
  wrap.style.display = "block";
  fit();
  running = true;
  nativeDet = await getNative();
  if (nativeDet) {
    say("Apunta al código de barras y mantenlo quieto…");
    loopNative();
  } else {
    say("Cargando lector…");
    try { await ensureZXing(); }
    catch (e) { stopCam(); say("⚠️ No se pudo cargar el lector (revisa tu internet). Usa «Tomar foto del código».", "err"); return; }
    say("Apunta al código de barras y mantenlo quieto…");
    loopZXing();
  }
}

async function decodeFile(file) {
  const url = URL.createObjectURL(file);
  const img = await new Promise((res, rej) => {
    const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = url;
  });
  const det = await getNative();
  if (det) {
    try { const r = await det.detect(img); if (r && r.length) return r[0].rawValue; } catch (e) {}
  }
  await ensureZXing();
  const reader = makeReader();
  const scale = Math.min(1, 1600 / Math.max(img.naturalWidth, img.naturalHeight));
  const w = Math.max(1, Math.round(img.naturalWidth * scale)), h = Math.max(1, Math.round(img.naturalHeight * scale));
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  const intentos = [[0, 0, 1, 1], [0, 0.25, 1, 0.5], [0.1, 0.1, 0.8, 0.8]];
  for (const [fx, fy, fw, fh] of intentos) {
    canvas.width = Math.round(w * fw); canvas.height = Math.round(h * fh);
    ctx.drawImage(img, fx * img.naturalWidth, fy * img.naturalHeight, fw * img.naturalWidth, fh * img.naturalHeight,
                  0, 0, canvas.width, canvas.height);
    try { const t = zxingDecodeCanvas(reader, canvas); if (t) return t; } catch (e) {}
  }
  return null;
}

$("btnCam").addEventListener("click", startCam);
$("btnStop").addEventListener("click", () => { stopCam(); say(""); });
$("foto").addEventListener("change", async (ev) => {
  const f = ev.target.files && ev.target.files[0];
  if (!f) return;
  say("Leyendo la foto…");
  try {
    const code = await decodeFile(f);
    if (code) found(code);
    else say("⚠️ No se detectó ningún código en la foto. Acércate más, con buena luz y bien enfocado.", "err");
  } catch (e) {
    say("⚠️ No se pudo leer la foto (" + (e.message || e) + ").", "err");
  }
  ev.target.value = "";
});
window.addEventListener("pagehide", stopCam);
</script>
</body>
</html>
'''


def _declarar_componente_escaner():
    """Crea (si hace falta) la carpeta del componente de cámara y lo registra en Streamlit."""
    try:
        import streamlit.components.v1 as components
        carpeta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "componente_escaner")
        os.makedirs(carpeta, exist_ok=True)
        ruta = os.path.join(carpeta, "index.html")
        try:
            with open(ruta, encoding="utf-8") as f:
                actual = f.read()
        except Exception:
            actual = None
        if actual != _HTML_ESCANER:
            with open(ruta, "w", encoding="utf-8") as f:
                f.write(_HTML_ESCANER)
        return components.declare_component("minici_barcode_scanner", path=carpeta)
    except Exception:
        return None


_scanner_component = _declarar_componente_escaner()


def escanear_codigo_barras(key_suffix, label="📷 Escanear código de barras con la cámara"):
    """
    Botón de cámara para leer un código de barras (celular o computadora).
    El código leído se guarda en st.session_state[f"codigo_{key_suffix}"] y el campo
    de texto con esa misma clave se llena solo, SIN recargar la página.
    Los lectores físicos USB/Bluetooth no necesitan esto: escriben directo en el campo.
    """
    res = None
    with st.expander(label):
        if _scanner_component is None:
            st.warning("No se pudo preparar el escáner. Escribe el código a mano.")
        else:
            res = _scanner_component(key=f"scanner_{key_suffix}", default=None)
    if isinstance(res, dict) and res.get("code"):
        ts = res.get("ts")
        if st.session_state.get(f"scan_ts_{key_suffix}") != ts:
            st.session_state[f"scan_ts_{key_suffix}"] = ts
            st.session_state[f"codigo_{key_suffix}"] = str(res["code"]).strip()
            st.toast(f"✅ Código escaneado: {res['code']}")


# ==========================================
# TALLAS DE CAJAS DE EMPRENDEDORES
# ==========================================
# Precio fijo en dólares por talla. La app guarda el precio en colones (₡),
# convertido con el tipo de cambio que se indique al asignar la caja.
TALLAS_CAJA_USD = {"S": 135.0, "M": 185.0, "L": 220.0, "XL": 265.0}
TIPO_CAMBIO_DEFAULT = 520.0  # ₡ por cada US$1 (ajustable en pantalla)


def talla_desde_nombre(nombre_caja):
    """Devuelve la talla (S/M/L/XL) si el nombre guardado es 'Talla X', si no None."""
    if nombre_caja:
        txt = str(nombre_caja).strip().upper().replace("TALLA", "").strip()
        if txt in TALLAS_CAJA_USD:
            return txt
    return None


def selector_talla_caja(key_prefix, nombre_actual=None):
    """Muestra selector de talla + tipo de cambio y devuelve (nombre_caja, precio_crc)."""
    tallas = list(TALLAS_CAJA_USD.keys())
    actual = talla_desde_nombre(nombre_actual)
    c1, c2 = st.columns(2)
    with c1:
        talla = st.selectbox(
            "Talla de la caja", tallas,
            index=tallas.index(actual) if actual else 1,
            format_func=lambda t: f"{t} — US${TALLAS_CAJA_USD[t]:,.0f}",
            key=f"{key_prefix}_talla",
        )
    with c2:
        tc = st.number_input("Tipo de cambio (₡ por US$1)", min_value=1.0,
                             value=float(TIPO_CAMBIO_DEFAULT), step=1.0,
                             key=f"{key_prefix}_tc")
    usd = TALLAS_CAJA_USD[talla]
    crc = round(usd * tc)
    st.caption(f"💵 Talla {talla}: US${usd:,.0f} × ₡{tc:,.0f} = **₡{crc:,.0f}**")
    if nombre_actual and not actual:
        st.caption(f"ℹ️ Caja actual (sin talla): {nombre_actual}. Al guardar se reemplaza por la talla elegida.")
    return f"Talla {talla}", float(crc)


# -------------------------------------------------------------
# 3. ESTILOS CSS GLOBALES
# -------------------------------------------------------------
_card_seq = [0]


def card(kind="card"):
    """Tarjeta real (contenedor con borde). 'sub' = tarjeta interna más ligera."""
    _card_seq[0] += 1
    try:
        return st.container(border=True, key=f"{kind}_{_card_seq[0]}")
    except TypeError:  # versiones antiguas de Streamlit sin parámetro key
        return st.container(border=True)


st.markdown(
    """
<style>
   @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&display=swap');

   :root {
       --brand: #F3B2C9;
       --brand-soft: #FDE8F0;
       --raspberry: #BE185D;
       --raspberry-dark: #9D174D;
       --plum: #831843;
       --ink: #2B1B24;
       --muted: #7C6572;
       --line: #F3D2E0;
       --paper: #FFF6FA;
   }

   /* ---------- Base ---------- */
   .stApp {
       background-color: var(--paper) !important;
       color: var(--ink) !important;
   }
   html, body, .stApp, p, label, li, input, textarea, button,
   [data-testid="stMarkdownContainer"] {
       font-family: 'DM Sans', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
   }
   h1, h2, h3, h4, h5, h6 {
       font-family: 'Fraunces', Georgia, serif !important;
       letter-spacing: -0.01em;
   }

   #MainMenu, footer, header {visibility: hidden;}

   .block-container {
       padding-top: 1.6rem !important;
       padding-bottom: 3rem !important;
   }

   label, p, span, div, h1, h2, h3, h4, h5, h6 {
       color: var(--ink);
   }
   [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] * {
       color: var(--muted) !important;
   }
   hr {
       border-color: var(--line) !important;
   }

   /* ---------- Encabezado de la app ---------- */
   .app-header {
       display: flex;
       align-items: baseline;
       flex-wrap: wrap;
       gap: 4px 12px;
       padding-top: 4px;
   }
   .app-brand {
       font-family: 'Fraunces', Georgia, serif;
       font-size: 28px;
       font-weight: 700;
       color: var(--raspberry) !important;
       line-height: 1.15;
   }
   .app-role {
       font-size: 14px;
       color: var(--muted) !important;
   }

   .login-hero {
       text-align: center;
       margin: 6px 0 18px;
   }
   .login-hero h2 {
       color: var(--raspberry) !important;
       font-size: 28px;
       margin-bottom: 4px;
   }
   .login-hero p {
       color: var(--muted) !important;
       font-size: 15px;
       margin: 0;
   }

   /* ---------- Banner de sección ---------- */
   .top-banner {
       background: #ffffff;
       border: 1px solid var(--line);
       border-left: 6px solid var(--brand);
       padding: 16px 20px;
       border-radius: 16px;
       margin-bottom: 18px;
       box-shadow: 0 1px 3px rgba(131, 24, 67, 0.06);
   }
   .top-banner * {
       color: var(--plum) !important;
       font-family: 'Fraunces', Georgia, serif;
   }

   .section-title {
       font-size: 13.5px;
       font-weight: 600;
       color: var(--raspberry-dark) !important;
       margin: 12px 0 4px;
   }

   /* ---------- Tarjetas (st.container con borde) ---------- */
   div[class*="st-key-card_"],
   div[data-testid="stVerticalBlockBorderWrapper"]:has(> div[class*="st-key-card_"]) {
       background: #ffffff !important;
       border: 1px solid var(--line) !important;
       border-radius: 18px !important;
       box-shadow: 0 4px 18px rgba(190, 24, 93, 0.06) !important;
   }
   div[class*="st-key-sub_"],
   div[data-testid="stVerticalBlockBorderWrapper"]:has(> div[class*="st-key-sub_"]) {
       background: #FFF9FC !important;
       border: 1px solid var(--line) !important;
       border-radius: 14px !important;
       box-shadow: none !important;
   }
   /* evita doble borde cuando Streamlit envuelve el contenedor */
   div[data-testid="stVerticalBlockBorderWrapper"] > div[class*="st-key-card_"],
   div[data-testid="stVerticalBlockBorderWrapper"] > div[class*="st-key-sub_"] {
       background: transparent !important;
       border: none !important;
       box-shadow: none !important;
       border-radius: 0 !important;
   }

   /* ---------- Métricas ---------- */
   div[data-testid="stMetric"] {
       background-color: #ffffff !important;
       padding: 14px 16px !important;
       border-radius: 16px !important;
       border: 1px solid var(--line) !important;
       box-shadow: none !important;
   }
   div[data-testid="stMetricLabel"] > div,
   div[data-testid="stMetricLabel"] label,
   div[data-testid="stMetricLabel"] p {
       color: var(--raspberry-dark) !important;
       font-weight: 600 !important;
       font-size: 13px !important;
   }
   div[data-testid="stMetricValue"] > div {
       color: var(--plum) !important;
       font-family: 'Fraunces', Georgia, serif !important;
       font-weight: 700 !important;
   }

   /* ---------- Campos de formulario ---------- */
   div[data-baseweb="select"],
   div[data-baseweb="select"] *,
   div[data-baseweb="input"],
   div[data-baseweb="input"] *,
   div[data-baseweb="base-input"],
   div[data-baseweb="base-input"] *,
   .stSelectbox > div > div,
   .stTextInput > div > div,
   .stNumberInput > div > div {
       background-color: #ffffff !important;
       color: var(--ink) !important;
       fill: var(--ink) !important;
   }
   div[data-baseweb="input"],
   div[data-baseweb="select"] > div,
   div[data-baseweb="textarea"] {
       border-radius: 10px !important;
       border: 1px solid #EDC3D5 !important;
   }
   div[data-baseweb="input"]:focus-within,
   div[data-baseweb="select"] > div:focus-within,
   div[data-baseweb="textarea"]:focus-within {
       border-color: var(--raspberry) !important;
       box-shadow: 0 0 0 3px rgba(190, 24, 93, 0.15) !important;
   }

   input, textarea, [role="option"], [role="combobox"] {
       color: var(--ink) !important;
       -webkit-text-fill-color: var(--ink) !important;
       background-color: #ffffff !important;
   }

   div[data-baseweb="popover"],
   div[data-baseweb="popover"] * {
       background-color: #ffffff !important;
       color: var(--ink) !important;
   }

   .stNumberInput button {
       background-color: var(--brand-soft) !important;
       color: var(--plum) !important;
       border: none !important;
   }

   /* ---------- Botones ---------- */
   div.stButton > button:first-child {
       background-color: var(--raspberry) !important;
       color: #ffffff !important;
       border-radius: 12px;
       font-weight: 600;
       border: none;
       padding: 10px 18px;
       width: 100%;
       font-size: 15px;
       box-shadow: 0 2px 8px rgba(190, 24, 93, 0.25);
       transition: background-color 0.15s ease, transform 0.05s ease;
   }
   div.stButton > button:first-child * {
       color: #ffffff !important;
   }
   div.stButton > button:first-child:hover {
       background-color: var(--raspberry-dark) !important;
   }
   div.stButton > button:first-child:active {
       transform: translateY(1px);
   }

   div.stDownloadButton > button:first-child {
       background-color: var(--brand-soft) !important;
       color: var(--plum) !important;
       border-radius: 12px;
       font-weight: 600;
       border: 1px solid var(--brand) !important;
       padding: 10px 18px;
       width: 100%;
       font-size: 15px;
   }
   div.stDownloadButton > button:first-child * {
       color: var(--plum) !important;
   }
   div.stDownloadButton > button:first-child:hover {
       background-color: var(--brand) !important;
   }

   /* Botones secundarios (salir / eliminar) */
   .st-key-btn_salir_admin div.stButton > button:first-child,
   .st-key-btn_salir_client div.stButton > button:first-child,
   .st-key-btn_eliminar_pedido div.stButton > button:first-child,
   .st-key-btn_inv_del div.stButton > button:first-child {
       background-color: #ffffff !important;
       border: 1.5px solid var(--brand) !important;
       box-shadow: none !important;
   }
   .st-key-btn_salir_admin div.stButton > button:first-child *,
   .st-key-btn_salir_client div.stButton > button:first-child *,
   .st-key-btn_eliminar_pedido div.stButton > button:first-child *,
   .st-key-btn_inv_del div.stButton > button:first-child * {
       color: var(--raspberry) !important;
   }
   .st-key-btn_salir_admin div.stButton > button:first-child:hover,
   .st-key-btn_salir_client div.stButton > button:first-child:hover,
   .st-key-btn_eliminar_pedido div.stButton > button:first-child:hover,
   .st-key-btn_inv_del div.stButton > button:first-child:hover {
       background-color: var(--brand-soft) !important;
   }

   .btn-whatsapp {
       display: inline-block;
       width: 100%;
       background-color: #25D366 !important;
       color: #ffffff !important;
       text-align: center;
       font-weight: 600;
       padding: 10px 15px;
       border-radius: 12px;
       text-decoration: none;
       font-size: 15px;
       margin-top: 10px;
       box-shadow: 0 3px 8px rgba(37, 211, 102, 0.3);
   }
   .btn-whatsapp:hover {
       background-color: #1da851 !important;
       color: #ffffff !important;
   }

   /* ---------- Menús tipo pastilla ---------- */
   div[data-testid="stRadio"] > div {
       flex-direction: row !important;
       gap: 6px !important;
       flex-wrap: wrap !important;
   }
   div[data-testid="stRadio"] label:not([data-testid="stWidgetLabel"]) {
       background-color: #ffffff !important;
       border: 1.5px solid var(--line) !important;
       padding: 6px 14px !important;
       border-radius: 999px !important;
       font-weight: 600 !important;
       cursor: pointer !important;
       transition: border-color 0.15s ease, background-color 0.15s ease;
   }
   div[data-testid="stRadio"] label:not([data-testid="stWidgetLabel"]):hover {
       border-color: var(--brand) !important;
   }
   div[data-testid="stRadio"] label:not([data-testid="stWidgetLabel"]) p,
   div[data-testid="stRadio"] label:not([data-testid="stWidgetLabel"]) span,
   div[data-testid="stRadio"] label:not([data-testid="stWidgetLabel"]) div {
       color: var(--raspberry) !important;
       -webkit-text-fill-color: var(--raspberry) !important;
       font-size: 13.5px !important;
   }
   div[data-testid="stRadio"] label:has(input:checked) {
       background-color: var(--raspberry) !important;
       border-color: var(--raspberry) !important;
   }
   div[data-testid="stRadio"] label:has(input:checked) p,
   div[data-testid="stRadio"] label:has(input:checked) span,
   div[data-testid="stRadio"] label:has(input:checked) div {
       color: #ffffff !important;
       -webkit-text-fill-color: #ffffff !important;
   }
   div[data-testid="stRadio"] input[type="radio"] {
       display: none !important;
   }

   /* ---------- Tablas, alertas, expanders, imágenes ---------- */
   div[data-testid="stDataFrame"] {
       border: 1px solid var(--line);
       border-radius: 12px;
       overflow: hidden;
   }
   div[data-testid="stAlert"] {
       border-radius: 12px !important;
   }
   div[data-testid="stExpander"] {
       background: #ffffff !important;
       border: 1px solid var(--line) !important;
       border-radius: 14px !important;
   }
   div[data-testid="stImage"] img {
       border-radius: 12px;
   }
   div[data-testid="stForm"] {
       border: 1px solid var(--line) !important;
       border-radius: 14px !important;
       background: #FFFBFD !important;
   }
</style>
""",
    unsafe_allow_html=True,
)

# -------------------------------------------------------------
# 4. CONTROL DE SESIÓN Y AUTENTICACIÓN
# -------------------------------------------------------------
if "user_role" not in st.session_state:
    st.session_state.user_role = None
if "current_client" not in st.session_state:
    st.session_state.current_client = None

# --- VISTA DE LOGIN ---
if st.session_state.user_role is None:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        logo_path = None
        for posible_nombre in ["logo.jpg", "logo.png", "logo.jpeg", "21237.jpg"]:
            if os.path.exists(posible_nombre):
                logo_path = posible_nombre
                break

        if logo_path:
            st.image(logo_path)
        else:
            st.markdown(
                "<h1 style='text-align: center; color: #be185d;'>🛍️ Minici Store</h1>",
                unsafe_allow_html=True,
            )

    st.markdown(
        """
        <div class="login-hero">
            <h2>¡Bienvenidos a Minici Store!</h2>
            <p>Ingresa tu código de acceso para continuar.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with card():
        codigo_ingresado = st.text_input("Código de acceso", placeholder="Ej. MIN-0001 o EMP-0001", key="login_input")

        st.write("")
        if st.button("🚀 Ingresar al Sistema", key="btn_login"):
            codigo_limpio = codigo_ingresado.strip().upper()
        
            if codigo_limpio == str(_secret("ADMIN_CODE", "KENDRA5412")).strip().upper():
                st.session_state.user_role = "admin"
                st.rerun()
            else:
                c.execute("SELECT id_cliente FROM clientes WHERE UPPER(id_cliente) = ?", (codigo_limpio,))
                res = c.fetchone()
                if res:
                    st.session_state.user_role = "client"
                    st.session_state.current_client = res[0]
                    st.rerun()
                else:
                    st.error("❌ Código incorrecto o no registrado.")

# -------------------------------------------------------------
# 5. PANEL DE ADMINISTRADOR
# -------------------------------------------------------------
elif st.session_state.user_role == "admin":
    col_a, col_save, col_b = st.columns([3, 1.5, 1])
    with col_a:
        st.markdown(
            '<div class="app-header"><span class="app-brand">Minici Store</span><span class="app-role">Panel administrador</span></div>',
            unsafe_allow_html=True,
        )
    with col_save:
        if st.button("💾 Guardar datos", key="btn_respaldo"):
            with st.spinner("Guardando todos los datos..."):
                try:
                    st.session_state["_respaldo"] = hacer_respaldo()
                except Exception as e:
                    st.session_state["_respaldo"] = {"error": str(e)}
    with col_b:
        if st.button("🚪 Salir", key="btn_salir_admin"):
            st.session_state.user_role = None
            st.rerun()

    mostrar_flash()
    if st.session_state.pop("_restaurado_gh", False):
        st.info("♻️ El servidor se había reiniciado: tus datos se recuperaron automáticamente del último respaldo de GitHub.")

    _rsp = st.session_state.get("_respaldo")
    if _rsp:
        if _rsp.get("error"):
            st.error(f"No se pudo crear el respaldo: {_rsp['error']}")
        else:
            if _rsp["github"] is True:
                st.success(f"✅ Todos los datos quedaron guardados en GitHub ({_rsp['hora']}).")
            elif _rsp["github"] is False:
                st.warning(f"⚠️ El respaldo se creó pero NO se pudo subir a GitHub: {_rsp['detalle']}")
            else:
                st.info(f"💾 Respaldo creado ({_rsp['hora']}). Descárgalo abajo. Para que también se guarde solo en GitHub, abre «Respaldos y GitHub».")
            st.download_button("⬇️ Descargar respaldo (datos + fotos + Excel)", data=_rsp["zip"],
                               file_name=_rsp["nombre"], mime="application/zip", key="dl_respaldo")

    with st.expander("⚙️ Respaldos y GitHub (configurar / restaurar)"):
        if _gh_config():
            _cfg = _gh_config()
            st.success(f"GitHub conectado: {_cfg['repo']} (rama {_cfg['branch']}). Si el servidor se reinicia, los datos se recuperan solos.")
        else:
            st.markdown(
                "**Para guardar los datos en GitHub automáticamente:**\n"
                "1. Crea un repositorio **privado** nuevo solo para respaldos (por ejemplo `minici-respaldos`). "
                "Debe ser privado porque contiene nombres y teléfonos de clientas, y distinto al repositorio de la app: "
                "si guardas en el repositorio de la app, cada respaldo reinicia la aplicación.\n"
                "2. En GitHub: *Settings → Developer settings → Fine-grained tokens* → crea un token con acceso solo a ese repositorio "
                "y el permiso **Contents: Read and write**.\n"
                "3. En Streamlit Cloud: *Manage app → Settings → Secrets* y pega:"
            )
            st.code('GITHUB_TOKEN = "github_pat_xxxxxxxx"\nGITHUB_REPO = "tu-usuario/minici-respaldos"\nGITHUB_BRANCH = "main"', language="toml")
        st.divider()
        st.markdown("**♻️ Restaurar un respaldo** (reemplaza todos los datos actuales)")
        _arch = st.file_uploader("Sube un respaldo .zip (o la base .db)", type=["zip", "db"], key="up_restaurar")
        _ok = st.checkbox("Entiendo que esto reemplaza los clientes, pedidos y abonos actuales", key="chk_restaurar")
        if st.button("♻️ Restaurar ahora", key="btn_restaurar", disabled=not (_arch and _ok)):
            _ok_r, _txt_r = restaurar_respaldo(_arch.getvalue())
            if _ok_r:
                flash("success", "✅ " + _txt_r)
                st.rerun()
            else:
                st.error(_txt_r)

    seccion_admin = st.radio(
        "Módulo General",
        ["📦 Panel Principal (Pedidos y Clientes)", "⚡ Módulo Postventa & POS"],
        horizontal=True,
    )
    st.write("")

    # =========================================================
    # SECCIÓN 1: PANEL PRINCIPAL
    # =========================================================
    if seccion_admin == "📦 Panel Principal (Pedidos y Clientes)":
        menu_principal = st.radio(
            "Acción Principal",
            [
                "📸 Registrar Compra",
                "👩 Clientes y Expedientes",
                "✏️ Gestor Pedidos",
                "💰 Registrar Abonos",
                "💸 Gastos Operativos",
                "📊 Finanzas y Reportes"
            ],
            horizontal=True,
            label_visibility="collapsed",
            key="menu_ppal"
        )
        st.write("")

        # 1. Registrar Compra
        if menu_principal == "📸 Registrar Compra":
            st.markdown(
                """
                <div class="top-banner">
                    <div style="font-size: 18px; font-weight: 700;">Registrar Compra / Encargo de Cliente o Emprendedor</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            with card():
            
                filtro_tipo_cli = st.radio("Filtrar Tipo de Cliente:", ["Todos", "Clientas Regulares (MIN)", "Emprendedores (EMP)"], horizontal=True, key="compra_filtro_cli")
            
                query_cli = """SELECT c.id_cliente || ' — ' || c.nombre || COALESCE(' (' || e.nombre_caja || ')', '') AS display, c.id_cliente 
                               FROM clientes c 
                               LEFT JOIN cajas_emprendedores e ON c.id_cliente = e.id_cliente"""
                if filtro_tipo_cli == "Clientas Regulares (MIN)":
                    query_cli += " WHERE c.id_cliente LIKE 'MIN-%'"
                elif filtro_tipo_cli == "Emprendedores (EMP)":
                    query_cli += " WHERE c.id_cliente LIKE 'EMP-%'"

                clientes_df = pd.read_sql(query_cli, conn)

                if clientes_df.empty:
                    st.warning("No hay clientes registrados en esta categoría.")
                else:
                    st.markdown('<p class="section-title">1. Seleccionar Cliente / Emprendedor</p>', unsafe_allow_html=True)
                    cli_selected = st.selectbox("Cliente", clientes_df["display"], label_visibility="collapsed", key="compra_cli")
                    id_cliente = clientes_df[clientes_df["display"] == cli_selected]["id_cliente"].values[0]

                    col_tienda, col_cat = st.columns(2)
                    with col_tienda:
                        st.markdown('<p class="section-title">2. Tienda</p>', unsafe_allow_html=True)
                        tienda_sel = st.selectbox("Tienda", ["Zara", "Guess", "Adidas", "Shein", "Amazon", "Nike", "Victoria's Secret", "Otra"], label_visibility="collapsed", key="compra_tienda")
                        tienda = st.text_input("Escribe tienda", placeholder="Tienda...") if tienda_sel == "Otra" else tienda_sel

                    with col_cat:
                        st.markdown('<p class="section-title">3. Categoría</p>', unsafe_allow_html=True)
                        cat_sel = st.selectbox("Categoría", ["Vestido", "Bolso", "Tenis", "Blusa", "Cosméticos", "Accesorios", "Otra"], label_visibility="collapsed", key="compra_cat")
                        categoria = st.text_input("Escribe categoría", placeholder="Categoría...") if cat_sel == "Otra" else cat_sel

                    st.markdown('<p class="section-title">4. Producto / Artículo</p>', unsafe_allow_html=True)
                    producto = st.text_input("Producto", placeholder="Ej. Vestido estampado", label_visibility="collapsed", key="compra_prod")

                    col_c, col_d = st.columns(2)
                    with col_c:
                        st.markdown('<p class="section-title">5. Precio (₡ CRC)</p>', unsafe_allow_html=True)
                        precio = st.number_input("Precio", min_value=0.0, value=15000.0, step=500.0, label_visibility="collapsed", key="compra_precio")
                    with col_d:
                        st.markdown('<p class="section-title">6. Cantidad</p>', unsafe_allow_html=True)
                        cantidad = st.number_input("Cantidad", min_value=1, value=1, step=1, label_visibility="collapsed", key="compra_cant")

                    st.markdown('<p class="section-title">7. Estado del producto</p>', unsafe_allow_html=True)
                    estado = st.selectbox("Estado", ["🇺🇸 Comprado en USA", "📦 En tránsito", "🇨🇷 Recibido en CR", "✅ Entregado"], label_visibility="collapsed", key="compra_estado")

                    st.markdown('<p class="section-title">8. Fotos del producto</p>', unsafe_allow_html=True)
                    metodo_foto = st.radio("Cargar foto desde:", ["Subir archivo", "Usar cámara"], horizontal=True, key="compra_foto_modo")
                    foto_file = st.file_uploader("Subir foto", type=["jpg", "png", "jpeg"], label_visibility="collapsed", key="compra_file") if metodo_foto == "Subir archivo" else st.camera_input("Tomar foto", key="compra_cam")

                    st.markdown('<p class="section-title">9. Observaciones</p>', unsafe_allow_html=True)
                    observaciones = st.text_area("Observaciones", height=70, label_visibility="collapsed", key="compra_obs")

                    st.write("")
                    if st.button("💾 Guardar compra", key="btn_save_compra"):
                        if not producto or not tienda or not categoria:
                            st.error("Debes completar el producto, la tienda y la categoría.")
                        else:
                            prod_clean = cap1(producto)
                            tienda_clean = cap1(tienda)
                            cat_clean = cap1(categoria)
                            obs_clean = cap1(observaciones) if observaciones else ""

                            foto_filename = ""
                            if foto_file:
                                foto_filename = f"{id_cliente}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
                                filepath = os.path.join("fotos_productos", foto_filename)
                                with open(filepath, "wb") as f:
                                    f.write(foto_file.getbuffer())

                            c.execute("""INSERT INTO productos (id_cliente, tienda, categoria, descripcion, precio, moneda, cantidad, estado, observaciones, foto_path)
                                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                      (id_cliente, tienda_clean, cat_clean, prod_clean, precio, "CRC", cantidad, estado, obs_clean, foto_filename))
                            c.execute("""INSERT INTO notificaciones (id_cliente, titulo, mensaje, fecha) VALUES (?, ?, ?, ?)""",
                                      (id_cliente, "Nuevo pedido registrado", f"Se agregó '{prod_clean}' ({tienda_clean}) a tus compras por ₡{precio:,.0f}.", datetime.now().strftime("%Y-%m-%d %H:%M")))
                            conn.commit()
                            st.success(f"¡Compra de '{prod_clean}' registrada exitosamente para {id_cliente}!")

                            path_fac, link_wa = generar_factura_imagen(id_cliente)
                            if path_fac and os.path.exists(path_fac):
                                st.success("🖼️ ¡Factura digital en imagen generada con éxito!")
                                with open(path_fac, "rb") as file:
                                    st.download_button(
                                        label="📥 Descargar Factura en Imagen (para WhatsApp)",
                                        data=file,
                                        file_name=f"Factura_{id_cliente}.jpg",
                                        mime="image/jpeg",
                                        key="dl_factura_compra"
                                    )
                            if link_wa:
                                st.markdown(f'<a href="{link_wa}" target="_blank" class="btn-whatsapp">📲 Enviar Mensaje por WhatsApp</a>', unsafe_allow_html=True)

        # 2. Clientes y Expedientes
        elif menu_principal == "👩 Clientes y Expedientes":
            st.markdown(
                """
                <div class="top-banner">
                    <div style="font-size: 18px; font-weight: 700;">Registro de Clientes / Emprendedores y Expedientes</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            with card():
                st.markdown("<h4 style='color:#be185d;'>👤 Registrar Nuevo Cliente / Emprendedor</h4>", unsafe_allow_html=True)

                tipo_registro = st.radio("Tipo de Registro:", ["🌸 Clienta Regular (MIN)", "💼 Emprendedor (EMP)"], horizontal=True, key="tipo_reg_cli")

                c.execute("SELECT id_cliente FROM clientes")
                rows = c.fetchall()

                if "EMP" in tipo_registro:
                    numeros = [int(str(r[0]).replace("EMP-", "").strip()) for r in rows if r[0] and str(r[0]).startswith("EMP-") and str(r[0]).replace("EMP-", "").strip().isdigit()]
                    nuevo_id = f"EMP-{max(numeros) + 1:04d}" if numeros else "EMP-0001"
                else:
                    numeros = [int(str(r[0]).replace("MIN-", "").strip()) for r in rows if r[0] and str(r[0]).startswith("MIN-") and str(r[0]).replace("MIN-", "").strip().isdigit()]
                    nuevo_id = f"MIN-{max(numeros) + 1:04d}" if numeros else "MIN-0001"

                st.markdown(f"🏷️ <span style='font-size:16px; font-weight:bold; color:#be185d;'>Código asignado: {nuevo_id}</span>", unsafe_allow_html=True)
                st.write("")

                with st.form("form_registro_cliente", clear_on_submit=True):
                    nombre = st.text_input("Nombre Completo", placeholder="Ej. Maria Lopez")
                    tel = st.text_input("Teléfono / WhatsApp", placeholder="Ej. 88888888")
                    correo = st.text_input("Correo Electrónico", placeholder="Ej. correo@ejemplo.com")

                    nombre_caja_in, precio_caja_in = "", 0.0
                    if "EMP" in tipo_registro:
                        st.markdown("---")
                        st.markdown("💼 **Configuración de Caja para Emprendedor:**")
                        nombre_caja_in, precio_caja_in = selector_talla_caja("reg_caja")

                    btn_guardar_cli = st.form_submit_button("Guardar Registro")

                    if btn_guardar_cli:
                        if nombre.strip():
                            nombre_clean = nombre.strip().title()
                            c.execute("INSERT OR REPLACE INTO clientes (id_cliente, nombre, telefono, correo) VALUES (?, ?, ?, ?)", (nuevo_id, nombre_clean, tel, correo))
                        
                            if "EMP" in tipo_registro:
                                caja_clean = nombre_caja_in or "Caja Asignada"
                                c.execute("INSERT OR REPLACE INTO cajas_emprendedores (id_cliente, nombre_caja, precio_caja) VALUES (?, ?, ?)", (nuevo_id, caja_clean, precio_caja_in))

                            conn.commit()
                            flash("success", f"¡Registro de {nombre_clean} guardado exitosamente ({nuevo_id})!")
                            st.rerun()
                        else:
                            st.error("Debes ingresar el nombre del cliente.")

            with card():
                st.markdown("<h4 style='color:#be185d;'>📋 Directorio de Clientes y Emprendedores</h4>", unsafe_allow_html=True)
            
                clientas_todas = pd.read_sql("""
                    SELECT c.id_cliente as Código, c.nombre as Nombre, c.telefono as Teléfono, 
                           COALESCE(e.nombre_caja, 'N/A') as 'Caja', COALESCE(e.precio_caja, 0.0) as 'Precio Caja (₡)'
                    FROM clientes c 
                    LEFT JOIN cajas_emprendedores e ON c.id_cliente = e.id_cliente 
                    ORDER BY c.id_cliente DESC""", conn)
            
                id_cli_exp = None
                if not clientas_todas.empty:
                    evento_seleccion = st.dataframe(clientas_todas, hide_index=True, on_select="rerun", selection_mode="single-row", key="df_dir_clientas")

                    selected_rows = []
                    if isinstance(evento_seleccion, dict):
                        selected_rows = evento_seleccion.get("selection", {}).get("rows", [])
                    elif hasattr(evento_seleccion, "selection"):
                        sel = getattr(evento_seleccion, "selection")
                        if isinstance(sel, dict):
                            selected_rows = sel.get("rows", [])
                        elif hasattr(sel, "rows"):
                            selected_rows = getattr(sel, "rows", [])

                    if selected_rows:
                        fila_idx = selected_rows[0]
                        id_cli_exp = clientas_todas.iloc[fila_idx]["Código"]

                    st.divider()
                    st.markdown("<h5 style='color:#be185d;'>🔍 Expediente del Cliente Seleccionado</h5>", unsafe_allow_html=True)

                    if id_cli_exp:
                        c.execute("SELECT nombre, telefono, correo FROM clientes WHERE id_cliente = ?", (id_cli_exp,))
                        info_cli = c.fetchone()

                        c.execute("SELECT nombre_caja, precio_caja FROM cajas_emprendedores WHERE id_cliente = ?", (id_cli_exp,))
                        res_caja_exp = c.fetchone()
                        caja_nombre_exp, caja_precio_exp = (res_caja_exp[0], res_caja_exp[1] or 0.0) if res_caja_exp else (None, 0.0)

                        prods_cli = pd.read_sql("SELECT descripcion, tienda, precio, cantidad, estado, foto_path FROM productos WHERE id_cliente = ?", conn, params=(id_cli_exp,))
                        abonos_cli = pd.read_sql("SELECT monto_crc, fecha FROM abonos WHERE id_cliente = ?", conn, params=(id_cli_exp,))
                    
                        tot_articulos = (prods_cli['precio'].fillna(0.0) * prods_cli['cantidad'].fillna(1)).sum() if not prods_cli.empty else 0.0
                        tot_comp = tot_articulos + caja_precio_exp
                        tot_ab = abonos_cli['monto_crc'].fillna(0.0).sum() if not abonos_cli.empty else 0.0
                        saldo_p = tot_comp - tot_ab

                        with st.expander(f"👤 EXPEDIENTE: {info_cli[0]} ({id_cli_exp})", expanded=True):
                            col_m1, col_m2, col_m3 = st.columns(3)
                            col_m1.metric("Total Cargos", f"₡{tot_comp:,.0f}")
                            col_m2.metric("Total Abonado", f"₡{tot_ab:,.0f}")
                            col_m3.metric("Saldo Pendiente", f"₡{max(0.0, saldo_p):,.0f}")

                            if id_cli_exp.startswith("EMP-"):
                                st.write("")
                                with card("sub"):
                                    st.markdown(f"💼 **Configuración de Caja de Emprendedor:**")
                                    edit_nombre_caja, edit_precio_caja = selector_talla_caja(f"exp_cj_{id_cli_exp}", caja_nombre_exp)
                                    if st.button("💾 Actualizar Caja", key=f"btn_update_cj_{id_cli_exp}"):
                                        c.execute("INSERT OR REPLACE INTO cajas_emprendedores (id_cliente, nombre_caja, precio_caja) VALUES (?, ?, ?)",
                                                  (id_cli_exp, edit_nombre_caja, edit_precio_caja))
                                        conn.commit()
                                        flash("success", "¡Caja actualizada!")
                                        st.rerun()

                            st.write("")
                            path_fac, link_wa = generar_factura_imagen(id_cli_exp)
                            if path_fac and os.path.exists(path_fac):
                                with open(path_fac, "rb") as file:
                                    st.download_button(
                                        label="🖼️ Descargar Factura en Imagen (para WhatsApp)",
                                        data=file,
                                        file_name=f"Factura_{id_cli_exp}.jpg",
                                        mime="image/jpeg",
                                        key="dl_factura_exp"
                                    )
                            if link_wa:
                                st.markdown(f'<a href="{link_wa}" target="_blank" class="btn-whatsapp">📲 Enviar Mensaje por WhatsApp</a>', unsafe_allow_html=True)

                            st.divider()
                            st.markdown("**📦 Pedidos / Artículos comprados:**")
                            if prods_cli.empty:
                                st.info("Sin artículos individuales registrados.")
                            else:
                                for _, r in prods_cli.iterrows():
                                    with card("sub"):
                                        col_img, col_info = st.columns([1, 2])
                                        with col_img:
                                            if r["foto_path"] and os.path.exists(os.path.join("fotos_productos", r["foto_path"])):
                                                st.image(os.path.join("fotos_productos", r["foto_path"]))
                                            else:
                                                st.caption("📷 Sin foto")
                                        with col_info:
                                            st.markdown(f"✨ **{r['descripcion']}**")
                                            st.caption(f"🛍️ Tienda: {r['tienda']} | 📌 Estado: {r['estado']}")
                                            precio_item = r['precio'] if pd.notna(r['precio']) else 0.0
                                            cant_item = r['cantidad'] if pd.notna(r['cantidad']) else 1
                                            st.markdown(f"📦 Cantidad: {int(cant_item)} | 💰 **Total: ₡{(precio_item * cant_item):,.0f}**")
                    else:
                        st.info("👆 Haz clic en cualquier cliente de la tabla superior para cargar automáticamente su expediente.")
                else:
                    st.info("No hay clientes registrados.")

        # 3. Gestor de pedidos: al elegir un pedido (en la tabla o en la lista) sus datos se cargan solos
        elif menu_principal == "✏️ Gestor Pedidos":
            st.markdown(
                """
                <div class="top-banner">
                    <div style="font-size: 18px; font-weight: 700;">Gestión y Edición de Pedidos por Cliente</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            mensaje_flash = st.session_state.pop("gestor_msg", None)
            if mensaje_flash:
                st.success(mensaje_flash)

            cli_pedidos = pd.read_sql("SELECT DISTINCT p.id_cliente, c.nombre FROM productos p JOIN clientes c ON p.id_cliente = c.id_cliente WHERE p.id_cliente IS NOT NULL AND p.id_cliente != ''", conn)

            with card():
                if cli_pedidos.empty:
                    st.info("No hay encargos asignados a clientes específicos.")
                else:
                    dict_cli_p = {f"{r['id_cliente']} — {r['nombre']}": r['id_cliente'] for _, r in cli_pedidos.iterrows()}
                    if st.session_state.get("gestor_cli_sel") not in dict_cli_p:
                        st.session_state.pop("gestor_cli_sel", None)
                    cli_p_sel = st.selectbox("Cliente", list(dict_cli_p.keys()), key="gestor_cli_sel")
                    id_cli_gest = dict_cli_p[cli_p_sel]

                    prods_de_cli = pd.read_sql(
                        "SELECT id, tienda, categoria, descripcion, precio, cantidad, estado, foto_path FROM productos WHERE id_cliente = ? ORDER BY id DESC",
                        conn, params=(id_cli_gest,))

                    if prods_de_cli.empty:
                        st.warning("Este cliente no tiene pedidos registrados.")
                    else:
                        ids_pedidos = [int(i) for i in prods_de_cli["id"]]
                        etiquetas = {
                            int(r["id"]): f"Pedido #{int(r['id'])} - {r['descripcion']} (₡{(r['precio'] if pd.notna(r['precio']) else 0.0):,.0f})"
                            for _, r in prods_de_cli.iterrows()
                        }

                        # --- Tabla: al tocar una fila se selecciona ese pedido ---
                        st.caption("Toca una fila de la tabla (o usa la lista de abajo) para cargar el pedido en el formulario.")
                        tabla_pedidos = prods_de_cli[["id", "tienda", "descripcion", "cantidad", "precio", "estado"]].rename(
                            columns={"id": "N.º", "tienda": "Tienda", "descripcion": "Producto", "cantidad": "Cant.", "precio": "Precio (₡)", "estado": "Estado"})
                        evento_ped = st.dataframe(
                            tabla_pedidos, hide_index=True, on_select="rerun", selection_mode="single-row",
                            key=f"gestor_tabla_{id_cli_gest}")

                        filas_sel = []
                        sel_obj = getattr(evento_ped, "selection", None)
                        if sel_obj is None and isinstance(evento_ped, dict):
                            sel_obj = evento_ped.get("selection")
                        if isinstance(sel_obj, dict):
                            filas_sel = sel_obj.get("rows", [])
                        elif sel_obj is not None:
                            filas_sel = getattr(sel_obj, "rows", [])

                        if filas_sel and filas_sel[0] < len(ids_pedidos):
                            marca_fila = (id_cli_gest, ids_pedidos[filas_sel[0]])
                            if st.session_state.get("gestor_ultima_fila") != marca_fila:
                                st.session_state.gestor_ultima_fila = marca_fila
                                st.session_state.gestor_pedido_id = ids_pedidos[filas_sel[0]]
                        else:
                            st.session_state.pop("gestor_ultima_fila", None)

                        # --- Lista desplegable (sincronizada con la tabla) ---
                        if st.session_state.get("gestor_pedido_id") not in ids_pedidos:
                            st.session_state.gestor_pedido_id = ids_pedidos[0]

                        id_pedido_actual = st.selectbox(
                            "Pedido a modificar",
                            ids_pedidos,
                            format_func=lambda i: etiquetas.get(i, f"Pedido #{i}"),
                            key="gestor_pedido_id",
                        )

                        # --- Carga automática de los datos del pedido en el formulario ---
                        fila_actual = prods_de_cli[prods_de_cli["id"] == id_pedido_actual].iloc[0]
                        if st.session_state.get("gestor_cargado_id") != id_pedido_actual:
                            st.session_state.edit_tienda = fila_actual["tienda"] if pd.notna(fila_actual["tienda"]) else ""
                            st.session_state.edit_categoria = fila_actual["categoria"] if pd.notna(fila_actual["categoria"]) else ""
                            st.session_state.edit_desc = fila_actual["descripcion"] if pd.notna(fila_actual["descripcion"]) else ""
                            st.session_state.edit_cant = int(fila_actual["cantidad"]) if pd.notna(fila_actual["cantidad"]) else 1
                            st.session_state.edit_precio = float(fila_actual["precio"]) if pd.notna(fila_actual["precio"]) else 0.0
                            st.session_state.edit_estado = fila_actual["estado"] if pd.notna(fila_actual["estado"]) and fila_actual["estado"] else "Pendiente"
                            st.session_state.gestor_cargado_id = id_pedido_actual

                        st.divider()
                        st.markdown(f"<h5 style='color:#be185d;'>Editando el pedido #{id_pedido_actual}</h5>", unsafe_allow_html=True)

                        foto_pedido = fila_actual["foto_path"]
                        if isinstance(foto_pedido, str) and foto_pedido and os.path.exists(os.path.join("fotos_productos", foto_pedido)):
                            st.image(os.path.join("fotos_productos", foto_pedido), width=150)

                        col1, col2 = st.columns(2)

                        with col1:
                            nueva_tienda = st.text_input("Tienda", key="edit_tienda")
                            nueva_descripcion = st.text_input("Descripción", key="edit_desc")
                            nuevo_precio = st.number_input("Precio (₡)", step=100.0, key="edit_precio")

                        with col2:
                            nueva_categoria = st.text_input("Categoría", key="edit_categoria")
                            nueva_cantidad = st.number_input("Cantidad", step=1, key="edit_cant")

                            opciones_estado = ["Pendiente", "🇺🇸 Comprado en USA", "📦 En tránsito", "🇨🇷 Recibido en CR", "✅ Entregado"]
                            if st.session_state.edit_estado not in opciones_estado:
                                opciones_estado.insert(0, st.session_state.edit_estado)
                            nuevo_estado = st.selectbox("Estado", options=opciones_estado, key="edit_estado")

                        st.markdown(f"**Total del pedido:** ₡{(nuevo_precio * nueva_cantidad):,.0f}")

                        col_btn1, col_btn2 = st.columns(2)
                        with col_btn1:
                            if st.button("💾 Guardar cambios", key="btn_guardar_cambios_pedido"):
                                c.execute("""
                                    UPDATE productos
                                    SET tienda = ?, categoria = ?, descripcion = ?, cantidad = ?, precio = ?, estado = ?
                                    WHERE id = ?
                                """, (cap1(nueva_tienda), cap1(nueva_categoria), cap1(nueva_descripcion), nueva_cantidad, nuevo_precio, nuevo_estado, id_pedido_actual))
                                conn.commit()
                                st.session_state.gestor_msg = f"¡Pedido #{id_pedido_actual} actualizado exitosamente!"
                                st.session_state.gestor_cargado_id = None
                                st.rerun()
                        with col_btn2:
                            if st.button("🗑 Eliminar Pedido", key="btn_eliminar_pedido"):
                                c.execute("DELETE FROM productos WHERE id=?", (id_pedido_actual,))
                                conn.commit()
                                st.session_state.gestor_msg = f"Pedido #{id_pedido_actual} eliminado."
                                st.session_state.gestor_cargado_id = None
                                st.rerun()

        # 4. Registrar Abonos
        elif menu_principal == "💰 Registrar Abonos":
            st.markdown(
                """
                <div class="top-banner">
                    <div style="font-size: 18px; font-weight: 700;">Gestión de Abonos</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            with card():
                cli_list = pd.read_sql("SELECT id_cliente || ' - ' || nombre AS disp, id_cliente FROM clientes", conn)
                if cli_list.empty:
                    st.warning("No hay clientes registrados.")
                else:
                    sel = st.selectbox("Seleccionar Cliente para Abonar", cli_list["disp"], key="abono_sel")
                    id_c = sel.split(" - ")[0]
                
                    monto = st.number_input("Monto Abonado (₡ CRC)", min_value=0.0, step=1000.0, key="abono_monto_num")
                    st.write("")
                    if st.button("Guardar Abono", key="btn_save_abono"):
                        c.execute("INSERT INTO abonos (id_cliente, monto_crc, fecha) VALUES (?, ?, ?)", (id_c, monto, datetime.now().strftime("%Y-%m-%d %H:%M")))
                        conn.commit()
                        st.success("Abono registrado con éxito.")
                    
                        path_fac, link_wa = generar_factura_imagen(id_c)
                        if path_fac and os.path.exists(path_fac):
                            with open(path_fac, "rb") as file:
                                st.download_button(
                                    label="🖼️ Descargar Factura Actualizada en Imagen",
                                    data=file,
                                    file_name=f"Factura_{id_c}.jpg",
                                    mime="image/jpeg",
                                    key="dl_factura_abono"
                                )
                        if link_wa:
                            st.markdown(f'<a href="{link_wa}" target="_blank" class="btn-whatsapp">📲 Enviar Mensaje por WhatsApp</a>', unsafe_allow_html=True)

            with card():
                st.markdown("<h4 style='color:#be185d;'>📋 Historial de Abonos Recibidos</h4>", unsafe_allow_html=True)
                abonos_detalle = pd.read_sql("SELECT a.fecha as Fecha, c.id_cliente as Código, c.nombre as Cliente, a.monto_crc as 'Monto Abonado (₡)' FROM abonos a JOIN clientes c ON a.id_cliente = c.id_cliente ORDER BY a.id DESC", conn)
                if not abonos_detalle.empty:
                    st.dataframe(abonos_detalle)
                else:
                    st.info("No hay abonos registrados en el sistema.")

        # 5. Gastos Operativos
        elif menu_principal == "💸 Gastos Operativos":
            st.markdown(
                """
                <div class="top-banner">
                    <div style="font-size: 18px; font-weight: 700;">Registrar Gastos u Operaciones</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            with card():
            
                with st.form("form_registro_gastos", clear_on_submit=True):
                    cat_gasto = st.selectbox("Categoría del Gasto", ["🏨 Hospedaje / Hotel", "🍽️ Comida / Alimentación", "🚗 Transporte / Combustible", "✈️ Fletes / Envíos USA-CR", "🛃 Aduana / Impuestos", "📦 Material de Empaque", "💡 Servicios y Operación", "🧩 Otros Gastos"], key="gasto_cat_sel")
                    concepto = st.text_input("Concepto o Descripción", placeholder="Ej. Noche en hotel Miami...", key="gasto_conc_txt")
                    monto_gasto = st.number_input("Monto en Colones (₡ CRC)", min_value=0.0, step=500.0, key="gasto_monto_num")
                    obs_gasto = st.text_area("Notas / Observaciones adicionales", height=60, placeholder="Ej. Factura #1024", key="gasto_obs_txt")

                    st.write("")
                    btn_save_gasto = st.form_submit_button("💾 Registrar Gasto Operativo")

                    if btn_save_gasto:
                        if concepto and monto_gasto > 0:
                            c.execute("INSERT INTO gastos (concepto, categoria, monto_crc, fecha, observaciones) VALUES (?, ?, ?, ?, ?)",
                                      (cap1(concepto), cat_gasto, monto_gasto, datetime.now().strftime("%Y-%m-%d"), cap1(obs_gasto) if obs_gasto else ""))
                            conn.commit()
                            st.success("¡Gasto registrado e integrado correctamente!")
                        else:
                            st.error("Por favor completa la descripción y un monto superior a 0.")

            st.markdown("<h4 style='color:#be185d;'>Últimos Gastos Registrados</h4>", unsafe_allow_html=True)
            gastos_df = pd.read_sql("SELECT fecha as Fecha, concepto as Concepto, categoria as Categoria, monto_crc as 'Monto (CRC)', observaciones as Observaciones FROM gastos ORDER BY id DESC LIMIT 10", conn)
            if not gastos_df.empty:
                st.dataframe(gastos_df)

        # 6. Finanzas y Reportes
        elif menu_principal == "📊 Finanzas y Reportes":
            st.markdown(
                """
                <div class="top-banner">
                    <div style="font-size: 18px; font-weight: 700;">Balance General y Ganancia Real</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            df_pedidos = pd.read_sql("SELECT COALESCE(SUM(precio * cantidad), 0) as total FROM productos WHERE id_cliente IS NOT NULL AND id_cliente != ''", conn)
            df_cajas = pd.read_sql("SELECT COALESCE(SUM(precio_caja), 0) as total FROM cajas_emprendedores", conn)
            df_abonos = pd.read_sql("SELECT COALESCE(SUM(monto_crc), 0) as total FROM abonos", conn)
            df_gastos = pd.read_sql("SELECT COALESCE(SUM(monto_crc), 0) as total FROM gastos", conn)
            df_vr = pd.read_sql("SELECT COALESCE(SUM(total), 0) as total FROM ventas_rapidas", conn)

            total_ventas = (df_pedidos.iloc[0]["total"] or 0.0) + (df_cajas.iloc[0]["total"] or 0.0) + (df_vr.iloc[0]["total"] or 0.0)
            total_ingresos_reales = (df_abonos.iloc[0]["total"] or 0.0) + (df_vr.iloc[0]["total"] or 0.0)
            total_gastos = df_gastos.iloc[0]["total"] or 0.0
            ganancia_real = total_ingresos_reales - total_gastos
            cuentas_por_cobrar = total_ventas - total_ingresos_reales

            col_f1, col_f2, col_f3 = st.columns(3)
            col_f1.metric("Ventas Totales Proyectadas", f"₡{total_ventas:,.0f}")
            col_f2.metric("Cobrado (Ingresos Reales)", f"₡{total_ingresos_reales:,.0f}")
            col_f3.metric("Por Cobrar (Saldos)", f"₡{max(0, cuentas_por_cobrar):,.0f}")

            col_f4, col_f5 = st.columns(2)
            col_f4.metric("Total Gastos (Egresos)", f"₡{total_gastos:,.0f}")
            col_f5.metric("💎 GANANCIA REAL NETA", f"₡{ganancia_real:,.0f}")

            st.write("")
            with card():
                st.markdown("<h4 style='color:#be185d;'>📥 Exportar Reporte Financiero Completo</h4>", unsafe_allow_html=True)

                output = io.BytesIO()
                wb = openpyxl.Workbook()
                wb.remove(wb.active)

                header_fill = PatternFill(start_color="F3B2C9", end_color="F3B2C9", fill_type="solid")
                header_font = Font(name="Segoe UI", size=11, bold=True, color="0f172a")
                data_font = Font(name="Segoe UI", size=10)
                title_font = Font(name="Segoe UI", size=14, bold=True, color="BE185D")
                thin_border = Border(left=Side(style='thin', color='FBCFE8'), right=Side(style='thin', color='FBCFE8'), top=Side(style='thin', color='FBCFE8'), bottom=Side(style='thin', color='FBCFE8'))

                tablas_config = [
                    ("Resumen_Financiero", pd.DataFrame([
                        {"Concepto": "Total Ventas Proyectadas (CRC)", "Monto CRC": total_ventas},
                        {"Concepto": "Ingresos Reales Cobrados", "Monto CRC": total_ingresos_reales},
                        {"Concepto": "Cuentas Por Cobrar", "Monto CRC": max(0, cuentas_por_cobrar)},
                        {"Concepto": "Total Gastos Operativos (Egresos)", "Monto CRC": total_gastos},
                        {"Concepto": "GANANCIA REAL / UTILIDAD NETA", "Monto CRC": ganancia_real}
                    ])),
                    ("Clientes_y_Emprendedores", pd.read_sql("SELECT c.id_cliente as Código, c.nombre as Nombre, c.telefono as Teléfono, c.correo as Correo, COALESCE(e.nombre_caja, 'N/A') as Caja, COALESCE(e.precio_caja, 0.0) as Precio_Caja_CRC FROM clientes c LEFT JOIN cajas_emprendedores e ON c.id_cliente = e.id_cliente", conn)),
                    ("Ventas_Productos", pd.read_sql("SELECT p.id_cliente as Cliente, p.tienda as Tienda, p.categoria as Categoría, p.descripcion as Producto, p.precio as Precio_CRC, p.cantidad as Cantidad, (p.precio * p.cantidad) as Total_CRC, p.estado as Estado FROM productos p WHERE p.id_cliente IS NOT NULL AND p.id_cliente != ''", conn)),
                    ("Ventas_Rapidas_POS", pd.read_sql("SELECT fecha as Fecha, producto as Producto, tipo as Tipo, cantidad as Cantidad, precio_u as Precio_Unitario, descuento as Descuento, total as Total_CRC FROM ventas_rapidas", conn)),
                    ("Ingresos_Abonos", pd.read_sql("SELECT id_cliente as Cliente, monto_crc as Monto_CRC, fecha as Fecha FROM abonos", conn)),
                    ("Gastos_Operativos", pd.read_sql("SELECT fecha as Fecha, categoria as Categoría, concepto as Concepto, monto_crc as Monto_CRC, observaciones as Observaciones FROM gastos", conn))
                ]

                for sheet_name, df_data in tablas_config:
                    ws = wb.create_sheet(title=sheet_name)
                    ws.views.sheetView[0].showGridLines = True
                    ws.cell(row=1, column=1, value=f"Reporte: {sheet_name.replace('_', ' ')}").font = title_font
                    if not df_data.empty:
                        headers = list(df_data.columns)
                        for col_idx, header in enumerate(headers, 1):
                            cell = ws.cell(row=3, column=col_idx, value=header)
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                            cell.border = thin_border
                        for row_idx, row_data in enumerate(df_data.values, 4):
                            for col_idx, value in enumerate(row_data, 1):
                                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                                cell.font = data_font
                                cell.border = thin_border
                                if "CRC" in headers[col_idx-1] or "Monto" in headers[col_idx-1] or "Precio" in headers[col_idx-1] or "Total" in headers[col_idx-1]:
                                    cell.number_format = '₡#,##0'
                                    cell.alignment = Alignment(horizontal="right")
                        for col in ws.columns:
                            max_len = max([len(str(cell.value or '')) for cell in col])
                            col_letter = get_column_letter(col[0].column)
                            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

                wb.save(output)
                st.download_button(
                    label="📊 Descargar Reporte (.xlsx)",
                    data=output.getvalue(),
                    file_name=f"Reporte_Financiero_Minici_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    # =========================================================
    # SECCIÓN 2: MÓDULO POSTVENTA & POS
    # =========================================================
    elif seccion_admin == "⚡ Módulo Postventa & POS":
        menu_postventa = st.radio(
            "Acción Postventa",
            ["⚡ Punto de Venta (POS)", "📦 Inventario y Edición de Stock"],
            horizontal=True,
            label_visibility="collapsed",
            key="menu_post"
        )
        st.write("")

        # 1. Punto de Venta (POS)
        if menu_postventa == "⚡ Punto de Venta (POS)":
            st.markdown(
                """
                <div class="top-banner">
                    <div style="font-size: 18px; font-weight: 700;">Punto de Venta Rápido (POS)</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            with card():
                tipo_origen = st.radio("Tipo de Venta:", ["Del Stock General", "Producto Express (No Registrado)"], horizontal=True, key="pos_origen")

                nombre_prod_vr, precio_u_vr, stock_max_vr, id_prod_inv = "", 0.0, 999, None
                df_prods_disponibles = pd.read_sql("SELECT id, descripcion, precio, cantidad FROM productos WHERE (id_cliente IS NULL OR id_cliente = '') AND cantidad > 0", conn)

                if tipo_origen == "Del Stock General":
                    if not df_prods_disponibles.empty:
                        dict_prods = {f"{r['descripcion']} - ₡{(r['precio'] or 0.0):,.0f} (Stock: {r['cantidad']})": r for _, r in df_prods_disponibles.iterrows()}
                        prod_sel_key = st.selectbox("Seleccionar producto del stock:", list(dict_prods.keys()), key="pos_stock_sel")
                        p_data = dict_prods[prod_sel_key]
                        id_prod_inv, nombre_prod_vr, precio_u_vr, stock_max_vr = int(p_data["id"]), p_data["descripcion"], float(p_data["precio"] or 0.0), int(p_data["cantidad"] or 1)
                        cant_vr = st.number_input("Cantidad:", min_value=1, max_value=max(1, stock_max_vr), value=1, key="pos_cant_num")
                    else:
                        st.warning("⚠️ No hay productos registrados en el Stock General actualmente.")
                        cant_vr = 0
                else:
                    nombre_prod_vr = st.text_input("Descripción del producto:", value="Venta Express", key="pos_expr_desc")
                    precio_u_vr = st.number_input("Precio unitario (₡):", min_value=0.0, value=1000.0, step=500.0, key="pos_expr_precio")
                    cant_vr = st.number_input("Cantidad:", min_value=1, value=1, key="pos_expr_cant")

                c_d1, c_d2 = st.columns(2)
                with c_d1:
                    tipo_desc = st.selectbox("Tipo de Descuento:", ["Sin Descuento", "Porcentaje (%)", "Monto Fijo (₡)"], key="pos_tipo_desc")
                with c_d2:
                    val_desc = st.number_input("Valor del Descuento:", min_value=0.0, value=0.0, key="pos_val_desc")

                subtotal_vr = precio_u_vr * cant_vr
                monto_desc_vr = subtotal_vr * (val_desc / 100.0) if tipo_desc == "Porcentaje (%)" else (val_desc if tipo_desc == "Monto Fijo (₡)" else 0.0)
                total_final_vr = max(0.0, subtotal_vr - monto_desc_vr)

                df_cli_vr = pd.read_sql("SELECT id_cliente || ' — ' || nombre AS display, id_cliente FROM clientes", conn)
            
                opciones_cli_pos = ["Venta General / Anónima"]
                if not df_cli_vr.empty:
                    opciones_cli_pos.extend(list(df_cli_vr["display"]))
            
                cli_vr_selected = st.selectbox("Asignar a cliente (Opcional):", opciones_cli_pos, key="pos_cli_sel")

                st.markdown(f"### **Total Final:** ₡{total_final_vr:,.0f}")

                if st.button("⚡ Procesar y Cobrar Venta", key="btn_pos_cobrar"):
                    if nombre_prod_vr and nombre_prod_vr.strip():
                        prod_vr_clean = cap1(nombre_prod_vr)
                        id_cli_final = df_cli_vr[df_cli_vr["display"] == cli_vr_selected]["id_cliente"].values[0] if cli_vr_selected != "Venta General / Anónima" else None
                        if tipo_origen == "Del Stock General" and id_prod_inv is not None:
                            c.execute("UPDATE productos SET cantidad = cantidad - ? WHERE id = ?", (cant_vr, id_prod_inv))

                        fecha_ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        c.execute("INSERT INTO ventas_rapidas (fecha, id_cliente, producto, tipo, cantidad, precio_u, descuento, total) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                                  (fecha_ahora, id_cli_final, prod_vr_clean, tipo_origen, cant_vr, precio_u_vr, monto_desc_vr, total_final_vr))
                        conn.commit()
                        st.success("✅ Venta procesada correctamente.")
                        if id_cli_final:
                            path_fac, link_wa = generar_factura_imagen(id_cli_final)
                            if path_fac and os.path.exists(path_fac):
                                with open(path_fac, "rb") as file:
                                    st.download_button(
                                        label="🖼️ Descargar Factura de Venta en Imagen",
                                        data=file,
                                        file_name=f"Factura_{id_cli_final}.jpg",
                                        mime="image/jpeg",
                                        key="dl_factura_pos"
                                    )
                            if link_wa:
                                st.markdown(f'<a href="{link_wa}" target="_blank" class="btn-whatsapp">📲 Enviar Comprobante por WhatsApp</a>', unsafe_allow_html=True)
                    else:
                        st.error("Ingresa o selecciona un producto válido para realizar la venta.")

        # 2. Inventario de Stock en Galería Y Edición
        elif menu_postventa == "📦 Inventario y Edición de Stock":
            st.markdown(
                """
                <div class="top-banner">
                    <div style="font-size: 18px; font-weight: 700;">Inventario, Edición y Stock General</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            
            with card():
                st.markdown("<h4 style='color:#be185d;'>➕ Añadir Nuevo Producto al Inventario</h4>", unsafe_allow_html=True)

                # El escáner va fuera del formulario para poder llenar el campo del código al instante
                escanear_codigo_barras("new_stock", label="📷 Escanear código de barras con la cámara (opcional)")
                if "codigo_new_stock" not in st.session_state:
                    st.session_state["codigo_new_stock"] = ""

                with st.form("form_agregar_stock", clear_on_submit=True):
                    col_i1, col_i2 = st.columns(2)
                    with col_i1:
                        new_desc = st.text_input("Descripción del Producto", placeholder="Ej. Blusa elegante")
                        new_precio = st.number_input("Precio (₡ CRC)", min_value=0.0, value=5000.0, step=500.0)
                        new_cant = st.number_input("Stock Inicial", min_value=1, value=1, step=1)
                    with col_i2:
                        new_tienda = st.text_input("Tienda / Proveedor", placeholder="Ej. Zara o Local")
                        new_cat = st.text_input("Categoría", placeholder="Ej. Ropa")
                        new_barcode = st.text_input("Código de Barras (Opcional)", placeholder="Escanea con la cámara o escribe el código", key="codigo_new_stock")

                    new_obs = st.text_area("Observaciones", placeholder="Detalles adicionales...", height=60)

                    btn_add_stock = st.form_submit_button("💾 Guardar Producto en Inventario")

                    if btn_add_stock:
                        if new_desc.strip():
                            c.execute("""INSERT INTO productos (id_cliente, tienda, categoria, descripcion, precio, moneda, cantidad, estado, observaciones, codigo_barras)
                                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                      (None, cap1(new_tienda), cap1(new_cat), cap1(new_desc), new_precio, "CRC", new_cant, "Disponible en Inventario", cap1(new_obs) if new_obs else "", new_barcode.strip() if new_barcode else ""))
                            conn.commit()
                            flash("success", f"¡Producto '{new_desc}' añadido al inventario exitosamente!")
                            st.rerun()
                        else:
                            st.error("Debes ingresar al menos la descripción del producto.")
            
            with card():
                st.markdown("<h4 style='color:#be185d;'>📊 Tabla de Inventario (Stock General)</h4>", unsafe_allow_html=True)
                df_stock_table = pd.read_sql("SELECT id, codigo_barras as 'Cód. Barras', descripcion as Producto, tienda as Tienda, categoria as Categoría, precio as Precio, cantidad as Stock FROM productos WHERE id_cliente IS NULL OR id_cliente = ''", conn)
                st.dataframe(df_stock_table)

            with card():
                st.markdown("<h4 style='color:#be185d;'>✏️ Editar Producto del Stock</h4>", unsafe_allow_html=True)
                if not df_stock_table.empty:
                    prod_edit_dict = {f"ID: {r['id']} - {r['Producto']}": r['id'] for _, r in df_stock_table.iterrows()}
                    prod_edit_sel = st.selectbox("Seleccionar producto a editar:", list(prod_edit_dict.keys()), key="inv_edit_sel")
                    id_prod_edit = prod_edit_dict[prod_edit_sel]

                    c.execute("SELECT descripcion, tienda, categoria, precio, cantidad, codigo_barras FROM productos WHERE id = ?", (id_prod_edit,))
                    p_edit_data = c.fetchone()
                    if p_edit_data:
                        # Si aún no se ha escaneado nada para este producto, se parte del valor guardado en la BD
                        clave_barcode_edit = f"inv_ed_bar_{id_prod_edit}"
                        escanear_codigo_barras(f"edit_stock_{id_prod_edit}", label="📷 Escanear nuevo código de barras con la cámara (opcional)")
                        if clave_barcode_edit not in st.session_state:
                            st.session_state[clave_barcode_edit] = p_edit_data[5] or ""
                        if st.session_state.get(f"codigo_edit_stock_{id_prod_edit}"):
                            st.session_state[clave_barcode_edit] = st.session_state.pop(f"codigo_edit_stock_{id_prod_edit}")

                        c_e1, c_e2 = st.columns(2)
                        with c_e1:
                            n_desc = st.text_input("Descripción", value=p_edit_data[0] or "", key=f"inv_ed_desc_{id_prod_edit}")
                            n_precio = st.number_input("Precio", value=float(p_edit_data[3]) if p_edit_data[3] is not None else 0.0, step=500.0, key=f"inv_ed_pre_{id_prod_edit}")
                            n_cant = st.number_input("Stock", value=int(p_edit_data[4]) if p_edit_data[4] is not None else 1, step=1, key=f"inv_ed_can_{id_prod_edit}")
                        with c_e2:
                            n_tienda = st.text_input("Tienda", value=p_edit_data[1] if p_edit_data[1] else "", key=f"inv_ed_tie_{id_prod_edit}")
                            n_cat = st.text_input("Categoría", value=p_edit_data[2] if p_edit_data[2] else "", key=f"inv_ed_cat_{id_prod_edit}")
                            n_barcode = st.text_input("Código de Barras", key=clave_barcode_edit)

                        c_b1, c_b2 = st.columns(2)
                        with c_b1:
                            if st.button("💾 Guardar Cambios de Stock", key="btn_inv_save"):
                                c.execute("UPDATE productos SET descripcion=?, tienda=?, categoria=?, precio=?, cantidad=?, codigo_barras=? WHERE id=?", 
                                          (cap1(n_desc), cap1(n_tienda), cap1(n_cat), n_precio, n_cant, n_barcode.strip() if n_barcode else "", id_prod_edit))
                                conn.commit()
                                flash("success", "¡Producto en stock actualizado!")
                                st.rerun()
                        with c_b2:
                            if st.button("🗑 Eliminar Producto", key="btn_inv_del"):
                                c.execute("DELETE FROM productos WHERE id=?", (id_prod_edit,))
                                conn.commit()
                                flash("warning", "Producto eliminado del inventario.")
                                st.rerun()
                else:
                    st.info("No hay productos en el stock general para editar.")
            
            with card():
                st.markdown("<h4 style='color:#be185d;'>📸 Galería de Productos Locales</h4>", unsafe_allow_html=True)
                df_stock = pd.read_sql("SELECT id, descripcion, precio, cantidad, categoria, tienda, foto_path, codigo_barras FROM productos WHERE id_cliente IS NULL OR id_cliente = ''", conn)
            
                if df_stock.empty:
                    st.info("No hay productos generales en stock actualmente.")
                else:
                    for _, r in df_stock.iterrows():
                        with card("sub"):
                            col_img, col_info = st.columns([1, 2])
                            with col_img:
                                if r["foto_path"] and os.path.exists(os.path.join("fotos_productos", r["foto_path"])):
                                    st.image(os.path.join("fotos_productos", r["foto_path"]))
                                else:
                                    st.caption("📷 Sin foto disponible")
                            with col_info:
                                st.markdown(f"✨ **{r['descripcion']}**")
                                st.markdown(f"📦 Stock Disponible: **{int(r['cantidad'] if pd.notna(r['cantidad']) else 1)}**")
                        
                                barcode_txt = r['codigo_barras'] if pd.notna(r['codigo_barras']) and r['codigo_barras'] != '' else 'N/A'
                                st.caption(f"🏷️ Categoría: {r['categoria']} | Tienda: {r['tienda']} | 📌 Cód. Barras: {barcode_txt}")
                        
                                st.markdown(f"💰 Precio: **₡{(r['precio'] or 0.0):,.0f}**")

# -------------------------------------------------------------
# 6. PANEL DE CLIENTE / EMPRENDEDOR
# -------------------------------------------------------------
elif st.session_state.user_role == "client":
    id_cli = st.session_state.current_client

    c.execute("SELECT nombre FROM clientes WHERE id_cliente = ?", (id_cli,))
    res_cli = c.fetchone()
    client_name = res_cli[0] if res_cli else "Cliente"

    c.execute("SELECT nombre_caja, precio_caja FROM cajas_emprendedores WHERE id_cliente = ?", (id_cli,))
    res_caja_client = c.fetchone()
    caja_nombre_c, caja_precio_c = (res_caja_client[0], res_caja_client[1] or 0.0) if res_caja_client else (None, 0.0)

    prods = pd.read_sql("SELECT descripcion, tienda, precio, cantidad, estado, foto_path FROM productos WHERE id_cliente = ?", conn, params=(id_cli,))
    abonos_df = pd.read_sql("SELECT COALESCE(SUM(monto_crc), 0.0) as total FROM abonos WHERE id_cliente = ?", conn, params=(id_cli,))
    
    total_abonos = abonos_df.iloc[0]["total"] or 0.0
    total_articulos = (prods['precio'].fillna(0.0) * prods['cantidad'].fillna(1)).sum() if not prods.empty else 0.0
    total_compras = total_articulos + caja_precio_c
    saldo_pendiente = total_compras - total_abonos

    col_cli1, col_cli2 = st.columns([4, 1])
    with col_cli1:
        st.markdown(
            f'<div class="app-header"><span class="app-brand">Hola, {html.escape(str(client_name))}</span><span class="app-role">Tu cuenta</span></div>',
            unsafe_allow_html=True,
        )
    with col_cli2:
        if st.button("🚪 Salir", key="btn_salir_client"):
            st.session_state.user_role = None
            st.session_state.current_client = None
            st.rerun()

    with card():
        st.markdown("<h4 style='color:#be185d;'>💰 Estado de Cuenta</h4>", unsafe_allow_html=True)
        if caja_nombre_c:
            st.caption(f"💼 Caja Asignada: **{caja_nombre_c}** | Cuota: **₡{caja_precio_c:,.0f}**")
        
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Total Cargos", f"₡{total_compras:,.0f}")
        col_m2.metric("Total Abonado", f"₡{total_abonos:,.0f}")
        col_m3.metric("Saldo Pendiente", f"₡{max(0.0, saldo_pendiente):,.0f}")

    notifs = pd.read_sql("SELECT titulo, mensaje, fecha FROM notificaciones WHERE id_cliente = ? ORDER BY id DESC LIMIT 5", conn, params=(id_cli,))
    if not notifs.empty:
        with card():
            st.markdown("<h4 style='color:#be185d;'>🔔 Notificaciones Recientes</h4>", unsafe_allow_html=True)
            for _, row in notifs.iterrows():
                st.markdown(f"**{row['titulo']}** ({row['fecha']})")
                st.caption(row["mensaje"])
                st.divider()

    with card():
        st.markdown("<h4 style='color:#be185d;'>📦 Tus Pedidos / Artículos</h4>", unsafe_allow_html=True)
        if prods.empty:
            st.info("Aún no tienes artículos registrados.")
        else:
            for _, row in prods.iterrows():
                col_img, col_info = st.columns([1, 2])
                with col_img:
                    if row["foto_path"] and os.path.exists(os.path.join("fotos_productos", row["foto_path"])):
                        st.image(os.path.join("fotos_productos", row["foto_path"]))
                    else:
                        st.text("📷 Sin foto")
                with col_info:
                    st.markdown(f"**{row['descripcion']}**")
                    st.caption(f"Tienda: {row['tienda']} | Estado: {row['estado']}")
                
                    cant_prod = int(row['cantidad'] if pd.notna(row['cantidad']) else 1)
                    precio_prod = row['precio'] if pd.notna(row['precio']) else 0.0
                    tot_prod = precio_prod * cant_prod
                    st.markdown(f"📦 Cantidad: {cant_prod} | 💰 **Total: ₡{tot_prod:,.0f}**")
                st.divider()