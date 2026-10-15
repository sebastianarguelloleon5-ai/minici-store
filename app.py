import os
os.environ["OPENBLAS_NUM_THREADS"] = "1"

import sqlite3
from datetime import datetime, timedelta
import io
import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageChops, ImageOps
import streamlit as st
import urllib.parse
import urllib.request
import urllib.error
import html
import json
import zipfile
import tempfile
import base64
import hmac
import secrets
import threading
import time
import shutil

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
# 1A. CARPETA DE TRABAJO Y MODO SERVIDOR
# -------------------------------------------------------------
# Todo (base de datos, fotos, comprobantes, logo) se busca junto a este archivo, sin importar desde dónde se arranque la app.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)
# En un servidor VPS el disco es permanente. Con MINICI_DISCO_PERSISTENTE=1 la base local manda siempre:
# nunca se reemplaza por un respaldo al arrancar (solo se recupera si el archivo no existe).
DISCO_PERSISTENTE = os.environ.get("MINICI_DISCO_PERSISTENTE", "").strip().lower() in ("1", "true", "si", "sí", "yes")


# -------------------------------------------------------------
# 1B. CONFIGURACIÓN SEGURA Y RESPALDOS EN GITHUB
# -------------------------------------------------------------
DB_FILE = "minici_store.db"
GH_RUTA_ZIP = "respaldos/respaldo_completo.zip"
GH_RUTA_DB = "respaldos/minici_store.db"
_ESTADO_FILE = "respaldo_estado.json"


def _secret(nombre, default=None):
    """Lee un valor de los 'Secrets' de Streamlit sin romper si no existen."""
    try:
        valor = st.secrets[nombre]
        return valor if valor not in (None, "") else default
    except Exception:
        return default


_GH_CACHE = {}


def _gh_config():
    token, repo = _secret("GITHUB_TOKEN"), _secret("GITHUB_REPO")
    if not token or not repo:
        return _GH_CACHE.get("cfg")  # un hilo en segundo plano puede no ver los Secrets: usa lo ya leído
    repo = str(repo).strip().strip("/")
    for prefijo in ("https://github.com/", "http://github.com/", "github.com/"):
        if repo.lower().startswith(prefijo):
            repo = repo[len(prefijo):]
    if repo.lower().endswith(".git"):
        repo = repo[:-4]
    cfg = {"token": str(token).strip().strip('"').strip("'"),
           "repo": repo.strip("/"),
           "branch": str(_secret("GITHUB_BRANCH", "main")).strip()}
    _GH_CACHE["cfg"] = cfg
    return cfg


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


def _gh_texto_http(e):
    """Mensaje que GitHub devuelve junto al error (por ejemplo 'Branch main not found')."""
    try:
        return str(json.loads(e.read().decode("utf-8", "ignore")).get("message", ""))
    except Exception:
        return ""


def _gh_mensaje_error(e):
    if isinstance(e, urllib.error.HTTPError):
        detalle = _gh_texto_http(e)
        extra = f" GitHub dice: «{detalle}»." if detalle else ""
        if e.code == 401:
            return "GitHub rechazó el token (401). Revisa que GITHUB_TOKEN sea correcto y no haya vencido." + extra
        if e.code == 403:
            return "GitHub negó el permiso (403). El token necesita 'Contents: Read and write' sobre el repositorio." + extra
        if e.code == 404:
            return "GitHub no encontró el repositorio (404). Revisa GITHUB_REPO (usuario/repositorio) y que el token tenga acceso a él." + extra
        if e.code == 409 and "empty" in detalle.lower():
            return "El repositorio está vacío. Créale un archivo README.md desde GitHub y vuelve a intentar." + extra
        if e.code == 422:
            return "GitHub rechazó los datos del respaldo (422)." + extra
        return f"GitHub respondió con el error {e.code}." + extra
    return f"No se pudo conectar con GitHub: {e}"


def _gh_rama(cfg):
    """Devuelve la rama a usar. Si la configurada no existe, usa la rama principal del repositorio.
    Devuelve None si el repositorio aún no tiene ramas (GitHub crea la principal al guardar el primer archivo)."""
    base = f"https://api.github.com/repos/{cfg['repo']}"
    try:
        _gh_request("GET", f"{base}/branches/{urllib.parse.quote(cfg['branch'])}", cfg["token"])
        return cfg["branch"]
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise
    info = json.loads(_gh_request("GET", base, cfg["token"]))
    principal = info.get("default_branch")
    if principal:
        try:
            _gh_request("GET", f"{base}/branches/{urllib.parse.quote(principal)}", cfg["token"])
            return principal
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise
    return None


def github_subir(ruta_repo, contenido, mensaje):
    """Crea o actualiza un archivo en el repositorio de respaldos. Devuelve (ok, texto)."""
    cfg = _gh_config()
    if not cfg:
        return False, "GitHub no está configurado."
    url = f"https://api.github.com/repos/{cfg['repo']}/contents/{urllib.parse.quote(ruta_repo)}"
    try:
        rama = _gh_rama(cfg)

        def sha_actual():
            consulta = f"{url}?ref={urllib.parse.quote(rama)}" if rama else url
            try:
                info = json.loads(_gh_request("GET", consulta, cfg["token"]))
                return info.get("sha") if isinstance(info, dict) else None
            except urllib.error.HTTPError as e:
                if e.code != 404:
                    raise
                return None

        payload = {"message": mensaje, "content": base64.b64encode(contenido).decode("ascii")}
        if rama:
            payload["branch"] = rama
        sha = sha_actual()
        if sha:
            payload["sha"] = sha
        try:
            _gh_request("PUT", url, cfg["token"], payload)
        except urllib.error.HTTPError as e:
            # el archivo cambió o ya existía: se vuelve a leer su versión y se reintenta una vez
            if e.code in (409, 422):
                nuevo_sha = sha_actual()
                if nuevo_sha and nuevo_sha != payload.get("sha"):
                    payload["sha"] = nuevo_sha
                    _gh_request("PUT", url, cfg["token"], payload)
                    return True, "ok"
            raise
        return True, "ok"
    except Exception as e:
        return False, _gh_mensaje_error(e)


def github_leer(ruta_repo):
    cfg = _gh_config()
    rama = _gh_rama(cfg)
    url = f"https://api.github.com/repos/{cfg['repo']}/contents/{urllib.parse.quote(ruta_repo)}"
    if rama:
        url += f"?ref={urllib.parse.quote(rama)}"
    return _gh_request("GET", url, cfg["token"], raw=True)


def _extraer_fotos_zip(z):
    """Extrae solo las fotos de productos de un respaldo (sin rutas peligrosas)."""
    for nombre in z.namelist():
        if nombre.startswith("fotos_productos/") and not nombre.endswith("/") and ".." not in nombre:
            destino = os.path.join("fotos_productos", os.path.basename(nombre))
            os.makedirs("fotos_productos", exist_ok=True)
            with open(destino, "wb") as f:
                f.write(z.read(nombre))


def _filas_en_bd(ruta):
    """Cantidad de registros importantes (clientes + productos + abonos) de un archivo de base de datos."""
    total = 0
    try:
        cx = sqlite3.connect(ruta)
        try:
            for tabla in ("clientes", "productos", "abonos"):
                try:
                    total += cx.execute(f"SELECT COUNT(*) FROM {tabla}").fetchone()[0]
                except Exception:
                    pass
        finally:
            cx.close()
    except Exception:
        pass
    return total


def _restaurar_al_iniciar():
    """Recupera los datos de GitHub cuando el servidor arrancó de cero.
    - Si no hay base de datos: usa la del respaldo.
    - Si hay una base pero este servidor nunca ha respaldado (típico cuando la base viene dentro del
      repositorio y está desactualizada): usa la del respaldo si tiene más datos, y guarda una copia de la otra."""
    if not _gh_config():
        return False
    hay_local = os.path.exists(DB_FILE)
    if hay_local and (DISCO_PERSISTENTE or os.path.exists(_ESTADO_FILE)):
        return False  # disco permanente o servidor que ya respalda: la base local es la vigente
    bd, zdatos = None, None
    try:
        bd = github_leer(GH_RUTA_DB)
    except Exception:
        pass
    try:
        zdatos = github_leer(GH_RUTA_ZIP)
    except Exception:
        pass
    if bd is None and zdatos:
        try:
            with zipfile.ZipFile(io.BytesIO(zdatos)) as z:
                if DB_FILE in z.namelist():
                    bd = z.read(DB_FILE)
        except Exception:
            bd = None
    if bd is None:
        return False
    try:
        if hay_local:
            tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
            tmp.write(bd)
            tmp.close()
            try:
                filas_respaldo = _filas_en_bd(tmp.name)
            finally:
                os.unlink(tmp.name)
            if _filas_en_bd(DB_FILE) >= filas_respaldo:
                return False  # la base local tiene igual o más datos: no se toca
            shutil.copy2(DB_FILE, "minici_store.antes_de_restaurar.db")
        with open(DB_FILE, "wb") as f:
            f.write(bd)
        if zdatos:
            with zipfile.ZipFile(io.BytesIO(zdatos)) as z:
                _extraer_fotos_zip(z)
        return True
    except Exception:
        return False


@st.cache_resource
def _restaurar_una_vez():
    """Se ejecuta una sola vez por arranque del servidor (no en cada recarga de pantalla)."""
    return _restaurar_al_iniciar()


_gh_config()  # deja lista la configuración de GitHub para el respaldo automático
if _restaurar_una_vez() and not st.session_state.get("_rest_avisado"):
    st.session_state["_restaurado_gh"] = True
    st.session_state["_rest_avisado"] = True

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

def asegurar_columnas():
    """Agrega columnas nuevas a bases de datos viejas (también se usa tras restaurar un respaldo)."""
    nuevas = {
        "productos": [
            ("categoria", "TEXT"), ("precio", "REAL DEFAULT 0.0"), ("moneda", "TEXT DEFAULT 'CRC'"),
            ("cantidad", "INTEGER DEFAULT 1"), ("observaciones", "TEXT"), ("foto_path", "TEXT"),
            ("codigo_barras", "TEXT"),
        ],
        "clientes": [
            ("provincia", "TEXT"), ("canton", "TEXT"), ("distrito", "TEXT"),
            ("senas", "TEXT"), ("metodo_envio", "TEXT"),
            ("origen", "TEXT"), ("fecha_registro", "TEXT"),
        ],
    }
    nuevas["ventas_rapidas"] = [("id_venta", "INTEGER"), ("metodo_pago", "TEXT"), ("id_producto", "INTEGER")]
    nuevas["abonos"] = [("metodo_pago", "TEXT"), ("referencia", "TEXT")]
    # Tablas básicas (por si se restauró un respaldo muy viejo que no las traía)
    conn.execute("CREATE TABLE IF NOT EXISTS clientes (id_cliente TEXT PRIMARY KEY, nombre TEXT NOT NULL, telefono TEXT, correo TEXT)")
    conn.execute("CREATE TABLE IF NOT EXISTS cajas_emprendedores (id INTEGER PRIMARY KEY AUTOINCREMENT, id_cliente TEXT UNIQUE, nombre_caja TEXT, precio_caja REAL DEFAULT 0.0)")
    conn.execute("""CREATE TABLE IF NOT EXISTS productos (id INTEGER PRIMARY KEY AUTOINCREMENT, id_cliente TEXT, tienda TEXT, categoria TEXT,
                   descripcion TEXT, precio REAL DEFAULT 0.0, moneda TEXT DEFAULT 'CRC', cantidad INTEGER DEFAULT 1, estado TEXT,
                   observaciones TEXT, foto_path TEXT, codigo_barras TEXT)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS ventas_rapidas (id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, id_cliente TEXT, producto TEXT,
                   tipo TEXT, cantidad INTEGER, precio_u REAL, descuento REAL, total REAL)""")
    conn.execute("CREATE TABLE IF NOT EXISTS abonos (id INTEGER PRIMARY KEY AUTOINCREMENT, id_cliente TEXT, monto_crc REAL, fecha TEXT)")
    conn.execute("""CREATE TABLE IF NOT EXISTS gastos (id INTEGER PRIMARY KEY AUTOINCREMENT, concepto TEXT NOT NULL, categoria TEXT,
                   monto_crc REAL NOT NULL, fecha TEXT NOT NULL, observaciones TEXT)""")
    conn.execute("CREATE TABLE IF NOT EXISTS notificaciones (id INTEGER PRIMARY KEY AUTOINCREMENT, id_cliente TEXT, titulo TEXT, mensaje TEXT, leida INTEGER DEFAULT 0, fecha TEXT)")
    conn.execute("CREATE TABLE IF NOT EXISTS config (clave TEXT PRIMARY KEY, valor TEXT)")
    conn.execute("""CREATE TABLE IF NOT EXISTS ventas (
                   id INTEGER PRIMARY KEY AUTOINCREMENT, folio TEXT, fecha TEXT, id_cliente TEXT, cliente_nombre TEXT,
                   subtotal REAL, descuento REAL, total REAL, metodo_pago TEXT, recibido REAL, vuelto REAL,
                   referencia TEXT, estado TEXT DEFAULT 'Pagada')""")
    conn.execute("""CREATE TABLE IF NOT EXISTS apartados (
                   id INTEGER PRIMARY KEY AUTOINCREMENT, folio TEXT, fecha TEXT, fecha_limite TEXT, id_cliente TEXT,
                   total REAL, prima_pct REAL, cancelacion_pct REAL, estado TEXT DEFAULT 'Activo', notas TEXT, fecha_cierre TEXT)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS apartado_items (
                   id INTEGER PRIMARY KEY AUTOINCREMENT, id_apartado INTEGER, id_producto INTEGER,
                   descripcion TEXT, cantidad INTEGER, precio_u REAL)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS apartado_pagos (
                   id INTEGER PRIMARY KEY AUTOINCREMENT, id_apartado INTEGER, fecha TEXT, monto REAL,
                   metodo_pago TEXT, referencia TEXT, tipo TEXT)""")
    for tabla, cols in nuevas.items():
        existentes = {r[1] for r in conn.execute(f"PRAGMA table_info({tabla})")}
        if not existentes:
            continue  # la tabla todavía no existe (se crea más abajo y luego se vuelve a llamar)
        for nombre_col, tipo in cols:
            if nombre_col not in existentes:
                conn.execute(f"ALTER TABLE {tabla} ADD COLUMN {nombre_col} {tipo}")
    conn.commit()


asegurar_columnas()

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

asegurar_columnas()  # ya existen todas las tablas: agrega las columnas nuevas (abonos, ventas...)

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


def titulo_es(texto):
    """'san rafael de alajuela' -> 'San Rafael de Alajuela' (respeta de, del, la, los, las, y)."""
    chicas = {"de", "del", "la", "las", "los", "y", "el"}
    palabras = str(texto or "").strip().split()
    return " ".join(w.lower() if (i > 0 and w.lower() in chicas) else w[:1].upper() + w[1:].lower()
                    for i, w in enumerate(palabras))


def flash(tipo, mensaje):
    """Guarda un mensaje para mostrarlo después de st.rerun() (si no, se pierde)."""
    st.session_state["_flash"] = (tipo, mensaje)


def mostrar_flash():
    f = st.session_state.pop("_flash", None)
    if f:
        getattr(st, f[0], st.info)(f[1])


def _copiar_bd(cx):
    """Copia consistente de la base de datos (bytes), aunque otra persona esté guardando."""
    cx.commit()
    tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
    tmp.close()
    try:
        destino = sqlite3.connect(tmp.name)
        cx.backup(destino)
        destino.close()
        with open(tmp.name, "rb") as f:
            return f.read()
    finally:
        os.unlink(tmp.name)


def crear_respaldo_zip(con=None):
    """Zip con la base de datos, todas las fotos y un Excel legible. Devuelve (zip, db)."""
    cx = conn if con is None else con
    db_bytes = _copiar_bd(cx)

    xlsx_bytes = None
    try:
        buf = io.BytesIO()
        hojas = 0
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            for tabla, hoja in [("clientes", "Clientes"), ("cajas_emprendedores", "Cajas"),
                                ("productos", "Productos"), ("abonos", "Abonos"),
                                ("gastos", "Gastos"), ("ventas", "Ventas POS"), ("ventas_rapidas", "Detalle ventas"),
                                ("apartados", "Apartados"), ("apartado_pagos", "Pagos apartados"),
                                ("notificaciones", "Notificaciones")]:
                try:
                    pd.read_sql(f"SELECT * FROM {tabla}", cx).to_excel(w, sheet_name=hoja, index=False)
                    hojas += 1
                except Exception:
                    pass
            if hojas == 0:
                pd.DataFrame({"info": ["sin datos"]}).to_excel(w, sheet_name="Datos", index=False)
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


# ---------- estado del respaldo (archivo aparte, para no alterar la base de datos) ----------
INTERVALO_RESPALDO_SEG = 10 * 60          # como máximo cada 10 minutos (solo si hubo cambios)
REINTENTO_RESPALDO_SEG = 5 * 60           # si falla, reintenta a los 5 minutos
RESPALDO_ZIP_CADA_SEG = 24 * 60 * 60      # el zip con fotos se sube como máximo 1 vez al día (o si hay fotos nuevas)


def _leer_estado():
    try:
        with open(_ESTADO_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _guardar_estado(**cambios):
    try:
        est = _leer_estado()
        est.update(cambios)
        tmp = _ESTADO_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(est, f)
        os.replace(tmp, _ESTADO_FILE)
    except Exception:
        pass


def _mtime_fotos():
    ms = [0.0]
    if os.path.isdir("fotos_productos"):
        for nombre in os.listdir("fotos_productos"):
            try:
                ms.append(os.path.getmtime(os.path.join("fotos_productos", nombre)))
            except OSError:
                pass
    return max(ms)


def _ultimo_cambio():
    base = os.path.getmtime(DB_FILE) if os.path.exists(DB_FILE) else 0.0
    return max(base, _mtime_fotos())


def hacer_respaldo(origen="manual", con=None, solo_bd=False):
    cx = conn if con is None else con
    inicio = time.time()
    ahora = datetime.now()
    if solo_bd:
        db_bytes, zip_bytes = _copiar_bd(cx), None
    else:
        zip_bytes, db_bytes = crear_respaldo_zip(cx)
    res = {"zip": zip_bytes, "nombre": f"respaldo_minici_{ahora:%Y-%m-%d_%H%M}.zip",
           "hora": ahora.strftime("%d/%m/%Y %H:%M"), "github": None, "detalle": "", "origen": origen}
    if _gh_config():
        msg = f"Respaldo {origen} Minici Store {ahora:%Y-%m-%d %H:%M}"
        ok_db, m_db = github_subir(GH_RUTA_DB, db_bytes, msg)
        ok_zip, m_zip = True, ""
        if zip_bytes is not None:
            if len(zip_bytes) < 90 * 1024 * 1024:
                ok_zip, m_zip = github_subir(GH_RUTA_ZIP, zip_bytes, msg)
            else:
                ok_zip, m_zip = False, "El respaldo con fotos pesa más de 90 MB; solo se subió la base de datos."
        res["github"] = bool(ok_db and ok_zip)
        res["detalle"] = "" if res["github"] else (m_db if not ok_db else m_zip)
        if res["github"]:
            cambios = {"epoch_ok": inicio, "txt_ok": res["hora"], "origen": origen, "error": None,
                       "epoch_intento": time.time()}
            if zip_bytes is not None:
                cambios["epoch_zip"] = inicio
            _guardar_estado(**cambios)
        else:
            _guardar_estado(error=res["detalle"] or "Error desconocido", epoch_intento=time.time())
    return res


def _respaldo_auto_si_toca():
    """Hace el respaldo automático si ya pasaron 2 horas, hubo cambios y la base no está vacía."""
    if not _gh_config() or not os.path.exists(DB_FILE):
        return False
    est = _leer_estado()
    ahora = time.time()
    ok = est.get("epoch_ok")
    if ok and ahora - ok < INTERVALO_RESPALDO_SEG:
        return False
    if est.get("error") and ahora - est.get("epoch_intento", 0) < REINTENTO_RESPALDO_SEG:
        return False
    if ok and _ultimo_cambio() <= ok:
        return False  # nada cambió desde el último respaldo
    con = sqlite3.connect(DB_FILE, timeout=30)
    try:
        total = con.execute("SELECT (SELECT COUNT(*) FROM clientes) + (SELECT COUNT(*) FROM productos)").fetchone()[0]
        if not total:
            return False  # nunca reemplazar un buen respaldo con una base vacía
        ez = est.get("epoch_zip")
        con_zip = (not ez) or (ahora - ez >= RESPALDO_ZIP_CADA_SEG) or (_mtime_fotos() > ez)
        hacer_respaldo(origen="automático", con=con, solo_bd=not con_zip)
        return True
    except Exception as e:
        _guardar_estado(error=f"{e}", epoch_intento=time.time())
        return False
    finally:
        con.close()


def _bucle_respaldo_auto():
    time.sleep(60)  # deja que la app termine de arrancar
    while True:
        try:
            _respaldo_auto_si_toca()
        except Exception:
            pass
        time.sleep(300)


@st.cache_resource
def _iniciar_respaldo_automatico():
    if any(t.name == "respaldo-auto" and t.is_alive() for t in threading.enumerate()):
        return None
    _gh_config()
    hilo = threading.Thread(target=_bucle_respaldo_auto, name="respaldo-auto", daemon=True)
    hilo.start()
    return hilo


def estado_respaldo_auto():
    """(tipo, texto) para mostrar arriba en el panel."""
    if not _gh_config():
        if DISCO_PERSISTENTE:
            return "off", "🕒 Respaldo en GitHub: apagado (opcional). Tus datos están guardados en el disco del servidor."
        return "off", "🕒 Respaldo automático: apagado (falta conectar GitHub). Los datos se pierden si la app se reinicia."
    est = _leer_estado()
    ok, intento = est.get("epoch_ok"), est.get("epoch_intento", 0)
    if est.get("error") and (not ok or intento > ok):
        return "error", f"⚠️ El último intento de respaldo falló y se reintenta solo: {str(est['error'])[:150]}"
    if ok:
        return "ok", f"🕒 Respaldo automático activo (cada 10 min si hay cambios) · último guardado: {est.get('txt_ok', '')} ({est.get('origen', 'manual')})"
    return "espera", "🕒 Respaldo automático activo (cada 10 min si hay cambios). El primero se hará en unos minutos."


_iniciar_respaldo_automatico()


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
        asegurar_columnas()
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
        import importlib
        mpl = importlib.import_module("matplotlib")
        mp = os.path.join(mpl.get_data_path(), "fonts", "ttf")
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


# ---------- 1) MOTOR GENERAL DE FACTURACIÓN: imagen ----------
def _fecha_corta(f):
    try:
        return datetime.strptime(str(f)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return str(f)[:10]


def _color_tono(tono):
    K = FACT_COLORS
    return {"normal": K["text"], "green": K["green"], "muted": K["muted"], "red": K["raspberry"],
            "raspberry": K["raspberry"], "orange": K["orange"]}.get(tono, K["text"])


def _link_wa(telefono, mensaje):
    tel = "".join(filter(str.isdigit, str(telefono))) if telefono else ""
    if len(tel) == 8:
        tel = "506" + tel
    return f"https://wa.me/{tel}?text={urllib.parse.quote(mensaje)}" if tel else None


def render_factura_imagen(doc):
    """Dibuja cualquier comprobante (estado de cuenta, venta, apartado) con el mismo diseño. Devuelve la ruta del .jpg."""
    K = FACT_COLORS
    S, W = 2, 1000
    logo, bg_logo = _logo_factura()
    header_bg = bg_logo if bg_logo else "#FEEFF6"
    meas = ImageDraw.Draw(Image.new("RGB", (4, 4)))

    def F(bold, size):
        return _fuente(bold, size * S)

    def tw(txt, font):
        return meas.textlength(txt, font=font) / S

    def ajustar(txt, font, maxw):
        if tw(txt, font) <= maxw:
            return txt
        while txt and tw(txt + "…", font) > maxw:
            txt = txt[:-1]
        return txt.rstrip() + "…"

    def envolver(txt, font, maxw):
        lineas, actual = [], ""
        for palabra in str(txt).split():
            prueba = (actual + " " + palabra).strip()
            if tw(prueba, font) <= maxw or not actual:
                actual = prueba
            else:
                lineas.append(actual)
                actual = palabra
        if actual:
            lineas.append(actual)
        return lineas

    filas = doc.get("filas", [])
    lineas_cli = doc.get("cliente_lineas") or []
    totales = [t for t in doc.get("totales", []) if t[0] != "__pagos__"]
    dest = doc.get("destacado")
    notas_txt = doc.get("notas") or []
    f_nota = F(False, 19)
    notas = []
    for n in notas_txt:
        notas.extend(envolver(n, f_nota, 880))

    header_h, row_h = 250, 58
    client_top = header_h + 36
    client_h = 100 if not lineas_cli else 130 + (34 if len(lineas_cli) > 1 else 0)
    table_top = client_top + client_h + 36
    rows_top = table_top + 54
    rows_h = max(len(filas), 1) * row_h
    tot_top = rows_top + rows_h + 40
    y_after = tot_top + 46 * len(totales) + (8 if totales else 0)
    dest_top = y_after
    if dest:
        y_after = dest_top + 96
    notas_top = y_after + 22
    if notas:
        y_after = notas_top + 28 + 27 * len(notas)
    footer_top = y_after + 48
    H = footer_top + 120

    img = Image.new("RGB", (W * S, H * S), K["white"])
    draw = ImageDraw.Draw(img)

    def R(box, fill=None, outline=None, width=1, r=0):
        b = [v * S for v in box]
        if r:
            draw.rounded_rectangle(b, radius=r * S, fill=fill, outline=outline, width=width * S)
        else:
            draw.rectangle(b, fill=fill, outline=outline, width=width * S)

    def T(x, y, txt, font, fill, align="l"):
        w = tw(txt, font)
        if align == "r":
            x -= w
        elif align == "c":
            x -= w / 2
        draw.text((x * S, y * S), txt, font=font, fill=fill)

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
    titulo = doc.get("titulo", "COMPROBANTE")
    T(W - 44, 74, titulo, F(True, 30 if len(titulo) <= 22 else 25), K["plum"], "r")
    T(W - 44, 122, datetime.now().strftime("%d/%m/%Y  ·  %H:%M"), F(False, 20), K["muted"], "r")
    T(W - 44, 152, f"Ref. {doc.get('ref', '')}", F(True, 20), K["raspberry"], "r")
    R((0, header_h, W, header_h + 6), fill=K["raspberry"])
    R((0, header_h, 150, header_h + 6), fill=K["orange"])

    # ---------- Tarjeta de cliente ----------
    R((40, client_top, W - 40, client_top + client_h), fill=K["soft"], outline=K["line"], width=2, r=20)
    T(66, client_top + 20, doc.get("cliente_rotulo", "CLIENTA"), F(True, 15), K["raspberry"])
    T(66, client_top + 44, ajustar(str(doc.get("cliente_nombre", "")), F(True, 32), 520), F(True, 32), K["text"])
    if lineas_cli:
        T(66, client_top + 90, ajustar(lineas_cli[0], F(False, 20), 870), F(False, 20), K["muted"])
    if len(lineas_cli) > 1:
        T(66, client_top + 124, ajustar(lineas_cli[1], F(False, 19), 870), F(False, 19), K["muted"])
    if doc.get("pill"):
        etq = doc["pill"]
        fp = F(True, 20)
        pw = tw(etq, fp) + 44
        R((W - 66 - pw, client_top + 43, W - 66, client_top + 87), fill=K["raspberry"], r=22)
        T(W - 66 - pw / 2, client_top + 53, etq, fp, K["white"], "c")

    # ---------- Tabla de detalle ----------
    cols = doc.get("cols", ("Descripción", "Cant.", "Estado", "Total"))
    R((40, table_top, W - 40, table_top + 52), fill=K["raspberry"], r=14)
    fh = F(True, 19)
    T(66, table_top + 14, cols[0], fh, K["white"])
    T(600, table_top + 14, cols[1], fh, K["white"], "c")
    T(670, table_top + 14, cols[2], fh, K["white"])
    T(W - 66, table_top + 14, cols[3], fh, K["white"], "r")
    if not filas:
        T(W / 2, rows_top + 16, doc.get("vacio", "Aún no hay artículos registrados"), F(False, 21), K["muted"], "c")
    fd, fb, fe = F(False, 21), F(True, 21), F(False, 19)
    for i, fila in enumerate(filas):
        y = rows_top + i * row_h
        if i % 2 == 0:
            R((40, y, W - 40, y + row_h), fill=K["zebra"])
        T(66, y + 17, ajustar(fila["desc"], fd, 480), fd, K["text"])
        T(600, y + 17, str(fila["cant"]), fd, K["text"], "c")
        T(670, y + 18, ajustar(str(fila["col3"]), fe, 170), fe, K["muted"])
        T(W - 66, y + 17, _crc(fila["total"]), fb, K["text"], "r")
    R((40, rows_top + rows_h, W - 40, rows_top + rows_h + 2), fill=K["brand"])

    # ---------- Totales ----------
    for i, t in enumerate(totales):
        etiqueta, valor, tono = t[0], t[1], t[2]
        y = tot_top + i * 46
        T(500, y + 6, etiqueta, F(False, 21), K["muted"])
        T(W - 66, y + 4, ajustar(str(valor), F(True, 24), 400), F(True, 24), _color_tono(tono), "r")
    if dest:
        fondo = {"green": K["green"], "orange": K["orange"]}.get(dest.get("tone"), K["raspberry"])
        R((470, dest_top, W - 40, dest_top + 96), fill=fondo, r=20)
        T(500, dest_top + 16, dest["label"], F(True, 17), K["white"])
        T(500, dest_top + 44, dest["value"], F(True, 38), K["white"])

    # ---------- Notas / condiciones ----------
    if notas:
        R((40, notas_top, W - 40, notas_top + 28 + 27 * len(notas)), fill=K["soft"], outline=K["line"], width=2, r=14)
        for i, linea in enumerate(notas):
            T(64, notas_top + 14 + i * 27, linea, f_nota, K["muted"])

    # ---------- Pie ----------
    R((0, footer_top, W, H), fill=K["soft"])
    R((0, footer_top, W, footer_top + 5), fill=K["brand"])
    T(W / 2, footer_top + 28, doc.get("pie", "¡Gracias por tu preferencia en Minici Store!"), F(True, 25), K["plum"], "c")
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
    filepath = os.path.join("facturas_generadas", doc.get("archivo", "comprobante.jpg"))
    img.save(filepath, "JPEG", quality=92, optimize=True)
    return filepath


# ---------- 2) Datos comunes de la clienta ----------
def _info_cliente(id_cliente):
    """Datos de la clienta listos para los comprobantes. None si no existe."""
    r = c.execute("SELECT nombre, telefono, provincia, canton, distrito, senas, metodo_envio FROM clientes WHERE id_cliente = ?", (id_cliente,)).fetchone()
    if not r:
        return None
    nombre, tel, prov, canton, distrito, senas, envio = r
    ubic = texto_ubicacion(prov, canton, distrito)
    extra = ([f"Envío: {envio}"] if envio else []) + ([ubic] if ubic else [])
    t_lineas = []
    if tel:
        t_lineas.append((f"Tel: {tel}", False))
    if envio:
        t_lineas.append((f"Envío: {envio}", False))
    if ubic:
        t_lineas.append((f"Dir: {ubic}", False))
    if senas:
        t_lineas.append((str(senas), True))
    return {"nombre": nombre, "telefono": tel, "ubic": ubic, "envio": envio, "senas": senas,
            "lineas_img": [f"Tel: {tel if tel else 'No registrado'}"] + (["   ·   ".join(extra)] if extra else []),
            "lineas_ticket": t_lineas}


# ---------- 3) Estado de cuenta (pedidos + caja + abonos) ----------
def construir_doc_estado_cuenta(id_cliente):
    cli = _info_cliente(id_cliente)
    if not cli:
        return None
    res_caja = c.execute("SELECT nombre_caja, precio_caja FROM cajas_emprendedores WHERE id_cliente = ?", (id_cliente,)).fetchone()
    nombre_caja, precio_caja = (res_caja[0], res_caja[1] or 0.0) if res_caja else (None, 0.0)
    prods = c.execute("SELECT descripcion, precio, cantidad, estado FROM productos WHERE id_cliente = ?", (id_cliente,)).fetchall()
    abonos = c.execute("SELECT monto_crc, fecha, metodo_pago FROM abonos WHERE id_cliente = ? ORDER BY id", (id_cliente,)).fetchall()
    total_abonos = sum((a[0] or 0.0) for a in abonos)
    total_articulos = sum([((p[1] or 0.0) * (p[2] if p[2] else 1)) for p in prods])
    total_compras = total_articulos + precio_caja
    saldo = total_compras - total_abonos

    filas = []
    if nombre_caja and precio_caja > 0:
        filas.append({"desc": f"{_etiqueta_caja(nombre_caja)} · cuota mensual", "cant": "1", "col3": "Mensual",
                      "unit": precio_caja, "total": precio_caja, "ticket_desc": f"{_etiqueta_caja(nombre_caja)} (cuota mensual)"})
    for p in prods:
        cant = p[2] if p[2] else 1
        filas.append({"desc": str(p[0] or ""), "cant": str(cant), "col3": str(p[3] or ""),
                      "unit": (p[1] or 0.0), "total": (p[1] or 0.0) * cant, "ticket_desc": str(p[0] or "Artículo")})

    al_dia = saldo <= 0
    return {
        "titulo": "ESTADO DE CUENTA", "ref": id_cliente, "archivo": f"factura_{id_cliente}.jpg",
        "cliente_rotulo": "CLIENTA", "cliente_nombre": cli["nombre"], "cliente_lineas": cli["lineas_img"],
        "cliente_ticket": cli["lineas_ticket"], "pill": _etiqueta_caja(nombre_caja) if nombre_caja else None,
        "cols": ("Descripción", "Cant.", "Estado", "Total"), "filas": filas,
        "totales": [("Total compras", _crc(total_compras), "normal"),
                    ("__pagos__", None, None),
                    ("Total abonado", _crc(total_abonos), "green", "-" + _crc(total_abonos))],
        "pagos_titulo": "Abonos recibidos:",
        "pagos": [(a[1], a[0], a[2]) for a in abonos],
        "destacado": {"label": "CUENTA AL DÍA" if al_dia else "SALDO PENDIENTE", "value": _crc(max(0.0, saldo)),
                      "tone": "green" if al_dia else "raspberry",
                      "ticket_centrado": "*** CUENTA AL DÍA ***" if al_dia else None},
        "telefono": cli["telefono"],
        "mensaje": _mensaje_whatsapp(cli["nombre"], prods, nombre_caja, precio_caja, total_compras, total_abonos, saldo),
        "_prods": prods,
    }


def generar_factura_imagen(id_cliente):
    """Estado de cuenta en imagen + link de WhatsApp (misma interfaz de siempre)."""
    doc = construir_doc_estado_cuenta(id_cliente)
    if not doc:
        return None, None
    return render_factura_imagen(doc), _link_wa(doc["telefono"], doc["mensaje"])


# -------------------------------------------------------------
# TICKET IMPRIMIBLE (ESTILO SUPERMERCADO, 80 mm / 58 mm)
# -------------------------------------------------------------
@functools.lru_cache(maxsize=1)
def _logo_ticket_b64():
    """Logo en blanco y negro (fondo blanco) para que salga limpio en impresoras térmicas."""
    logo, bg = _logo_factura()
    if logo is None:
        return None
    try:
        w = 380
        im = logo.resize((w, max(1, int(logo.height * w / logo.width))), Image.LANCZOS)
        diff = ImageChops.difference(im, Image.new("RGB", im.size, bg)).convert("L")
        mask = diff.point(lambda v: min(255, v * 6))
        base = Image.new("RGB", im.size, "#FFFFFF")
        base.paste(im, (0, 0), mask)
        gris = ImageOps.autocontrast(base.convert("L"), cutoff=1)
        buf = io.BytesIO()
        gris.save(buf, "PNG", optimize=True)
        return base64.b64encode(buf.getvalue()).decode("ascii")
    except Exception:
        return None


_TICKET_PLANTILLA = """<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>%%TITULO%%</title>
<style id="pg">@page { margin: 2mm; }</style>
<style>
  html, body { margin: 0; padding: 0; background: #f3f4f6; }
  .bar { text-align: center; padding: 10px 8px 4px; font-family: system-ui, -apple-system, 'Segoe UI', sans-serif; font-size: 14px; color: #4a2c3b; }
  .bar select { font-size: 14px; padding: 6px 8px; border-radius: 8px; border: 1px solid #F3B2C9; margin-right: 8px; }
  .bar button { background: #BE185D; color: #fff; border: 0; border-radius: 999px; padding: 11px 22px; font-size: 15px; font-weight: 700; cursor: pointer; }
  #t { width: 72mm; margin: 10px auto 22px; padding: 4mm 3mm 8mm; background: #fff; color: #000;
       font-family: 'Courier New', Courier, 'Liberation Mono', monospace; font-size: 12.5px; line-height: 1.38;
       box-shadow: 0 1px 8px rgba(0,0,0,.25); box-sizing: border-box; }
  #t.p58 { font-size: 11px; }
  #t img { display: block; margin: 0 auto 4px; width: 46mm; max-width: 80%; }
  #t.p58 img { width: 34mm; }
  .c { text-align: center; } .b { font-weight: 700; }
  .nombre { font-size: 1.45em; font-weight: 700; letter-spacing: 1px; text-align: center; }
  .sep { border-top: 1px dashed #000; margin: 7px 0; }
  .sep2 { border-top: 2px solid #000; margin: 7px 0; }
  .r { display: flex; justify-content: space-between; gap: 8px; }
  .r span:last-child { white-space: nowrap; text-align: right; }
  .it { margin-bottom: 5px; } .it .d { font-weight: 700; word-break: break-word; }
  .big { font-size: 1.3em; font-weight: 700; }
  .dim { opacity: .85; }
  @media print {
    html, body { background: #fff; }
    .bar { display: none !important; }
    #t { width: auto !important; margin: 0 !important; padding: 0 0 8mm !important; box-shadow: none !important; }
  }
</style></head>
<body>
<div class="bar">
  <label>Papel: <select id="papel"><option value="80">80 mm (estándar)</option><option value="58">58 mm (pequeño)</option></select></label>
  <button onclick="imprimir()">🖨️ Imprimir ticket</button>
</div>
<div id="t">%%CUERPO%%</div>
<script>
  function aplicar() {
    var mm = parseInt(document.getElementById('papel').value, 10);
    var t = document.getElementById('t');
    t.style.width = (mm - 8) + 'mm';
    t.className = mm < 70 ? 'p58' : '';
  }
  function imprimir() { aplicar(); setTimeout(function () { window.print(); }, 80); }
  document.getElementById('papel').addEventListener('change', aplicar);
</script>
</body></html>"""


def render_ticket_html(doc):
    """Ticket tipo supermercado (80 mm / 58 mm) para cualquier comprobante. Devuelve (html, alto_sugerido)."""
    E = html.escape

    def fila(izq, der, clase=""):
        return f'<div class="r {clase}"><span>{E(str(izq))}</span><span>{E(str(der))}</span></div>'

    b = []
    logo64 = _logo_ticket_b64()
    if logo64:  # el logo ya trae el nombre y "Personal Shopper"
        b.append(f'<img alt="Minici Store" src="data:image/png;base64,{logo64}">')
    else:
        b.append('<div class="nombre">MINICI STORE</div><div class="c dim">PERSONAL SHOPPER</div>')
    b.append(f'<div class="sep2"></div><div class="c b">{E(doc.get("titulo", "COMPROBANTE"))}</div><div class="sep"></div>')
    b.append(fila("Fecha:", datetime.now().strftime("%d/%m/%Y %H:%M")))
    b.append(fila("Ref:", doc.get("ref", "")))
    b.append(f'<div class="b" style="margin-top:3px">{E(str(doc.get("cliente_nombre", "")))}</div>')
    for texto, tenue in doc.get("cliente_ticket", []):
        b.append(f'<div class="{"dim" if tenue else ""}">{E(str(texto))}</div>')

    cols = doc.get("cols", ("Descripción", "Cant.", "Estado", "Total"))
    filas = doc.get("filas", [])
    b.append('<div class="sep"></div>')
    b.append(fila(str(cols[0]).upper(), "TOTAL", "b"))
    b.append('<div class="sep"></div>')
    if not filas:
        b.append(f'<div class="c dim">{E(doc.get("vacio", "Sin artículos registrados"))}</div>')
    for f in filas:
        detalle = f"{f['cant']} x {_crc(f['unit'])}"
        b.append('<div class="it"><div class="d">' + E(f.get("ticket_desc", f["desc"])) + '</div>'
                 + fila(detalle, _crc(f["total"])) + '</div>')

    b.append('<div class="sep"></div>')
    pagos = doc.get("pagos", [])
    for t in doc.get("totales", []):
        if t[0] == "__pagos__":
            if pagos:
                b.append(f'<div style="margin-top:4px" class="dim">{E(doc.get("pagos_titulo", "Pagos:"))}</div>')
                for fecha, monto, metodo in pagos[-6:]:
                    etiqueta = f"  {_fecha_corta(fecha)}" + (f" {metodo}" if metodo else "")
                    b.append(fila(etiqueta, ("-" if (monto or 0) >= 0 else "+") + _crc(abs(monto or 0))))
                if len(pagos) > 6:
                    b.append(f'<div class="dim">  (+{len(pagos) - 6} movimientos anteriores)</div>')
            continue
        valor = t[3] if len(t) > 3 else t[1]
        b.append(fila(str(t[0]).upper(), valor, "" if t[2] == "muted" else "b"))

    b.append('<div class="sep2"></div>')
    d = doc.get("destacado")
    if d:
        if d.get("ticket_centrado"):
            b.append(f'<div class="c big">{E(d["ticket_centrado"])}</div>')
        else:
            b.append(fila(d["label"], d["value"], "big"))
        b.append('<div class="sep2"></div>')
    for nota in doc.get("notas", []):
        b.append(f'<div class="dim" style="margin-top:3px">{E(str(nota))}</div>')
    b.append('<div class="c b" style="margin-top:6px">¡GRACIAS POR SU COMPRA!</div>')
    b.append('<div class="c dim">Conserve este comprobante</div>')
    b.append('<div class="c dim" style="margin-top:2px">Minici Store · Personal Shopper</div>')

    cuerpo = _TICKET_PLANTILLA.replace("%%TITULO%%", E(f"Ticket {doc.get('ref', '')}")).replace("%%CUERPO%%", "\n".join(b))
    extra_cli = 22 * len(doc.get("cliente_ticket", []))
    alto = 560 + 62 * len(filas) + 22 * min(len(pagos), 6) + extra_cli + 40 * len(doc.get("notas", [])) + 26 * len(doc.get("totales", []))
    return cuerpo, alto


def generar_ticket_html(id_cliente):
    """Ticket del estado de cuenta (misma interfaz de siempre). Devuelve (html, alto) o (None, 0)."""
    doc = construir_doc_estado_cuenta(id_cliente)
    if not doc:
        return None, 0
    return render_ticket_html(doc)


def preparar_documento(doc):
    """Genera imagen + ticket + link de WhatsApp de un comprobante y lo deja listo para mostrar."""
    ruta = render_factura_imagen(doc)
    ticket, alto = render_ticket_html(doc)
    return {"titulo": doc.get("titulo", "Comprobante"), "ref": doc.get("ref", ""), "path": ruta,
            "archivo": os.path.basename(ruta), "ticket": ticket, "alto": alto,
            "link": _link_wa(doc.get("telefono"), doc.get("mensaje", "")) if doc.get("telefono") else None}


def mostrar_documento(info, key):
    """Muestra un comprobante: descargar imagen, WhatsApp e impresión de ticket."""
    import streamlit.components.v1 as components
    if not info:
        return
    st.markdown(f"#### 🧾 {info['titulo']} · {info['ref']}")
    if info.get("path") and os.path.exists(info["path"]):
        with open(info["path"], "rb") as f:
            st.download_button("🖼️ Descargar comprobante en imagen", data=f.read(), file_name=info["archivo"],
                               mime="image/jpeg", key=f"dl_img_{key}")
    if info.get("link"):
        st.markdown(f'<a href="{info["link"]}" target="_blank" class="btn-whatsapp">📲 Enviar por WhatsApp</a>', unsafe_allow_html=True)
    with st.expander("🖨️ Ticket para imprimir (estilo supermercado)", expanded=False):
        st.caption("Toca «Imprimir ticket», elige tu impresora (80 mm o 58 mm) y deja los márgenes en «ninguno».")
        components.html(info["ticket"], height=info["alto"], scrolling=True)
        st.download_button("⬇️ Descargar ticket (se abre en el navegador para imprimir)", data=info["ticket"].encode("utf-8"),
                           file_name=f"Ticket_{info['ref']}.html", mime="text/html", key=f"dl_tk_{key}")


def mostrar_ticket(id_cliente, key, con_checkbox=False):
    """Estado de cuenta del cliente: imagen, WhatsApp y ticket. Con checkbox queda oculto hasta que se pida."""
    if con_checkbox and not st.checkbox("🧾 Ver / imprimir ticket (estilo supermercado)", key=f"chk_ticket_{key}"):
        return
    doc = construir_doc_estado_cuenta(id_cliente)
    if not doc:
        return
    import streamlit.components.v1 as components
    ticket, alto = render_ticket_html(doc)
    st.caption("🧾 Ticket listo: toca «Imprimir ticket», elige tu impresora y deja los márgenes en «ninguno».")
    components.html(ticket, height=alto, scrolling=True)
    st.download_button("⬇️ Descargar ticket (se abre en el navegador para imprimir)", data=ticket.encode("utf-8"),
                       file_name=f"Ticket_{id_cliente}.html", mime="text/html", key=f"dl_ticket_{key}")


# -------------------------------------------------------------
# VENTAS (POS) Y APARTADOS: REGLAS, BASE DE DATOS Y COMPROBANTES
# -------------------------------------------------------------
METODOS_PAGO = ["Efectivo", "Tarjeta", "SINPE Móvil"]
ICONO_PAGO = {"Efectivo": "💵", "Tarjeta": "💳", "SINPE Móvil": "📲", "Apartado": "🕒"}


def _ahora():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _arriba(x):
    """Redondea hacia arriba al colón entero."""
    return float(int(x) if float(x) == int(x) else int(x) + 1)


def politica_apartado():
    def num(clave, defecto):
        try:
            return float(cfg_get(clave, defecto))
        except Exception:
            return float(defecto)
    return {"prima_pct": num("apartado_prima_pct", 20), "plazo_dias": int(num("apartado_plazo_dias", 30)),
            "cancel_pct": num("apartado_cancel_pct", 10)}


def guardar_politica_apartado(prima_pct, plazo_dias, cancel_pct):
    cfg_set("apartado_prima_pct", prima_pct)
    cfg_set("apartado_plazo_dias", int(plazo_dias))
    cfg_set("apartado_cancel_pct", cancel_pct)


def _validar_stock(carrito):
    """Revisa que haya existencias de todo lo del carrito. Devuelve un mensaje de error o None."""
    pedido = {}
    for it in carrito:
        if it.get("id_prod"):
            pedido[it["id_prod"]] = pedido.get(it["id_prod"], 0) + int(it["cant"])
    for id_prod, cant in pedido.items():
        fila = c.execute("SELECT cantidad, descripcion FROM productos WHERE id = ?", (id_prod,)).fetchone()
        if not fila:
            return "Un producto del carrito ya no existe en el inventario."
        if int(fila[0] or 0) < cant:
            return f"Stock insuficiente de «{fila[1]}»: hay {int(fila[0] or 0)} y pides {cant}."
    return None


def _total_carrito(carrito):
    return float(sum(int(i["cant"]) * float(i["precio"]) for i in carrito))


# ---------- Ventas ----------
def crear_venta(carrito, descuento, metodo, recibido, referencia, id_cliente=None, nombre_libre=None):
    """Registra una venta pagada de una vez. Devuelve (id_venta, mensaje_de_error)."""
    if not carrito:
        return None, "El carrito está vacío."
    if metodo not in METODOS_PAGO:
        return None, "Elige un método de pago."
    subtotal = _total_carrito(carrito)
    descuento = min(max(0.0, float(descuento or 0)), subtotal)
    total = subtotal - descuento
    if metodo == "Efectivo":
        recibido = float(recibido or 0)
        if recibido + 0.001 < total:
            return None, "El monto recibido es menor al total de la venta."
        vuelto = recibido - total
    else:
        recibido, vuelto = total, 0.0
    err = _validar_stock(carrito)
    if err:
        return None, err
    try:
        cur = c.execute(
            "INSERT INTO ventas (fecha, id_cliente, cliente_nombre, subtotal, descuento, total, metodo_pago, recibido, vuelto, referencia, estado) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pagada')",
            (_ahora(), id_cliente, (nombre_libre or "").strip() or None, subtotal, descuento, total, metodo, recibido, vuelto,
             (referencia or "").strip() or None))
        id_venta = cur.lastrowid
        c.execute("UPDATE ventas SET folio = ? WHERE id = ?", (f"V-{id_venta:05d}", id_venta))
        restante = descuento
        for idx, it in enumerate(carrito):
            bruto = int(it["cant"]) * float(it["precio"])
            if idx == len(carrito) - 1:
                d = min(restante, bruto)
            else:
                d = min(float(round(descuento * bruto / subtotal)) if subtotal else 0.0, restante, bruto)
            restante -= d
            c.execute("INSERT INTO ventas_rapidas (fecha, id_cliente, producto, tipo, cantidad, precio_u, descuento, total, id_venta, metodo_pago, id_producto) "
                      "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                      (_ahora(), id_cliente, cap1(it["desc"]), it.get("origen", "Del Stock General"), int(it["cant"]),
                       float(it["precio"]), d, bruto - d, id_venta, metodo, it.get("id_prod")))
            if it.get("id_prod"):
                c.execute("UPDATE productos SET cantidad = cantidad - ? WHERE id = ?", (int(it["cant"]), it["id_prod"]))
        conn.commit()
    except Exception as e:
        conn.rollback()
        return None, f"No se pudo registrar la venta: {e}"
    return id_venta, None


def anular_venta(id_venta):
    """Anula una venta: devuelve las existencias y deja de contar en las finanzas. Devuelve (ok, texto)."""
    v = c.execute("SELECT estado FROM ventas WHERE id = ?", (id_venta,)).fetchone()
    if not v:
        return False, "No se encontró la venta."
    if v[0] != "Pagada":
        return False, "Esa venta ya estaba anulada."
    try:
        for id_prod, cant in c.execute("SELECT id_producto, cantidad FROM ventas_rapidas WHERE id_venta = ? AND id_producto IS NOT NULL", (id_venta,)).fetchall():
            c.execute("UPDATE productos SET cantidad = cantidad + ? WHERE id = ?", (cant, id_prod))
        c.execute("UPDATE ventas SET estado = 'Anulada' WHERE id = ?", (id_venta,))
        conn.commit()
    except Exception as e:
        conn.rollback()
        return False, f"No se pudo anular: {e}"
    return True, "Venta anulada y existencias devueltas al inventario."


# ---------- Apartados ----------
def apartado_info(id_apartado):
    a = c.execute("SELECT id, folio, fecha, fecha_limite, id_cliente, total, prima_pct, cancelacion_pct, estado, notas, fecha_cierre FROM apartados WHERE id = ?", (id_apartado,)).fetchone()
    if not a:
        return None
    items = c.execute("SELECT id_producto, descripcion, cantidad, precio_u FROM apartado_items WHERE id_apartado = ? ORDER BY id", (id_apartado,)).fetchall()
    pagos = c.execute("SELECT id, fecha, monto, metodo_pago, referencia, tipo FROM apartado_pagos WHERE id_apartado = ? ORDER BY id", (id_apartado,)).fetchall()
    entradas = sum((p[2] or 0.0) for p in pagos if (p[2] or 0.0) > 0)
    salidas = -sum((p[2] or 0.0) for p in pagos if (p[2] or 0.0) < 0)
    neto = entradas - salidas
    estado, total = a[8], a[5] or 0.0
    saldo = max(0.0, total - neto) if estado == "Activo" else 0.0
    hoy = datetime.now().strftime("%Y-%m-%d")
    vencido = estado == "Activo" and saldo > 0.5 and bool(a[3]) and a[3] < hoy
    return {"id": a[0], "folio": a[1], "fecha": a[2], "limite": a[3], "id_cliente": a[4], "total": total,
            "prima_pct": a[6], "cancel_pct": a[7], "estado": estado, "notas": a[9], "cierre": a[10],
            "items": items, "pagos": pagos, "pagado": entradas, "reembolsado": salidas, "neto": neto,
            "saldo": saldo, "vencido": vencido}


def crear_apartado(carrito, id_cliente, dias_plazo, prima, metodo, referencia="", notas=""):
    """Reserva los productos y cobra la prima. Devuelve (id_apartado, mensaje_de_error)."""
    if not carrito:
        return None, "El carrito está vacío."
    if not id_cliente:
        return None, "Un apartado necesita una clienta registrada."
    if metodo not in METODOS_PAGO:
        return None, "Elige el método de pago de la prima."
    pol = politica_apartado()
    total = _total_carrito(carrito)
    prima = float(prima or 0)
    prima_min = _arriba(total * pol["prima_pct"] / 100.0)
    if prima + 0.001 < prima_min:
        return None, f"La prima mínima es ₡{prima_min:,.0f} ({pol['prima_pct']:g}% del total)."
    if prima > total + 0.001:
        return None, "La prima no puede ser mayor al total."
    err = _validar_stock(carrito)
    if err:
        return None, err
    try:
        limite = (datetime.now() + timedelta(days=int(dias_plazo))).strftime("%Y-%m-%d")
        cur = c.execute("INSERT INTO apartados (fecha, fecha_limite, id_cliente, total, prima_pct, cancelacion_pct, estado, notas) VALUES (?, ?, ?, ?, ?, ?, 'Activo', ?)",
                        (_ahora(), limite, id_cliente, total, pol["prima_pct"], pol["cancel_pct"], notas or ""))
        id_ap = cur.lastrowid
        c.execute("UPDATE apartados SET folio = ? WHERE id = ?", (f"A-{id_ap:04d}", id_ap))
        for it in carrito:
            c.execute("INSERT INTO apartado_items (id_apartado, id_producto, descripcion, cantidad, precio_u) VALUES (?, ?, ?, ?, ?)",
                      (id_ap, it.get("id_prod"), cap1(it["desc"]), int(it["cant"]), float(it["precio"])))
            if it.get("id_prod"):
                c.execute("UPDATE productos SET cantidad = cantidad - ? WHERE id = ?", (int(it["cant"]), it["id_prod"]))
        if prima > 0:
            c.execute("INSERT INTO apartado_pagos (id_apartado, fecha, monto, metodo_pago, referencia, tipo) VALUES (?, ?, ?, ?, ?, 'Prima')",
                      (id_ap, _ahora(), prima, metodo, (referencia or "").strip() or None))
        conn.commit()
    except Exception as e:
        conn.rollback()
        return None, f"No se pudo crear el apartado: {e}"
    return id_ap, None


def registrar_pago_apartado(id_apartado, monto, metodo, referencia=""):
    """Suma un abono al apartado. Devuelve (id_pago, mensaje_de_error)."""
    info = apartado_info(id_apartado)
    if not info:
        return None, "No se encontró el apartado."
    if info["estado"] != "Activo":
        return None, f"El apartado está {info['estado'].lower()}: no recibe más pagos."
    if metodo not in METODOS_PAGO:
        return None, "Elige el método de pago."
    monto = float(monto or 0)
    if monto <= 0:
        return None, "El abono debe ser mayor a 0."
    if monto > info["saldo"] + 0.001:
        return None, f"El abono no puede superar el saldo (₡{info['saldo']:,.0f})."
    try:
        cur = c.execute("INSERT INTO apartado_pagos (id_apartado, fecha, monto, metodo_pago, referencia, tipo) VALUES (?, ?, ?, ?, ?, 'Abono')",
                        (id_apartado, _ahora(), monto, metodo, (referencia or "").strip() or None))
        conn.commit()
        return cur.lastrowid, None
    except Exception as e:
        conn.rollback()
        return None, f"No se pudo registrar el abono: {e}"


def entregar_apartado(id_apartado):
    info = apartado_info(id_apartado)
    if not info:
        return False, "No se encontró el apartado."
    if info["estado"] != "Activo":
        return False, f"El apartado ya está {info['estado'].lower()}."
    if info["saldo"] > 0.5:
        return False, f"Aún hay un saldo de ₡{info['saldo']:,.0f}: no se puede entregar."
    c.execute("UPDATE apartados SET estado = 'Entregado', fecha_cierre = ? WHERE id = ?", (_ahora(), id_apartado))
    conn.commit()
    return True, "Apartado entregado."


def cargo_cancelacion_sugerido(info):
    """Cargo por cancelación según la política vigente al crear el apartado (nunca más de lo pagado)."""
    return min(info["neto"], _arriba(info["total"] * (info["cancel_pct"] or 0) / 100.0))


def cancelar_apartado(id_apartado, cargo, metodo_reembolso):
    """Cancela, devuelve los productos al inventario y reembolsa lo pagado menos el cargo. Devuelve (ok, texto)."""
    info = apartado_info(id_apartado)
    if not info:
        return False, "No se encontró el apartado."
    if info["estado"] != "Activo":
        return False, f"El apartado ya está {info['estado'].lower()}."
    cargo = min(max(0.0, float(cargo or 0)), info["neto"])
    reembolso = info["neto"] - cargo
    if reembolso > 0 and metodo_reembolso not in METODOS_PAGO:
        return False, "Elige cómo se devuelve el dinero."
    try:
        if reembolso > 0:
            c.execute("INSERT INTO apartado_pagos (id_apartado, fecha, monto, metodo_pago, referencia, tipo) VALUES (?, ?, ?, ?, ?, 'Reembolso')",
                      (id_apartado, _ahora(), -reembolso, metodo_reembolso, None))
        for id_prod, _desc, cant, _pu in info["items"]:
            if id_prod:
                c.execute("UPDATE productos SET cantidad = cantidad + ? WHERE id = ?", (cant, id_prod))
        c.execute("UPDATE apartados SET estado = 'Cancelado', fecha_cierre = ? WHERE id = ?", (_ahora(), id_apartado))
        conn.commit()
    except Exception as e:
        conn.rollback()
        return False, f"No se pudo cancelar: {e}"
    return True, f"Apartado cancelado. Se retuvo ₡{cargo:,.0f} y se reembolsó ₡{reembolso:,.0f}."


def ampliar_plazo_apartado(id_apartado, nueva_fecha):
    info = apartado_info(id_apartado)
    if not info or info["estado"] != "Activo":
        return False, "Solo se puede ampliar el plazo de un apartado activo."
    c.execute("UPDATE apartados SET fecha_limite = ? WHERE id = ?", (nueva_fecha.strftime("%Y-%m-%d"), id_apartado))
    conn.commit()
    return True, "Plazo actualizado."


def resumen_apartados():
    """Números para finanzas: dinero retenido, por cobrar y cargos conservados."""
    activos = c.execute("SELECT id FROM apartados WHERE estado = 'Activo'").fetchall()
    por_cobrar = retenido = 0.0
    for (i,) in activos:
        inf = apartado_info(i)
        por_cobrar += inf["saldo"]
        retenido += inf["neto"]
    total_vigente = c.execute("SELECT COALESCE(SUM(total), 0) FROM apartados WHERE estado IN ('Activo', 'Entregado')").fetchone()[0] or 0.0
    cargos = c.execute("SELECT COALESCE(SUM(p.monto), 0) FROM apartado_pagos p JOIN apartados a ON a.id = p.id_apartado WHERE a.estado = 'Cancelado'").fetchone()[0] or 0.0
    cobrado_neto = c.execute("SELECT COALESCE(SUM(monto), 0) FROM apartado_pagos").fetchone()[0] or 0.0
    return {"activos": len(activos), "por_cobrar": por_cobrar, "dinero_retenido": retenido,
            "total_vigente": total_vigente, "cargos_retenidos": cargos, "cobrado_neto": cobrado_neto}


# ---------- Comprobantes de ventas y apartados ----------
def _mensaje_venta(nombre, folio, fecha, items, subtotal, descuento, total, metodo, referencia):
    linea = "━━━━━━━━━━━━━━"
    m = "🛍️ *MINICI STORE* ✨\n_Personal Shopper_\n" + linea + "\n\n"
    m += f"¡Hola *{nombre}*! 💖\nGracias por tu compra del *{_fecha_corta(fecha)}*.\n\n🧾 *Comprobante {folio}*\n"
    for desc, cant, pu in items[:6]:
        x = f" ×{cant}" if cant > 1 else ""
        m += f"• {_recortar_msg(desc, 28)}{x} — {_crc(cant * pu)}\n"
    if len(items) > 6:
        m += f"• _y {len(items) - 6} más en tu comprobante_\n"
    m += "\n" + linea + "\n"
    if descuento > 0:
        m += f"🏷️ Descuento: -{_crc(descuento)}\n"
    m += f"💰 *Total: {_crc(total)}*\n{ICONO_PAGO.get(metodo, '💳')} Pagado con: {metodo}\n"
    if referencia:
        m += f"🔖 Ref.: {referencia}\n"
    m += linea + "\n\nTe comparto tu comprobante en imagen 🖼️\n¡Gracias por confiar en Minici Store! 🌸"
    return m


def construir_doc_venta(id_venta):
    v = c.execute("SELECT id, folio, fecha, id_cliente, cliente_nombre, subtotal, descuento, total, metodo_pago, recibido, vuelto, referencia, estado FROM ventas WHERE id = ?", (id_venta,)).fetchone()
    if not v:
        return None
    _id, folio, fecha, id_cli, nombre_libre, subtotal, descuento, total, metodo, recibido, vuelto, ref, estado = v
    items = c.execute("SELECT producto, cantidad, precio_u FROM ventas_rapidas WHERE id_venta = ? ORDER BY id", (id_venta,)).fetchall()
    cli = _info_cliente(id_cli) if id_cli else None
    nombre = cli["nombre"] if cli else (nombre_libre or "Venta general")
    filas = [{"desc": str(d), "cant": str(q), "col3": _crc(pu), "unit": pu, "total": q * pu, "ticket_desc": str(d)} for d, q, pu in items]
    totales = [("Subtotal", _crc(subtotal), "normal")]
    if (descuento or 0) > 0:
        totales.append(("Descuento", "-" + _crc(descuento), "red"))
    totales.append(("Método de pago", metodo or "—", "normal"))
    if metodo == "Efectivo":
        totales.append(("Recibido", _crc(recibido), "muted"))
        totales.append(("Vuelto", _crc(vuelto), "muted"))
    if ref:
        totales.append(("Referencia", str(ref), "muted"))
    anulada = estado != "Pagada"
    return {
        "titulo": "VENTA ANULADA" if anulada else "COMPROBANTE DE VENTA", "ref": folio or f"V-{id_venta:05d}",
        "archivo": f"venta_{folio or id_venta}.jpg", "cliente_rotulo": "CLIENTA" if cli else "CLIENTE",
        "cliente_nombre": nombre, "cliente_lineas": cli["lineas_img"] if cli else [],
        "cliente_ticket": cli["lineas_ticket"] if cli else [], "pill": estado if anulada else (metodo or None),
        "cols": ("Descripción", "Cant.", "P. Unit.", "Total"), "filas": filas, "totales": totales,
        "destacado": {"label": "VENTA ANULADA" if anulada else "TOTAL PAGADO", "value": _crc(total),
                      "tone": "orange" if anulada else "green"},
        "telefono": cli["telefono"] if cli else None,
        "mensaje": _mensaje_venta(nombre, folio or "", fecha, items, subtotal, descuento or 0, total, metodo, ref),
    }


def _mensaje_apartado(modo, nombre, inf, pago=None, cargo=0.0, reembolso=0.0):
    linea = "━━━━━━━━━━━━━━"
    m = "🛍️ *MINICI STORE* ✨\n_Personal Shopper_\n" + linea + "\n\n" + f"¡Hola *{nombre}*! 💖\n"
    if modo == "apertura":
        m += f"Tu apartado *{inf['folio']}* quedó registrado 🕒\n\n"
        for _i, desc, cant, pu in inf["items"][:6]:
            x = f" ×{cant}" if cant > 1 else ""
            m += f"• {_recortar_msg(desc, 28)}{x} — {_crc(cant * pu)}\n"
        m += f"\n💰 Total: {_crc(inf['total'])}\n✅ Pagado: {_crc(inf['neto'])}\n📌 *Saldo: {_crc(inf['saldo'])}*\n"
        m += f"📅 Fecha límite: *{_fecha_corta(inf['limite'])}*\n\nTus artículos quedan reservados hasta completar el pago (sin intereses). Puedes abonar cuando quieras por efectivo, tarjeta o SINPE Móvil."
    elif modo == "abono":
        m += f"Recibimos tu abono de *{_crc(pago[2])}* ({pago[3]}) en el apartado *{inf['folio']}* ✅\n\n"
        m += f"✅ Pagado: {_crc(inf['neto'])} de {_crc(inf['total'])}\n"
        m += (f"📌 *Saldo: {_crc(inf['saldo'])}*\n📅 Fecha límite: *{_fecha_corta(inf['limite'])}*" if inf["saldo"] > 0.5
              else "🎉 *¡Tu apartado está pagado completo!* Ya puedes pasar a retirarlo.")
    elif modo == "entrega":
        m += f"🎉 Tu apartado *{inf['folio']}* fue entregado y está pagado completo ({_crc(inf['total'])}).\n"
    elif modo == "cancelacion":
        m += f"Tu apartado *{inf['folio']}* fue cancelado.\n\n💰 Pagado: {_crc(inf['pagado'])}\n🏷️ Cargo por cancelación: {_crc(cargo)}\n↩️ Reembolso: {_crc(reembolso)}\n"
    else:
        m += f"Este es el estado de tu apartado *{inf['folio']}*:\n\n💰 Total: {_crc(inf['total'])}\n✅ Pagado: {_crc(inf['neto'])}\n📌 *Saldo: {_crc(inf['saldo'])}*\n📅 Fecha límite: *{_fecha_corta(inf['limite'])}*"
    m += "\n\n" + linea + "\nTe comparto tu comprobante en imagen 🖼️\n¡Gracias por confiar en Minici Store! 🌸"
    return m


TITULOS_APARTADO = {"apertura": "COMPROBANTE DE APARTADO", "abono": "RECIBO DE ABONO · APARTADO",
                    "entrega": "APARTADO ENTREGADO", "cancelacion": "APARTADO CANCELADO", "estado": "ESTADO DE APARTADO"}


def construir_doc_apartado(id_apartado, modo="estado", id_pago=None):
    inf = apartado_info(id_apartado)
    if not inf:
        return None
    cli = _info_cliente(inf["id_cliente"]) or {"nombre": "Cliente", "telefono": None, "lineas_img": [], "lineas_ticket": []}
    pago = next((p for p in inf["pagos"] if p[0] == id_pago), None)
    filas = [{"desc": d, "cant": str(q), "col3": _crc(pu), "unit": pu, "total": q * pu, "ticket_desc": d} for _i, d, q, pu in inf["items"]]
    pagos_lista = [(p[1], p[2], p[3]) for p in inf["pagos"]]

    totales = [("Total del apartado", _crc(inf["total"]), "normal"), ("__pagos__", None, None)]
    cargo = reembolso = 0.0
    if inf["estado"] == "Cancelado":
        reembolso = inf["reembolsado"]
        cargo = inf["pagado"] - inf["reembolsado"]
        totales += [("Total pagado", _crc(inf["pagado"]), "green", "-" + _crc(inf["pagado"])),
                    ("Cargo retenido", _crc(cargo), "red")]
        destacado = {"label": "REEMBOLSO AL CLIENTE", "value": _crc(reembolso), "tone": "orange"}
    else:
        totales.append(("Total pagado", _crc(inf["neto"]), "green", "-" + _crc(inf["neto"])))
        if modo == "abono" and pago:
            totales.insert(1, ("Abono de hoy", _crc(pago[2]), "green"))
        if inf["estado"] == "Entregado":
            destacado = {"label": "APARTADO ENTREGADO", "value": _crc(inf["total"]), "tone": "green"}
        elif inf["saldo"] > 0.5:
            destacado = {"label": "SALDO PENDIENTE", "value": _crc(inf["saldo"]), "tone": "raspberry"}
        else:
            destacado = {"label": "PAGADO COMPLETO", "value": "Listo para entregar", "tone": "green",
                         "ticket_centrado": "*** PAGADO COMPLETO ***"}

    notas = []
    if inf["estado"] == "Activo":
        notas.append(f"Fecha límite de pago: {_fecha_corta(inf['limite'])}" + ("  (VENCIDO)" if inf["vencido"] else ""))
    if modo == "abono" and pago:
        notas.append(f"Método de pago: {pago[3]}" + (f" · Ref. {pago[4]}" if pago[4] else ""))
    if modo == "apertura":
        notas.append("Tus artículos quedan reservados hasta completar el pago. Sin intereses.")
        notas.append(f"Si el apartado se cancela o no se completa, se retiene un cargo del {inf['cancel_pct']:g}% del total (máximo lo pagado) y se devuelve el resto.")
    if inf["estado"] == "Entregado" and inf["cierre"]:
        notas.append(f"Entregado el {_fecha_corta(inf['cierre'])}. ¡Gracias por tu compra!")
    if inf["estado"] == "Cancelado":
        notas.append(f"Cancelado el {_fecha_corta(inf['cierre'])}.")

    estado_txt = "Vencido" if inf["vencido"] else inf["estado"]
    return {
        "titulo": TITULOS_APARTADO.get(modo, "ESTADO DE APARTADO"), "ref": inf["folio"],
        "archivo": f"apartado_{inf['folio']}_{modo}.jpg", "cliente_rotulo": "CLIENTA", "cliente_nombre": cli["nombre"],
        "cliente_lineas": cli["lineas_img"], "cliente_ticket": cli["lineas_ticket"], "pill": estado_txt,
        "cols": ("Descripción", "Cant.", "P. Unit.", "Total"), "filas": filas, "totales": totales,
        "pagos_titulo": "Pagos y movimientos:", "pagos": pagos_lista, "destacado": destacado, "notas": notas,
        "telefono": cli["telefono"], "mensaje": _mensaje_apartado(modo, cli["nombre"], inf, pago, cargo, reembolso),
    }



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
# DIRECCIÓN (COSTA RICA) Y MÉTODO DE ENVÍO
# ==========================================
UBICACIONES_CR = {
    "San José": ["San José", "Escazú", "Desamparados", "Puriscal", "Tarrazú", "Aserrí", "Mora", "Goicoechea",
                 "Santa Ana", "Alajuelita", "Vázquez de Coronado", "Acosta", "Tibás", "Moravia", "Montes de Oca",
                 "Turrubares", "Dota", "Curridabat", "Pérez Zeledón", "León Cortés Castro"],
    "Alajuela": ["Alajuela", "San Ramón", "Grecia", "San Mateo", "Atenas", "Naranjo", "Palmares", "Poás", "Orotina",
                 "San Carlos", "Zarcero", "Sarchí", "Upala", "Los Chiles", "Guatuso", "Río Cuarto"],
    "Cartago": ["Cartago", "Paraíso", "La Unión", "Jiménez", "Turrialba", "Alvarado", "Oreamuno", "El Guarco"],
    "Heredia": ["Heredia", "Barva", "Santo Domingo", "Santa Bárbara", "San Rafael", "San Isidro", "Belén", "Flores",
                "San Pablo", "Sarapiquí"],
    "Guanacaste": ["Liberia", "Nicoya", "Santa Cruz", "Bagaces", "Carrillo", "Cañas", "Abangares", "Tilarán",
                   "Nandayure", "La Cruz", "Hojancha"],
    "Puntarenas": ["Puntarenas", "Esparza", "Buenos Aires", "Montes de Oro", "Osa", "Quepos", "Golfito", "Coto Brus",
                   "Parrita", "Corredores", "Garabito", "Monteverde", "Puerto Jiménez"],
    "Limón": ["Limón", "Pococí", "Siquirres", "Talamanca", "Matina", "Guácimo"],
}
METODOS_ENVIO = ["Correos de Costa Rica", "Encomiendas Tabo"]


def texto_ubicacion(provincia, canton, distrito):
    return ", ".join([str(x).strip() for x in (provincia, canton, distrito) if x and str(x).strip()])


def selector_ubicacion(key_prefix, actual=None):
    """Provincia → Cantón (según la provincia) → Distrito (+ otras señas). Devuelve (prov, cantón, distrito, señas)."""
    actual = actual or {}
    provs = [""] + list(UBICACIONES_CR.keys())
    p0 = actual.get("provincia") or ""
    col_p, col_c, col_d = st.columns(3)
    with col_p:
        prov = st.selectbox("Provincia", provs, index=provs.index(p0) if p0 in provs else 0,
                            format_func=lambda x: x or "Seleccionar…", key=f"{key_prefix}_prov")
    cantones = [""] + (UBICACIONES_CR.get(prov, []) if prov else [])
    c0 = actual.get("canton") or ""
    with col_c:
        canton = st.selectbox("Cantón", cantones, index=cantones.index(c0) if c0 in cantones else 0,
                              format_func=lambda x: x or ("Seleccionar…" if prov else "Elige la provincia"),
                              key=f"{key_prefix}_canton_{prov or 'x'}", disabled=not prov)
    with col_d:
        distrito = st.text_input("Distrito", value=actual.get("distrito") or "", placeholder="Ej. San Rafael",
                                 key=f"{key_prefix}_distrito")
    senas = st.text_input("Otras señas (opcional)", value=actual.get("senas") or "",
                          placeholder="Ej. 200 m norte de la iglesia, casa blanca", key=f"{key_prefix}_senas")
    return prov, canton, titulo_es(distrito), senas.strip()


def selector_envio(key, actual=None):
    return st.radio("Método de envío", METODOS_ENVIO,
                    index=METODOS_ENVIO.index(actual) if actual in METODOS_ENVIO else 0,
                    horizontal=True, key=key)


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
# PANTALLAS DEL POS: VENTA, APARTADOS E HISTORIAL
# -------------------------------------------------------------
_OPC_PAGO = ["💵 Efectivo", "💳 Tarjeta", "📲 SINPE Móvil"]
_OPC_PAGO_APARTADO = _OPC_PAGO + ["🕒 Apartado"]


def _metodo_de(opcion):
    return str(opcion).split(" ", 1)[1] if " " in str(opcion) else str(opcion)


def _hint_referencia(metodo):
    return {"Tarjeta": "Últimos 4 dígitos o N.º de autorización (opcional)",
            "SINPE Móvil": "N.º de comprobante SINPE (opcional)"}.get(metodo)


def cobros_por_metodo(desde=None):
    """Dinero cobrado por método de pago: ventas POS pagadas + pagos de apartados (netos) + abonos de cuentas."""
    filtro = " AND fecha >= ?" if desde else ""
    args = (desde,) if desde else ()
    out = {}

    def sumar(filas):
        for metodo, total in filas:
            clave = metodo or "Sin dato"
            out[clave] = out.get(clave, 0.0) + float(total or 0.0)

    sumar(c.execute(f"SELECT metodo_pago, SUM(total) FROM ventas WHERE estado = 'Pagada'{filtro} GROUP BY metodo_pago", args).fetchall())
    sumar(c.execute(f"SELECT metodo_pago, SUM(monto) FROM apartado_pagos WHERE 1 = 1{filtro} GROUP BY metodo_pago", args).fetchall())
    sumar(c.execute(f"SELECT metodo_pago, SUM(monto_crc) FROM abonos WHERE 1 = 1{filtro} GROUP BY metodo_pago", args).fetchall())
    sumar([("Sin dato", c.execute(f"SELECT COALESCE(SUM(total), 0) FROM ventas_rapidas WHERE id_venta IS NULL{filtro}", args).fetchone()[0])])
    return {k: v for k, v in out.items() if abs(v) > 0.001}


def _guardar_doc_pos(doc):
    st.session_state["pos_doc"] = preparar_documento(doc)


def _opciones_clientes():
    df = pd.read_sql("SELECT id_cliente, id_cliente || ' — ' || nombre AS display FROM clientes ORDER BY id_cliente", conn)
    return df, dict(zip(df["display"], df["id_cliente"]))


def ui_pos_nueva_venta():
    carrito = st.session_state.setdefault("pos_cart", [])
    n = st.session_state.get("pos_n", 0)

    with card():
        st.markdown("<h4 style='color:#be185d;'>🛒 Agregar productos</h4>", unsafe_allow_html=True)
        origen = st.radio("Tipo de producto:", ["Del Stock General", "Producto Express (No Registrado)"], horizontal=True, key="pos_origen")
        if origen == "Del Stock General":
            df = pd.read_sql("SELECT id, descripcion, precio, cantidad FROM productos WHERE (id_cliente IS NULL OR id_cliente = '') AND cantidad > 0", conn)
            en_carrito = {}
            for it in carrito:
                if it.get("id_prod"):
                    en_carrito[it["id_prod"]] = en_carrito.get(it["id_prod"], 0) + it["cant"]
            df["disp"] = [int(r["cantidad"]) - en_carrito.get(int(r["id"]), 0) for _, r in df.iterrows()]
            df = df[df["disp"] > 0]
            if df.empty:
                st.warning("⚠️ No hay productos disponibles en el Stock General (o ya están todos en el carrito).")
            else:
                opciones = {f"{r['descripcion']} - ₡{(r['precio'] or 0.0):,.0f} (Disp.: {int(r['disp'])})": r for _, r in df.iterrows()}
                sel = st.selectbox("Producto del stock:", list(opciones.keys()), key="pos_stock_sel")
                r = opciones[sel]
                cant = st.number_input("Cantidad:", min_value=1, max_value=int(r["disp"]), value=1, step=1, key="pos_cant_num")
                if st.button("➕ Agregar al carrito", key="pos_add_stock"):
                    existente = next((i for i in carrito if i.get("id_prod") == int(r["id"])), None)
                    if existente:
                        existente["cant"] += int(cant)
                    else:
                        carrito.append({"id_prod": int(r["id"]), "desc": r["descripcion"], "cant": int(cant),
                                        "precio": float(r["precio"] or 0.0), "origen": "Del Stock General"})
                    st.rerun()
        else:
            desc = st.text_input("Descripción del producto:", value="Venta Express", key="pos_expr_desc")
            precio = st.number_input("Precio unitario (₡):", min_value=0.0, value=1000.0, step=500.0, key="pos_expr_precio")
            cant = st.number_input("Cantidad:", min_value=1, value=1, step=1, key="pos_expr_cant")
            if st.button("➕ Agregar al carrito", key="pos_add_expr"):
                if not desc.strip():
                    st.error("Escribe la descripción del producto.")
                elif precio <= 0:
                    st.error("El precio debe ser mayor a 0.")
                else:
                    carrito.append({"id_prod": None, "desc": desc.strip(), "cant": int(cant), "precio": float(precio),
                                    "origen": "Producto Express (No Registrado)"})
                    st.rerun()

    with card():
        st.markdown("<h4 style='color:#be185d;'>🧺 Carrito</h4>", unsafe_allow_html=True)
        if not carrito:
            st.info("El carrito está vacío. Agrega productos arriba.")
        for i, it in enumerate(carrito):
            c1, c2, c3 = st.columns([6, 3, 1])
            with c1:
                st.markdown(f"**{it['desc']}**  \n{it['cant']} × ₡{it['precio']:,.0f}")
            with c2:
                st.markdown(f"**₡{it['cant'] * it['precio']:,.0f}**")
            with c3:
                quitar = st.button("🗑️", key=f"pos_del_{i}")
            if quitar:
                carrito.pop(i)
                st.rerun()
        subtotal = _total_carrito(carrito)
        if carrito:
            st.markdown(f"Subtotal: **₡{subtotal:,.0f}**")
            if st.button("Vaciar carrito", key="pos_vaciar"):
                st.session_state["pos_cart"] = []
                st.rerun()

    if not carrito:
        return

    with card():
        st.markdown("<h4 style='color:#be185d;'>💳 Cobro</h4>", unsafe_allow_html=True)
        c_d1, c_d2 = st.columns(2)
        with c_d1:
            tipo_desc = st.selectbox("Tipo de descuento:", ["Sin Descuento", "Porcentaje (%)", "Monto Fijo (₡)"], key=f"pos_tipo_desc_{n}")
        with c_d2:
            val_desc = st.number_input("Valor del descuento:", min_value=0.0, value=0.0, key=f"pos_val_desc_{n}")
        monto_desc = subtotal * (val_desc / 100.0) if tipo_desc == "Porcentaje (%)" else (val_desc if tipo_desc == "Monto Fijo (₡)" else 0.0)
        monto_desc = min(monto_desc, subtotal)
        total = max(0.0, subtotal - monto_desc)

        df_cli, mapa_cli = _opciones_clientes()
        sel_cli = st.selectbox("Clienta (opcional; obligatoria para apartados):", ["Venta general / anónima"] + list(df_cli["display"]), key=f"pos_cli_sel_{n}")
        id_cli = mapa_cli.get(sel_cli)
        nombre_libre = None
        if not id_cli:
            nombre_libre = st.text_input("Nombre en el comprobante (opcional):", key=f"pos_nombre_libre_{n}")

        st.markdown(f"### Total a cobrar: ₡{total:,.0f}")
        opc = st.radio("Método de pago:", _OPC_PAGO_APARTADO, horizontal=True, key=f"pos_metodo_{n}")
        metodo = _metodo_de(opc)

        if metodo != "Apartado":
            recibido, referencia = total, ""
            if metodo == "Efectivo":
                recibido = st.number_input("Monto recibido (₡):", min_value=0.0, value=float(total), step=1000.0, key=f"pos_recibido_{n}_{int(total)}")
                if recibido + 0.001 >= total:
                    st.success(f"💵 Vuelto a entregar: ₡{recibido - total:,.0f}")
                else:
                    st.warning(f"Falta ₡{total - recibido:,.0f} para completar el pago.")
            else:
                referencia = st.text_input(_hint_referencia(metodo), key=f"pos_ref_{n}_{metodo}")
            if st.button("⚡ Cobrar venta", key="btn_pos_cobrar"):
                id_venta, err = crear_venta(carrito, monto_desc, metodo, recibido, referencia, id_cli, nombre_libre)
                if err:
                    st.error(err)
                else:
                    _guardar_doc_pos(construir_doc_venta(id_venta))
                    folio = c.execute("SELECT folio FROM ventas WHERE id = ?", (id_venta,)).fetchone()[0]
                    st.session_state["pos_cart"] = []
                    st.session_state["pos_n"] = n + 1
                    flash("success", f"✅ Venta {folio} cobrada ({metodo}). Abajo está su comprobante.")
                    st.rerun()
        else:
            pol = politica_apartado()
            prima_min = _arriba(total * pol["prima_pct"] / 100.0)
            st.info(f"🕒 **Apartado**: los productos quedan reservados hasta completar el pago. Prima mínima: **₡{prima_min:,.0f}** ({pol['prima_pct']:g}% del total). Sin intereses.")
            dias_opc = sorted({15, 30, 45, 60, pol["plazo_dias"]})
            dias = st.selectbox("Plazo para completar el pago:", dias_opc, index=dias_opc.index(pol["plazo_dias"]),
                                format_func=lambda d: f"{d} días", key=f"pos_ap_dias_{n}")
            prima = st.number_input("Prima a pagar hoy (₡):", min_value=0.0, max_value=float(max(total, 1.0)), value=float(min(prima_min, total)), step=1000.0,
                                    key=f"pos_ap_prima_{n}_{int(total)}")
            opc_p = st.radio("Método de pago de la prima:", _OPC_PAGO, horizontal=True, key=f"pos_ap_metodo_{n}")
            metodo_p = _metodo_de(opc_p)
            ref_p = ""
            if metodo_p != "Efectivo":
                ref_p = st.text_input(_hint_referencia(metodo_p), key=f"pos_ap_ref_{n}_{metodo_p}")
            notas = st.text_input("Notas del apartado (opcional):", key=f"pos_ap_notas_{n}")
            if st.button("🕒 Crear apartado", key="btn_pos_apartar"):
                if monto_desc > 0:
                    st.error("Los descuentos no aplican a apartados. Quita el descuento o cobra la venta completa.")
                else:
                    id_ap, err = crear_apartado(carrito, id_cli, dias, prima, metodo_p, ref_p, notas)
                    if err:
                        st.error(err)
                    else:
                        _guardar_doc_pos(construir_doc_apartado(id_ap, "apertura"))
                        folio = c.execute("SELECT folio FROM apartados WHERE id = ?", (id_ap,)).fetchone()[0]
                        st.session_state["pos_cart"] = []
                        st.session_state["pos_n"] = n + 1
                        flash("success", f"🕒 Apartado {folio} creado: los productos quedaron reservados.")
                        st.rerun()


def ui_pos_apartados():
    pol = politica_apartado()
    with st.expander("⚙️ Política de apartados"):
        st.caption("Estas reglas se aplican a los apartados nuevos y se imprimen en su comprobante.")
        p1, p2, p3 = st.columns(3)
        with p1:
            n_prima = st.number_input("Prima mínima (%)", min_value=0.0, max_value=100.0, value=float(pol["prima_pct"]), step=5.0, key="pol_prima")
        with p2:
            n_plazo = st.number_input("Plazo por defecto (días)", min_value=1, max_value=365, value=int(pol["plazo_dias"]), step=5, key="pol_plazo")
        with p3:
            n_canc = st.number_input("Cargo por cancelar (% del total)", min_value=0.0, max_value=100.0, value=float(pol["cancel_pct"]), step=5.0, key="pol_cancel")
        if st.button("💾 Guardar política", key="btn_pol_apartado"):
            guardar_politica_apartado(n_prima, n_plazo, n_canc)
            flash("success", "Política de apartados guardada.")
            st.rerun()

    filtro = st.radio("Ver:", ["Activos", "Vencidos", "Entregados", "Cancelados", "Todos"], horizontal=True, key="pos_ap_filtro")
    ids = [r[0] for r in c.execute("SELECT id FROM apartados ORDER BY id DESC").fetchall()]
    infos = [apartado_info(i) for i in ids]

    def cumple(inf):
        return {"Activos": inf["estado"] == "Activo" and not inf["vencido"], "Vencidos": inf["vencido"],
                "Entregados": inf["estado"] == "Entregado", "Cancelados": inf["estado"] == "Cancelado", "Todos": True}[filtro]

    infos = [i for i in infos if cumple(i)]
    n_venc = sum(1 for i in [apartado_info(x) for x in ids] if i["vencido"])
    if n_venc and filtro != "Vencidos":
        st.warning(f"⏰ Hay {n_venc} apartado(s) vencido(s). Revisa la pestaña «Vencidos».")

    with card():
        st.markdown("<h4 style='color:#be185d;'>🕒 Apartados</h4>", unsafe_allow_html=True)
        if not infos:
            st.info("No hay apartados en esta vista. Para crear uno, ve a «Nueva venta» y elige el método «Apartado».")
            return
        nombres = {}
        filas = []
        for i in infos:
            cli = _info_cliente(i["id_cliente"])
            nom = cli["nombre"] if cli else i["id_cliente"]
            nombres[i["id"]] = nom
            filas.append({"Folio": i["folio"], "Clienta": nom, "Creado": _fecha_corta(i["fecha"]), "Límite": _fecha_corta(i["limite"]),
                          "Total (₡)": i["total"], "Pagado (₡)": i["neto"], "Saldo (₡)": i["saldo"],
                          "Estado": "Vencido" if i["vencido"] else i["estado"]})
        st.dataframe(pd.DataFrame(filas))
        etiquetas = {f"{i['folio']} — {nombres[i['id']]} — saldo ₡{i['saldo']:,.0f} ({'Vencido' if i['vencido'] else i['estado']})": i["id"] for i in infos}
        sel = st.selectbox("Selecciona un apartado:", list(etiquetas.keys()), key=f"pos_ap_sel_{filtro}")
        id_ap = etiquetas[sel]

    inf = apartado_info(id_ap)
    with card():
        st.markdown(f"<h4 style='color:#be185d;'>{inf['folio']} · {nombres[id_ap]}</h4>", unsafe_allow_html=True)
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total", f"₡{inf['total']:,.0f}")
        m2.metric("Pagado", f"₡{inf['neto']:,.0f}")
        m3.metric("Saldo", f"₡{inf['saldo']:,.0f}")
        m4.metric("Límite", _fecha_corta(inf["limite"]))
        if inf["estado"] == "Activo":
            st.progress(min(1.0, (inf["neto"] / inf["total"]) if inf["total"] else 1.0))
            if inf["vencido"]:
                st.error("⏰ Este apartado está vencido. Puedes ampliar el plazo o cancelarlo.")
        else:
            st.info(f"Estado: {inf['estado']}" + (f" el {_fecha_corta(inf['cierre'])}" if inf["cierre"] else ""))
        st.markdown("**Productos reservados:**")
        for _i, desc, cant, pu in inf["items"]:
            st.markdown(f"- {desc} — {cant} × ₡{pu:,.0f} = **₡{cant * pu:,.0f}**")
        st.markdown("**Pagos y movimientos:**")
        for pid, fecha, monto, metodo, ref, tipo in inf["pagos"]:
            st.markdown(f"- {_fecha_corta(fecha)} · {tipo} · {metodo}" + (f" (ref. {ref})" if ref else "") + f" — **₡{monto:,.0f}**")

    if inf["estado"] == "Activo":
        with card("sub"):
            st.markdown("**➕ Registrar abono**")
            a1, a2 = st.columns(2)
            with a1:
                monto = st.number_input("Monto (₡):", min_value=0.0, max_value=float(max(inf["saldo"], 1.0)), value=float(min(inf["saldo"], 10000.0)),
                                        step=1000.0, key=f"ap_monto_{id_ap}")
            with a2:
                metodo = _metodo_de(st.radio("Método de pago:", _OPC_PAGO, horizontal=True, key=f"ap_metodo_{id_ap}"))
            ref = ""
            if metodo != "Efectivo":
                ref = st.text_input(_hint_referencia(metodo), key=f"ap_ref_{id_ap}_{metodo}")
            if st.button("💾 Guardar abono", key=f"btn_ap_abono_{id_ap}"):
                id_pago, err = registrar_pago_apartado(id_ap, monto, metodo, ref)
                if err:
                    st.error(err)
                else:
                    _guardar_doc_pos(construir_doc_apartado(id_ap, "abono", id_pago))
                    flash("success", f"✅ Abono de ₡{monto:,.0f} registrado en {inf['folio']}.")
                    st.rerun()

        with card("sub"):
            b1, b2 = st.columns(2)
            with b1:
                st.markdown("**📦 Entregar al cliente**")
                st.caption("Se habilita cuando el saldo llega a ₡0.")
                if st.button("✅ Entregar apartado", key=f"btn_ap_entregar_{id_ap}", disabled=inf["saldo"] > 0.5):
                    ok, txt = entregar_apartado(id_ap)
                    if ok:
                        _guardar_doc_pos(construir_doc_apartado(id_ap, "entrega"))
                        flash("success", f"🎉 {inf['folio']} entregado.")
                        st.rerun()
                    else:
                        st.error(txt)
            with b2:
                st.markdown("**📅 Ampliar plazo**")
                hoy = datetime.now().date()
                try:
                    actual = datetime.strptime(inf["limite"], "%Y-%m-%d").date()
                except Exception:
                    actual = hoy
                nueva = st.date_input("Nueva fecha límite:", value=max(actual, hoy), min_value=hoy, key=f"ap_fecha_{id_ap}")
                if st.button("📅 Guardar nueva fecha", key=f"btn_ap_ampliar_{id_ap}"):
                    ok, txt = ampliar_plazo_apartado(id_ap, datetime.combine(nueva, datetime.min.time()))
                    flash("success" if ok else "warning", txt)
                    st.rerun()

        with card("sub"):
            st.markdown("**❌ Cancelar apartado**")
            sugerido = cargo_cancelacion_sugerido(inf)
            st.caption(f"Se devuelven los productos al inventario. Cargo según la política ({inf['cancel_pct']:g}% del total): ₡{sugerido:,.0f}. Puedes cambiarlo (por ejemplo, dejarlo en 0).")
            cargo = st.number_input("Cargo por cancelación (₡):", min_value=0.0, max_value=float(max(inf["neto"], 1.0)), value=float(min(sugerido, max(inf["neto"], 1.0))), step=500.0, key=f"ap_cargo_{id_ap}")
            reembolso = max(0.0, inf["neto"] - cargo)
            st.markdown(f"Pagado: ₡{inf['neto']:,.0f} · Cargo: ₡{cargo:,.0f} · **Se reembolsa: ₡{reembolso:,.0f}**")
            metodo_r = "Efectivo"
            if reembolso > 0:
                metodo_r = _metodo_de(st.radio("¿Cómo se devuelve el dinero?", _OPC_PAGO, horizontal=True, key=f"ap_metodo_r_{id_ap}"))
            conf = st.checkbox("Confirmo que quiero cancelar este apartado", key=f"ap_conf_{id_ap}")
            if st.button("❌ Cancelar apartado", key=f"btn_ap_cancelar_{id_ap}", disabled=not conf):
                ok, txt = cancelar_apartado(id_ap, cargo, metodo_r)
                if ok:
                    _guardar_doc_pos(construir_doc_apartado(id_ap, "cancelacion"))
                    flash("success", txt)
                    st.rerun()
                else:
                    st.error(txt)

    modo = {"Entregado": "entrega", "Cancelado": "cancelacion"}.get(inf["estado"], "estado")
    if st.button("🧾 Ver / reimprimir comprobante", key=f"btn_ap_doc_{id_ap}"):
        _guardar_doc_pos(construir_doc_apartado(id_ap, modo))
        st.rerun()


def ui_pos_historial():
    periodo = st.radio("Período:", ["Hoy", "Últimos 7 días", "Este mes", "Todo"], horizontal=True, key="pos_hist_periodo")
    hoy = datetime.now()
    desde = {"Hoy": hoy.strftime("%Y-%m-%d"), "Últimos 7 días": (hoy - timedelta(days=6)).strftime("%Y-%m-%d"),
             "Este mes": hoy.strftime("%Y-%m-01"), "Todo": None}[periodo]

    with card():
        st.markdown("<h4 style='color:#be185d;'>💰 Cobrado por método de pago</h4>", unsafe_allow_html=True)
        st.caption("Incluye ventas del POS, pagos de apartados y abonos de cuentas del período.")
        cobros = cobros_por_metodo(desde)
        if cobros:
            cols = st.columns(len(cobros) + 1)
            for col, (m, v) in zip(cols, sorted(cobros.items())):
                col.metric(f"{ICONO_PAGO.get(m, '•')} {m}", f"₡{v:,.0f}")
            cols[-1].metric("Total", f"₡{sum(cobros.values()):,.0f}")
        else:
            st.info("No hay cobros en este período.")

    with card():
        st.markdown("<h4 style='color:#be185d;'>🧾 Ventas del POS</h4>", unsafe_allow_html=True)
        filtro = " WHERE v.fecha >= ?" if desde else ""
        df = pd.read_sql(f"""SELECT v.id, v.folio AS Folio, v.fecha AS Fecha, COALESCE(c.nombre, v.cliente_nombre, 'Venta general') AS Cliente,
                             v.total AS 'Total (₡)', v.metodo_pago AS Método, v.estado AS Estado
                             FROM ventas v LEFT JOIN clientes c ON c.id_cliente = v.id_cliente{filtro} ORDER BY v.id DESC LIMIT 200""",
                         conn, params=(desde,) if desde else None)
        if df.empty:
            st.info("No hay ventas en este período.")
            return
        st.dataframe(df.drop(columns=["id"]))
        etiquetas = {f"{r['Folio']} — {r['Cliente']} — ₡{r['Total (₡)']:,.0f} ({r['Método']}, {r['Estado']})": int(r["id"]) for _, r in df.iterrows()}
        sel = st.selectbox("Selecciona una venta:", list(etiquetas.keys()), key="pos_hist_sel")
        id_v = etiquetas[sel]
        h1, h2 = st.columns(2)
        with h1:
            if st.button("🧾 Ver / reimprimir comprobante", key=f"btn_hist_doc_{id_v}"):
                _guardar_doc_pos(construir_doc_venta(id_v))
                st.rerun()
        with h2:
            estado = c.execute("SELECT estado FROM ventas WHERE id = ?", (id_v,)).fetchone()[0]
            conf = st.checkbox("Confirmo anular esta venta (devuelve el stock)", key=f"hist_conf_{id_v}", disabled=estado != "Pagada")
            if st.button("🗑️ Anular venta", key=f"btn_hist_anular_{id_v}", disabled=(estado != "Pagada" or not conf)):
                ok, txt = anular_venta(id_v)
                flash("success" if ok else "error", txt)
                st.rerun()


def ui_pos():
    st.markdown(
        """
        <div class="top-banner">
            <div style="font-size: 18px; font-weight: 700;">Punto de Venta (POS)</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    doc_pos = st.session_state.get("pos_doc")
    if doc_pos:
        with card():
            mostrar_documento(doc_pos, "pos")
            if st.button("✖️ Cerrar comprobante", key="pos_doc_cerrar"):
                st.session_state.pop("pos_doc", None)
                st.rerun()
    sub = st.radio("Sección del POS", ["🛒 Nueva venta", "🕒 Apartados", "🧾 Historial de ventas"], horizontal=True,
                   label_visibility="collapsed", key="pos_sub")
    if sub == "🛒 Nueva venta":
        ui_pos_nueva_venta()
    elif sub == "🕒 Apartados":
        ui_pos_apartados()
    else:
        ui_pos_historial()


# -------------------------------------------------------------
# 3B. REGISTRO PÚBLICO POR LINK (las clientas se registran solas)
# -------------------------------------------------------------
def cfg_get(clave, default=None):
    r = c.execute("SELECT valor FROM config WHERE clave = ?", (clave,)).fetchone()
    return r[0] if r and r[0] is not None else default


def cfg_set(clave, valor):
    c.execute("INSERT OR REPLACE INTO config (clave, valor) VALUES (?, ?)", (clave, str(valor)))
    conn.commit()


def _solo_digitos(texto):
    return "".join(ch for ch in str(texto or "") if ch.isdigit())


def _normalizar_telefono(texto):
    d = _solo_digitos(texto)
    if len(d) == 11 and d.startswith("506"):
        d = d[3:]
    return d


def siguiente_id_cliente(prefijo):
    filas = c.execute("SELECT id_cliente FROM clientes").fetchall()
    nums = []
    for (cod,) in filas:
        cod = str(cod or "")
        if cod.startswith(prefijo + "-") and cod[len(prefijo) + 1:].strip().isdigit():
            nums.append(int(cod[len(prefijo) + 1:]))
    return f"{prefijo}-{(max(nums) + 1 if nums else 1):04d}"


def _token_registro():
    """Código secreto del link. Se crea la primera vez que se necesita."""
    t = cfg_get("registro_token")
    if not t:
        t = secrets.token_urlsafe(9)
        cfg_set("registro_token", t)
    return t


def _url_base_app():
    guardada = cfg_get("url_base")
    if guardada:
        return str(guardada).strip().rstrip("/")
    try:
        h = st.context.headers
        host = h.get("Host") or h.get("host")
        if host:
            proto = h.get("X-Forwarded-Proto") or ("http" if str(host).startswith(("localhost", "127.")) else "https")
            return f"{proto}://{host}"
    except Exception:
        pass
    return ""


def _link_valido(token_url):
    token = cfg_get("registro_token")
    if not token or cfg_get("registro_activo", "1") != "1":
        return False
    return hmac.compare_digest(str(token_url), str(token))


def pagina_registro_publico(token_url):
    """Pantalla pública: solo se llega con el link secreto. No muestra ningún dato de nadie."""
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        logo_path = next((n for n in ["logo.jpg", "logo.png", "logo.jpeg", "21237.jpg"] if os.path.exists(n)), None)
        if logo_path:
            st.image(logo_path)
        else:
            st.markdown("<h1 style='text-align: center; color: #be185d;'>🛍️ Minici Store</h1>", unsafe_allow_html=True)

    if not _link_valido(token_url):
        st.markdown(
            """<div class="login-hero"><h2>Este link no está disponible</h2>
            <p>Pídele a Minici Store que te envíe el link de registro actualizado.</p></div>""",
            unsafe_allow_html=True)
        return

    ok = st.session_state.get("pub_ok")
    if ok:
        st.markdown(
            f"""<div class="login-hero"><h2>🎉 ¡Listo, {html.escape(ok['nombre'])}!</h2>
            <p>Tu registro quedó guardado. Este es tu código de acceso:</p></div>""",
            unsafe_allow_html=True)
        with card():
            st.markdown(f"<h2 style='text-align:center; color:#be185d; letter-spacing:2px;'>{html.escape(ok['id'])}</h2>",
                        unsafe_allow_html=True)
            st.info("📌 Guarda este código (puedes tomarle captura). Con él entras a ver tus pedidos y abonos.")
            st.markdown('<div style="text-align:center"><a href="?" target="_self">Ingresar a mi cuenta</a></div>',
                        unsafe_allow_html=True)
            if st.button("Registrar a otra persona", key="pub_otra"):
                st.session_state["pub_n"] = st.session_state.get("pub_n", 0) + 1
                st.session_state.pop("pub_ok", None)
                st.rerun()
        return

    st.markdown(
        """<div class="login-hero"><h2>¡Regístrate en Minici Store!</h2>
        <p>Completa tus datos para crear tu cuenta de clienta.</p></div>""",
        unsafe_allow_html=True)

    n = st.session_state.get("pub_n", 0)
    kp = f"pub{n}"
    with card():
        nombre = st.text_input("Nombre completo *", placeholder="Ej. María López", key=f"{kp}_nombre")
        tel = st.text_input("Teléfono / WhatsApp *", placeholder="Ej. 8888 8888", key=f"{kp}_tel")
        correo = st.text_input("Correo electrónico (opcional)", placeholder="correo@ejemplo.com", key=f"{kp}_correo")
        st.markdown("📍 **¿Dónde te entregamos?**")
        prov, canton, distrito, senas = selector_ubicacion(kp)
        st.caption("🔒 Tus datos se usan solo para gestionar tus pedidos y envíos.")

        if st.button("✅ Crear mi registro", key=f"{kp}_enviar"):
            if st.session_state.get("pub_conteo", 0) >= 5:
                st.error("Se alcanzó el máximo de registros desde este dispositivo. Escríbenos por WhatsApp.")
            else:
                faltan = []
                if len(nombre.strip()) < 3:
                    faltan.append("tu nombre")
                tel_n = _normalizar_telefono(tel)
                if not (8 <= len(tel_n) <= 15):
                    faltan.append("un teléfono válido (8 dígitos)")
                if not prov:
                    faltan.append("la provincia")
                if not canton:
                    faltan.append("el cantón")
                if not distrito:
                    faltan.append("el distrito")
                if faltan:
                    st.error("Falta completar: " + ", ".join(faltan) + ".")
                elif any(_normalizar_telefono(t)[-8:] == tel_n[-8:] for (t,) in c.execute("SELECT telefono FROM clientes").fetchall() if t):
                    st.warning("Este teléfono ya está registrado. Si necesitas actualizar tus datos, escríbenos por WhatsApp.")
                else:
                    nuevo_id = None
                    for _ in range(3):
                        cand = siguiente_id_cliente("MIN")
                        try:
                            c.execute("INSERT INTO clientes (id_cliente, nombre, telefono, correo, provincia, canton, distrito, senas, metodo_envio, origen, fecha_registro) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                      (cand, nombre.strip().title(), tel_n, correo.strip(), prov, canton, distrito, senas, None,
                                       "link", datetime.now().strftime("%Y-%m-%d %H:%M")))
                            conn.commit()
                            nuevo_id = cand
                            break
                        except sqlite3.IntegrityError:
                            conn.rollback()
                    if nuevo_id:
                        st.session_state["pub_conteo"] = st.session_state.get("pub_conteo", 0) + 1
                        st.session_state["pub_ok"] = {"nombre": nombre.strip().title(), "id": nuevo_id}
                        st.rerun()
                    else:
                        st.error("No se pudo completar el registro. Intenta de nuevo en unos segundos.")


def panel_link_registro():
    """Cuadro del administrador: copiar, compartir, desactivar o cambiar el link de registro."""
    with st.expander("🔗 Link para que las clientas se registren solas"):
        activo = cfg_get("registro_activo", "1") == "1"
        base = _url_base_app()
        token = _token_registro()
        if base:
            link = f"{base}/?registro={token}"
            st.markdown("**Pásales este link** (toca el ícono de la derecha para copiarlo):")
            st.code(link, language=None)
            msg = f"¡Hola! 💖 Regístrate en Minici Store con este link: {link}"
            st.markdown(f'<a href="https://wa.me/?text={urllib.parse.quote(msg)}" target="_blank" class="btn-whatsapp">📲 Compartir por WhatsApp</a>',
                        unsafe_allow_html=True)
        else:
            st.warning("No pude detectar la dirección de tu app. Escríbela abajo (ej. https://minici.streamlit.app) y guarda.")
        st.caption("Solo quien tenga este link puede registrarse. Las clientas nuevas aparecen en tu directorio con su dirección.")

        nuevo_activo = st.checkbox("Link activo (las clientas pueden registrarse)", value=activo, key="chk_reg_activo")
        if nuevo_activo != activo:
            cfg_set("registro_activo", "1" if nuevo_activo else "0")
            st.rerun()

        url_in = st.text_input("Dirección de tu app", value=base, placeholder="https://tu-app.streamlit.app", key="inp_url_base")
        if st.button("💾 Guardar dirección", key="btn_url_base") and url_in.strip():
            cfg_set("url_base", url_in.strip().rstrip("/"))
            flash("success", "Dirección guardada.")
            st.rerun()

        st.divider()
        conf = st.checkbox("Entiendo que el link anterior dejará de funcionar", key="chk_reg_nuevo")
        if st.button("🔄 Crear un link nuevo", key="btn_reg_nuevo", disabled=not conf):
            cfg_set("registro_token", secrets.token_urlsafe(9))
            flash("success", "Link nuevo creado. El anterior ya no funciona.")
            st.rerun()

        recientes = c.execute("SELECT id_cliente, nombre, telefono, fecha_registro FROM clientes WHERE origen = 'link' ORDER BY fecha_registro DESC LIMIT 5").fetchall()
        total = c.execute("SELECT COUNT(*) FROM clientes WHERE origen = 'link'").fetchone()[0]
        if total:
            st.markdown(f"**Registradas por link: {total}** (últimas 5)")
            for cod, nom, tl, fr in recientes:
                st.markdown(f"- `{cod}` · {nom} · {tl} · {fr}")


_tok_url = st.query_params.get("registro")
if _tok_url:
    pagina_registro_publico(str(_tok_url))
    st.stop()


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

    st.caption(estado_respaldo_auto()[1])
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
                st.info(f"💾 Respaldo creado ({_rsp['hora']}). Descárgalo abajo.")
            st.download_button("⬇️ Descargar respaldo (datos + fotos + Excel)", data=_rsp["zip"],
                               file_name=_rsp["nombre"], mime="application/zip", key="dl_respaldo")

    with st.expander("♻️ Restaurar respaldo"):
        if _gh_config():
            _cfg = _gh_config()
            st.success(f"GitHub conectado: {_cfg['repo']} (rama {_cfg['branch']}). Si el servidor se reinicia, los datos se recuperan solos.")
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
                            mostrar_ticket(id_cliente, "compra")

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
            panel_link_registro()

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

                n_reg = st.session_state.get("reg_cli_n", 0)
                kr = f"reg{n_reg}"
                nombre = st.text_input("Nombre Completo", placeholder="Ej. Maria Lopez", key=f"{kr}_nombre")
                tel = st.text_input("Teléfono / WhatsApp", placeholder="Ej. 88888888", key=f"{kr}_tel")
                correo = st.text_input("Correo Electrónico", placeholder="Ej. correo@ejemplo.com", key=f"{kr}_correo")

                st.markdown("📍 **Dirección:**")
                prov_in, canton_in, distrito_in, senas_in = selector_ubicacion(kr)

                nombre_caja_in, precio_caja_in, envio_in = "", 0.0, None
                if "EMP" in tipo_registro:
                    st.markdown("---")
                    st.markdown("🚚 **Envío de las cajas:**")
                    envio_in = selector_envio(f"{kr}_envio")
                    st.markdown("💼 **Configuración de Caja para Emprendedor:**")
                    nombre_caja_in, precio_caja_in = selector_talla_caja(f"{kr}_caja")

                if st.button("Guardar Registro", key=f"{kr}_guardar"):
                    faltan = []
                    if not nombre.strip():
                        faltan.append("el nombre")
                    if not prov_in:
                        faltan.append("la provincia")
                    if not canton_in:
                        faltan.append("el cantón")
                    if not distrito_in:
                        faltan.append("el distrito")
                    if faltan:
                        st.error("Falta completar: " + ", ".join(faltan) + ".")
                    else:
                        nombre_clean = nombre.strip().title()
                        c.execute("INSERT OR REPLACE INTO clientes (id_cliente, nombre, telefono, correo, provincia, canton, distrito, senas, metodo_envio) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                  (nuevo_id, nombre_clean, tel.strip(), correo.strip(), prov_in, canton_in, distrito_in, senas_in, envio_in))
                        if "EMP" in tipo_registro:
                            caja_clean = nombre_caja_in or "Caja Asignada"
                            c.execute("INSERT OR REPLACE INTO cajas_emprendedores (id_cliente, nombre_caja, precio_caja) VALUES (?, ?, ?)", (nuevo_id, caja_clean, precio_caja_in))
                        conn.commit()
                        st.session_state["reg_cli_n"] = n_reg + 1
                        flash("success", f"¡Registro de {nombre_clean} guardado exitosamente ({nuevo_id})!")
                        st.rerun()

            with card():
                st.markdown("<h4 style='color:#be185d;'>📋 Directorio de Clientes y Emprendedores</h4>", unsafe_allow_html=True)
            
                clientas_todas = pd.read_sql("""
                    SELECT c.id_cliente as Código, c.nombre as Nombre, c.telefono as Teléfono,
                           c.provincia as _prov, c.canton as _cant, c.distrito as _dist,
                           COALESCE(c.metodo_envio, '—') as 'Envío',
                           COALESCE(e.nombre_caja, 'N/A') as 'Caja', COALESCE(e.precio_caja, 0.0) as 'Precio Caja (₡)'
                    FROM clientes c 
                    LEFT JOIN cajas_emprendedores e ON c.id_cliente = e.id_cliente 
                    ORDER BY c.id_cliente DESC""", conn)
                clientas_todas.insert(3, "Ubicación", [texto_ubicacion(a_, b_, c_) or "— sin dirección —"
                                                       for a_, b_, c_ in zip(clientas_todas["_prov"], clientas_todas["_cant"], clientas_todas["_dist"])])
                clientas_todas = clientas_todas.drop(columns=["_prov", "_cant", "_dist"])
            
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
                        c.execute("SELECT nombre, telefono, correo, provincia, canton, distrito, senas, metodo_envio FROM clientes WHERE id_cliente = ?", (id_cli_exp,))
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

                            st.write("")
                            with card("sub"):
                                st.markdown("📍 **Dirección y envío:**")
                                _dir_actual = {"provincia": info_cli[3], "canton": info_cli[4], "distrito": info_cli[5], "senas": info_cli[6]}
                                ed_prov, ed_canton, ed_distrito, ed_senas = selector_ubicacion(f"exp_dir_{id_cli_exp}", _dir_actual)
                                ed_envio = None
                                if id_cli_exp.startswith("EMP-"):
                                    ed_envio = selector_envio(f"exp_env_{id_cli_exp}", info_cli[7])
                                if st.button("💾 Guardar dirección y envío", key=f"btn_dir_{id_cli_exp}"):
                                    if not (ed_prov and ed_canton and ed_distrito):
                                        st.error("Completa provincia, cantón y distrito.")
                                    else:
                                        c.execute("UPDATE clientes SET provincia=?, canton=?, distrito=?, senas=?, metodo_envio=? WHERE id_cliente=?",
                                                  (ed_prov, ed_canton, ed_distrito, ed_senas, ed_envio, id_cli_exp))
                                        conn.commit()
                                        flash("success", "¡Dirección y envío actualizados!")
                                        st.rerun()

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
                            mostrar_ticket(id_cli_exp, "exp", con_checkbox=True)

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
                    metodo_ab = _metodo_de(st.radio("Método de pago:", _OPC_PAGO, horizontal=True, key="abono_metodo"))
                    ref_ab = ""
                    if metodo_ab != "Efectivo":
                        ref_ab = st.text_input(_hint_referencia(metodo_ab), key=f"abono_ref_{metodo_ab}")
                    st.write("")
                    guardar_ab = st.button("Guardar Abono", key="btn_save_abono")
                    if guardar_ab and monto <= 0:
                        st.error("El monto del abono debe ser mayor a 0.")
                    if guardar_ab and monto > 0:
                        c.execute("INSERT INTO abonos (id_cliente, monto_crc, fecha, metodo_pago, referencia) VALUES (?, ?, ?, ?, ?)",
                                  (id_c, monto, datetime.now().strftime("%Y-%m-%d %H:%M"), metodo_ab, ref_ab.strip() or None))
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
                        mostrar_ticket(id_c, "abono")

            with card():
                st.markdown("<h4 style='color:#be185d;'>📋 Historial de Abonos Recibidos</h4>", unsafe_allow_html=True)
                abonos_detalle = pd.read_sql("SELECT a.fecha as Fecha, c.id_cliente as Código, c.nombre as Cliente, a.monto_crc as 'Monto Abonado (₡)', COALESCE(a.metodo_pago, '—') as Método FROM abonos a JOIN clientes c ON a.id_cliente = c.id_cliente ORDER BY a.id DESC", conn)
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
            df_vr = pd.read_sql("SELECT COALESCE(SUM(total), 0) as total FROM ventas_rapidas WHERE id_venta IS NULL OR id_venta IN (SELECT id FROM ventas WHERE estado = 'Pagada')", conn)
            ap = resumen_apartados()

            total_ventas = (df_pedidos.iloc[0]["total"] or 0.0) + (df_cajas.iloc[0]["total"] or 0.0) + (df_vr.iloc[0]["total"] or 0.0) + ap["total_vigente"] + ap["cargos_retenidos"]
            total_ingresos_reales = (df_abonos.iloc[0]["total"] or 0.0) + (df_vr.iloc[0]["total"] or 0.0) + ap["cobrado_neto"]
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
                st.markdown("<h4 style='color:#be185d;'>🕒 Apartados y 💳 cobros por método de pago</h4>", unsafe_allow_html=True)
                ca1, ca2, ca3 = st.columns(3)
                ca1.metric("Apartados activos", f"{ap['activos']}")
                ca2.metric("Por cobrar en apartados", f"₡{ap['por_cobrar']:,.0f}")
                ca3.metric("Dinero recibido en apartados activos", f"₡{ap['dinero_retenido']:,.0f}")
                cobros_all = cobros_por_metodo()
                if cobros_all:
                    cols_m = st.columns(len(cobros_all))
                    for col_m, (met, val) in zip(cols_m, sorted(cobros_all.items())):
                        col_m.metric(f"{ICONO_PAGO.get(met, '•')} {met}", f"₡{val:,.0f}")

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
                    ("Ventas_Rapidas_POS", pd.read_sql("SELECT vr.fecha as Fecha, COALESCE(v.folio, '—') as Folio, vr.producto as Producto, vr.tipo as Tipo, vr.cantidad as Cantidad, vr.precio_u as Precio_Unitario, vr.descuento as Descuento, vr.total as Total_CRC, COALESCE(v.metodo_pago, 'Sin dato') as Método FROM ventas_rapidas vr LEFT JOIN ventas v ON v.id = vr.id_venta WHERE vr.id_venta IS NULL OR v.estado = 'Pagada'", conn)),
                    ("Apartados", pd.read_sql("SELECT a.folio as Folio, a.fecha as Fecha, a.fecha_limite as Fecha_Límite, COALESCE(c.nombre, a.id_cliente) as Cliente, a.total as Total_CRC, a.estado as Estado FROM apartados a LEFT JOIN clientes c ON c.id_cliente = a.id_cliente", conn)),
                    ("Pagos_Apartados", pd.read_sql("SELECT a.folio as Folio, p.fecha as Fecha, p.tipo as Tipo, p.metodo_pago as Método, p.referencia as Referencia, p.monto as Monto_CRC FROM apartado_pagos p JOIN apartados a ON a.id = p.id_apartado", conn)),
                    ("Ingresos_Abonos", pd.read_sql("SELECT id_cliente as Cliente, monto_crc as Monto_CRC, fecha as Fecha, COALESCE(metodo_pago, 'Sin dato') as Método FROM abonos", conn)),
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
            ui_pos()

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