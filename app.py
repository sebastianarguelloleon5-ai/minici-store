import os
os.environ["OPENBLAS_NUM_THREADS"] = "1"

import sqlite3
from datetime import datetime
import io
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import streamlit as st
import urllib.parse

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
# 2. BASE DE DATOS SQLITE Y MIGRACIONES AUTOMÁTICAS
# -------------------------------------------------------------
conn = sqlite3.connect("minici_store.db", check_same_thread=False)
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
# FUNCIÓN PARA GENERAR IMAGEN DE FACTURA Y ENLACE WHATSAPP
# -------------------------------------------------------------
def generar_factura_imagen(id_cliente):
    c.execute("SELECT nombre, telefono FROM clientes WHERE id_cliente = ?", (id_cliente,))
    res_cli = c.fetchone()
    if not res_cli:
        return None, None

    nombre_cliente, telefono = res_cli
    
    # Datos de Caja si es Emprendedor
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

    img_w, img_h = 800, 1050
    img = Image.new("RGB", (img_w, img_h), color="#FFFFFF")
    draw = ImageDraw.Draw(img)

    draw.rectangle([(0, 0), (img_w, 140)], fill="#F3B2C9")
    
    try:
        font_title = ImageFont.truetype("arial.ttf", 28)
        font_subtitle = ImageFont.truetype("arial.ttf", 16)
        font_bold = ImageFont.truetype("arial.ttf", 16)
        font_regular = ImageFont.truetype("arial.ttf", 15)
    except IOError:
        font_title = font_subtitle = font_bold = font_regular = ImageFont.load_default()

    draw.text((40, 30), "🛍️ MINICI STORE", fill="#0f172a", font=font_title)
    draw.text((40, 75), "Comprobante de Estado de Cuenta & Pedidos", fill="#0f172a", font=font_subtitle)
    
    fecha_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    draw.text((530, 40), f"Fecha: {fecha_str}", fill="#0f172a", font=font_subtitle)
    draw.text((530, 70), f"Ref Clienta: {id_cliente}", fill="#0f172a", font=font_subtitle)

    draw.rectangle([(40, 160), (760, 230)], outline="#F3B2C9", width=2, fill="#fdf2f8")
    draw.text((60, 172), f"Cliente: {nombre_cliente}", fill="#0f172a", font=font_bold)
    draw.text((500, 172), f"Tel: {telefono if telefono else 'No registrado'}", fill="#0f172a", font=font_regular)
    
    if nombre_caja:
        draw.text((60, 200), f"💼 Caja Asignada: {nombre_caja} (₡{precio_caja:,.0f})", fill="#be185d", font=font_bold)

    y = 260
    draw.text((40, y), "DETALLE DE ARTÍCULOS / PEDIDOS:", fill="#be185d", font=font_bold)
    y += 35

    draw.rectangle([(40, y), (760, y + 35)], fill="#f1f5f9")
    draw.text((50, y + 8), "Descripción", fill="#0f172a", font=font_bold)
    draw.text((430, y + 8), "Cant", fill="#0f172a", font=font_bold)
    draw.text((520, y + 8), "Estado", fill="#0f172a", font=font_bold)
    draw.text((650, y + 8), "Total", fill="#0f172a", font=font_bold)
    y += 45

    if nombre_caja and precio_caja > 0:
        draw.text((50, y), f"Alquiler/Cuota: {nombre_caja}", fill="#334155", font=font_regular)
        draw.text((440, y), "1", fill="#334155", font=font_regular)
        draw.text((520, y), "Caja Mensual", fill="#64748b", font=font_regular)
        draw.text((640, y), f"₡{precio_caja:,.0f}", fill="#0f172a", font=font_bold)
        y += 35

    for p in prods:
        desc, precio, cant, estado = p[0], (p[1] or 0.0), (p[2] if p[2] else 1), (p[3] or "")
        subtot = precio * cant
        draw.text((50, y), str(desc)[:35], fill="#334155", font=font_regular)
        draw.text((440, y), str(cant), fill="#334155", font=font_regular)
        draw.text((520, y), str(estado)[:18], fill="#64748b", font=font_regular)
        draw.text((640, y), f"₡{subtot:,.0f}", fill="#0f172a", font=font_bold)
        y += 35

    y = max(y + 40, 720)
    draw.line([(40, y), (760, y)], fill="#cbd5e1", width=2)
    y += 20

    draw.text((450, y), "Total Cargos/Compras:", fill="#64748b", font=font_regular)
    draw.text((640, y), f"₡{total_compras:,.0f}", fill="#0f172a", font=font_bold)
    y += 35

    draw.text((450, y), "Total Abonado:", fill="#64748b", font=font_regular)
    draw.text((640, y), f"₡{total_abonos:,.0f}", fill="#059669", font=font_bold)
    y += 45

    draw.rectangle([(430, y), (760, y + 55)], fill="#fdf2f8", outline="#F3B2C9", width=2)
    draw.text((450, y + 15), "SALDO PENDIENTE:", fill="#be185d", font=font_bold)
    draw.text((630, y + 12), f"₡{max(0.0, saldo_pendiente):,.0f}", fill="#be185d", font=font_title)

    draw.text((250, 970), "✨ ¡Gracias por tu preferencia en Minici Store! ✨", fill="#94a3b8", font=font_subtitle)

    filename = f"factura_{id_cliente}.jpg"
    filepath = os.path.join("facturas_generadas", filename)
    img.save(filepath, "JPEG")

    telefono_clean = "".join(filter(str.isdigit, str(telefono))) if telefono else ""
    if len(telefono_clean) == 8:
        telefono_clean = "506" + telefono_clean

    mensaje = f"🛍 *MINICI STORE - ESTADO DE CUENTA*\n\n"
    mensaje += f"¡Hola *{nombre_cliente}*! 👋\n"
    mensaje += f"Te adjuntamos tu factura digital con el resumen de tus compras y abonos.\n\n"
    mensaje += f"🔴 *Saldo Pendiente:* ₡{max(0.0, saldo_pendiente):,.0f}\n\n"
    mensaje += "¡Muchas gracias por tu preferencia! ✨"

    link = f"https://wa.me/{telefono_clean}?text={urllib.parse.quote(mensaje)}" if telefono_clean else None

    return filepath, link

# -------------------------------------------------------------
# 3. ESTILOS CSS GLOBALES
# -------------------------------------------------------------
st.markdown(
    """
<style>
   .stApp {
       background-color: #fdf2f8 !important;
       font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
       color: #0f172a !important;
   }
   
   #MainMenu, footer, header {visibility: hidden;}

   label, p, span, div, h1, h2, h3, h4, h5, h6 {
       color: #0f172a;
   }

   div[data-testid="stMetric"] {
       background-color: #ffffff !important;
       padding: 12px 16px !important;
       border-radius: 12px !important;
       border: 1px solid #F3B2C9 !important;
       box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04) !important;
   }
   div[data-testid="stMetricLabel"] > div,
   div[data-testid="stMetricLabel"] label,
   div[data-testid="stMetricLabel"] p {
       color: #9d174d !important;
       font-weight: 700 !important;
       font-size: 12px !important;
       text-transform: uppercase !important;
   }
   div[data-testid="stMetricValue"] > div {
       color: #831843 !important;
       font-weight: 800 !important;
   }

   .top-banner {
       background: linear-gradient(135deg, #F3B2C9 0%, #e3a2b9 100%);
       padding: 18px;
       border-radius: 14px;
       color: #0f172a !important;
       margin-bottom: 20px;
       box-shadow: 0 4px 12px rgba(243, 178, 201, 0.4);
   }
   .top-banner * {
       color: #0f172a !important;
   }

   .form-card {
       background: #ffffff !important;
       padding: 22px;
       border-radius: 14px;
       box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
       margin-bottom: 20px;
       border: 1px solid #F3B2C9;
   }

   .section-title {
       font-size: 13px;
       font-weight: 700;
       color: #be185d !important;
       margin-bottom: 6px;
       margin-top: 10px;
       text-transform: uppercase;
   }

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
       color: #0f172a !important;
       fill: #0f172a !important;
   }

   input, textarea, [role="option"], [role="combobox"] {
       color: #0f172a !important;
       -webkit-text-fill-color: #0f172a !important;
       background-color: #ffffff !important;
   }

   div[data-baseweb="popover"],
   div[data-baseweb="popover"] * {
       background-color: #ffffff !important;
       color: #0f172a !important;
   }

   .stNumberInput button {
       background-color: #F3B2C9 !important;
       color: #0f172a !important;
       border: none !important;
   }

   div.stButton > button:first-child, div.stDownloadButton > button:first-child {
       background-color: #F3B2C9 !important;
       color: #0f172a !important;
       border-radius: 10px;
       font-weight: 700;
       border: none;
       padding: 10px 18px;
       width: 100%;
       font-size: 15px;
   }
   div.stButton > button:first-child *, div.stDownloadButton > button:first-child * {
       color: #0f172a !important;
   }
   div.stButton > button:first-child:hover, div.stDownloadButton > button:first-child:hover {
       background-color: #e3a2b9 !important;
   }

   .btn-whatsapp {
       display: inline-block;
       width: 100%;
       background-color: #25D366 !important;
       color: #ffffff !important;
       text-align: center;
       font-weight: bold;
       padding: 10px 15px;
       border-radius: 10px;
       text-decoration: none;
       font-size: 15px;
       margin-top: 10px;
       box-shadow: 0 3px 6px rgba(0,0,0,0.1);
   }
   .btn-whatsapp:hover {
       background-color: #1da851 !important;
       color: #ffffff !important;
   }

   div[data-testid="stRadio"] > div {
       flex-direction: row !important;
       gap: 6px !important;
       flex-wrap: wrap !important;
   }
   div[data-testid="stRadio"] label {
       background-color: #ffffff !important;
       border: 2px solid #F3B2C9 !important;
       padding: 6px 12px !important;
       border-radius: 10px !important;
       font-weight: 700 !important;
       cursor: pointer !important;
   }
   div[data-testid="stRadio"] label p,
   div[data-testid="stRadio"] label span,
   div[data-testid="stRadio"] label div {
       color: #be185d !important;
       -webkit-text-fill-color: #be185d !important;
       font-size: 13px !important;
   }
   div[data-testid="stRadio"] label:has(input:checked) {
       background-color: #F3B2C9 !important;
       border-color: #F3B2C9 !important;
   }
   div[data-testid="stRadio"] label:has(input:checked) p,
   div[data-testid="stRadio"] label:has(input:checked) span,
   div[data-testid="stRadio"] label:has(input:checked) div {
       color: #0f172a !important;
       -webkit-text-fill-color: #0f172a !important;
   }
   div[data-testid="stRadio"] input[type="radio"] {
       display: none !important;
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
        <div style="text-align: center; margin-bottom: 15px;">
            <h2 style="color: #be185d; font-size: 22px;">¡Bienvenidos a Minici Store!</h2>
            <p style="color: #64748b; font-size: 14px;">Ingresa tu código de acceso para continuar.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="form-card">', unsafe_allow_html=True)
    codigo_ingresado = st.text_input("Código de acceso", placeholder="Ej. MIN-0001 o EMP-0001", key="login_input")

    st.write("")
    if st.button("🚀 Ingresar al Sistema", key="btn_login"):
        codigo_limpio = codigo_ingresado.strip().upper()
        
        if codigo_limpio == "KENDRA5412":
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
    st.markdown("</div>", unsafe_allow_html=True)

# -------------------------------------------------------------
# 5. PANEL DE ADMINISTRADOR
# -------------------------------------------------------------
elif st.session_state.user_role == "admin":
    col_a, col_b = st.columns([3, 1])
    with col_a:
        st.markdown(
            "<h3 style='color: #be185d;'>⚙️ Panel Administrador</h3>", unsafe_allow_html=True
        )
    with col_b:
        if st.button("🚪 Salir", key="btn_salir_admin"):
            st.session_state.user_role = None
            st.rerun()

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

            st.markdown('<div class="form-card">', unsafe_allow_html=True)
            
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
                        prod_clean = producto.strip().capitalize()
                        tienda_clean = tienda.strip().capitalize()
                        cat_clean = categoria.strip().capitalize()
                        obs_clean = observaciones.strip().capitalize() if observaciones else ""

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
            st.markdown("</div>", unsafe_allow_html=True)

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

            st.markdown('<div class="form-card">', unsafe_allow_html=True)
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

            with st.form("form_registro_cliente"):
                nombre = st.text_input("Nombre Completo", placeholder="Ej. Maria Lopez")
                tel = st.text_input("Teléfono / WhatsApp", placeholder="Ej. 88888888")
                correo = st.text_input("Correo Electrónico", placeholder="Ej. correo@ejemplo.com")

                nombre_caja_in, precio_caja_in = "", 0.0
                if "EMP" in tipo_registro:
                    st.markdown("---")
                    st.markdown("💼 **Configuración de Caja para Emprendedor:**")
                    nombre_caja_in = st.text_input("Identificación / Nombre de la Caja", placeholder="Ej. Caja #01 - Accesorios", value="Caja #01")
                    precio_caja_in = st.number_input("Precio o Alquiler Mensual de Caja (₡ CRC)", min_value=0.0, value=0.0, step=1000.0)

                btn_guardar_cli = st.form_submit_button("Guardar Registro")

                if btn_guardar_cli:
                    if nombre.strip():
                        nombre_clean = nombre.strip().title()
                        c.execute("INSERT OR REPLACE INTO clientes (id_cliente, nombre, telefono, correo) VALUES (?, ?, ?, ?)", (nuevo_id, nombre_clean, tel, correo))
                        
                        if "EMP" in tipo_registro:
                            caja_clean = nombre_caja_in.strip().capitalize() if nombre_caja_in else "Caja Asignada"
                            c.execute("INSERT OR REPLACE INTO cajas_emprendedores (id_cliente, nombre_caja, precio_caja) VALUES (?, ?, ?)", (nuevo_id, caja_clean, precio_caja_in))

                        conn.commit()
                        st.success(f"¡Registro de {nombre_clean} guardado exitosamente ({nuevo_id})!")
                        st.rerun()
                    else:
                        st.error("Debes ingresar el nombre del cliente.")
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown('<div class="form-card">', unsafe_allow_html=True)
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
                            st.markdown('<div style="background:#fdf2f8; padding:12px; border-radius:10px; border:1px solid #F3B2C9;">', unsafe_allow_html=True)
                            st.markdown(f"💼 **Configuración de Caja de Emprendedor:**")
                            col_cj1, col_cj2, col_cj3 = st.columns([2, 2, 1])
                            with col_cj1:
                                edit_nombre_caja = st.text_input("Nombre de Caja", value=caja_nombre_exp or "Caja #01", key=f"exp_cj_nom_{id_cli_exp}")
                            with col_cj2:
                                edit_precio_caja = st.number_input("Precio de Caja (₡)", value=float(caja_precio_exp), step=1000.0, key=f"exp_cj_pre_{id_cli_exp}")
                            with col_cj3:
                                st.write("")
                                if st.button("💾 Actualizar Caja", key=f"btn_update_cj_{id_cli_exp}"):
                                    c.execute("INSERT OR REPLACE INTO cajas_emprendedores (id_cliente, nombre_caja, precio_caja) VALUES (?, ?, ?)",
                                              (id_cli_exp, edit_nombre_caja.strip().capitalize(), edit_precio_caja))
                                    conn.commit()
                                    st.success("¡Caja actualizada!")
                                    st.rerun()
                            st.markdown('</div>', unsafe_allow_html=True)

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
                                st.markdown('<div class="form-card" style="border: 1px solid #F3B2C9; padding: 15px; margin-bottom: 10px;">', unsafe_allow_html=True)
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
                                st.markdown('</div>', unsafe_allow_html=True)
                else:
                    st.info("👆 Haz clic en cualquier cliente de la tabla superior para cargar automáticamente su expediente.")
            else:
                st.info("No hay clientes registrados.")
            st.markdown("</div>", unsafe_allow_html=True)

        # 3. Gestor Pedidos
        elif menu_principal == "✏️ Gestor Pedidos":
            st.markdown(
                """
                <div class="top-banner">
                    <div style="font-size: 18px; font-weight: 700;">Gestión y Edición de Pedidos por Cliente</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            st.markdown('<div class="form-card">', unsafe_allow_html=True)
            cli_pedidos = pd.read_sql("SELECT DISTINCT p.id_cliente, c.nombre FROM productos p JOIN clientes c ON p.id_cliente = c.id_cliente WHERE p.id_cliente IS NOT NULL AND p.id_cliente != ''", conn)

            if cli_pedidos.empty:
                st.info("No hay encargos asignados a clientes específicos.")
            else:
                dict_cli_p = {f"{r['id_cliente']} — {r['nombre']}": r['id_cliente'] for _, r in cli_pedidos.iterrows()}
                cli_p_sel = st.selectbox("Selecciona el cliente:", list(dict_cli_p.keys()), key="gestor_cli_sel")
                id_cli_gest = dict_cli_p[cli_p_sel]

                prods_de_cli = pd.read_sql("SELECT id, tienda, descripcion, precio, cantidad, estado FROM productos WHERE id_cliente = ? ORDER BY id DESC", conn, params=(id_cli_gest,))
                st.dataframe(prods_de_cli, use_container_width=True)

                prod_options = {f"Pedido #{r['id']} - {r['descripcion']} (₡{(r['precio'] or 0.0):,.0f})": r['id'] for _, r in prods_de_cli.iterrows()}
                selected_prod_label = st.selectbox("Seleccionar encargo a modificar:", list(prod_options.keys()), key="gestor_prod_sel")
                selected_id = prod_options[selected_prod_label]

                c.execute("SELECT tienda, categoria, descripcion, precio, cantidad, estado FROM productos WHERE id = ?", (selected_id,))
                p_data = c.fetchone()

                if p_data:
                    estados_opciones = ["🇺🇸 Comprado en USA", "📦 En tránsito", "🇨🇷 Recibido en CR", "✅ Entregado"]
                    idx_estado = estados_opciones.index(p_data[5]) if p_data[5] in estados_opciones else 0

                    col_e1, col_e2 = st.columns(2)
                    with col_e1:
                        edit_tienda = st.text_input("Tienda", value=p_data[0] or "", key="ped_edit_tienda")
                        edit_desc = st.text_input("Descripción", value=p_data[2] or "", key="ped_edit_desc")
                        edit_precio = st.number_input("Precio (₡)", value=float(p_data[3]) if p_data[3] is not None else 0.0, step=500.0, key="ped_edit_precio")
                    with col_e2:
                        edit_cat = st.text_input("Categoría", value=p_data[1] if p_data[1] else "", key="ped_edit_cat")
                        edit_cant = st.number_input("Cantidad", value=int(p_data[4]) if p_data[4] is not None else 1, step=1, key="ped_edit_cant")
                        edit_est = st.selectbox("Estado", estados_opciones, index=idx_estado, key="ped_edit_est")

                    col_b1, col_b2 = st.columns(2)
                    with col_b1:
                        if st.button("💾 Guardar Cambios", key="btn_ped_save"):
                            c.execute("UPDATE productos SET tienda=?, categoria=?, descripcion=?, precio=?, cantidad=?, estado=? WHERE id=?", 
                                      (edit_tienda.strip().capitalize(), edit_cat.strip().capitalize(), edit_desc.strip().capitalize(), edit_precio, edit_cant, edit_est, selected_id))
                            conn.commit()
                            st.success("¡Pedido actualizado correctamente!")
                            st.rerun()
                    with col_b2:
                        if st.button("🗑 Eliminar Pedido", key="btn_ped_del"):
                            c.execute("DELETE FROM productos WHERE id=?", (selected_id,))
                            conn.commit()
                            st.warning("Pedido eliminado.")
                            st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

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

            st.markdown('<div class="form-card">', unsafe_allow_html=True)
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
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown('<div class="form-card">', unsafe_allow_html=True)
            st.markdown("<h4 style='color:#be185d;'>📋 Historial de Abonos Recibidos</h4>", unsafe_allow_html=True)
            abonos_detalle = pd.read_sql("SELECT a.fecha as Fecha, c.id_cliente as Código, c.nombre as Cliente, a.monto_crc as 'Monto Abonado (₡)' FROM abonos a JOIN clientes c ON a.id_cliente = c.id_cliente ORDER BY a.id DESC", conn)
            if not abonos_detalle.empty:
                st.dataframe(abonos_detalle, use_container_width=True)
            else:
                st.info("No hay abonos registrados en el sistema.")
            st.markdown("</div>", unsafe_allow_html=True)

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

            st.markdown('<div class="form-card">', unsafe_allow_html=True)
            cat_gasto = st.selectbox("Categoría del Gasto", ["🏨 Hospedaje / Hotel", "🍽️ Comida / Alimentación", "🚗 Transporte / Combustible", "✈️ Fletes / Envíos USA-CR", "🛃 Aduana / Impuestos", "📦 Material de Empaque", "💡 Servicios y Operación", "🧩 Otros Gastos"], key="gasto_cat_sel")
            concepto = st.text_input("Concepto o Descripción", placeholder="Ej. Noche en hotel Miami...", key="gasto_conc_txt")
            monto_gasto = st.number_input("Monto en Colones (₡ CRC)", min_value=0.0, step=500.0, key="gasto_monto_num")
            obs_gasto = st.text_area("Notas / Observaciones adicionales", height=60, placeholder="Ej. Factura #1024", key="gasto_obs_txt")

            st.write("")
            if st.button("💾 Registrar Gasto Operativo", key="btn_save_gasto"):
                if concepto and monto_gasto > 0:
                    c.execute("INSERT INTO gastos (concepto, categoria, monto_crc, fecha, observaciones) VALUES (?, ?, ?, ?, ?)",
                              (concepto.strip().capitalize(), cat_gasto, monto_gasto, datetime.now().strftime("%Y-%m-%d"), obs_gasto.strip().capitalize() if obs_gasto else ""))
                    conn.commit()
                    st.success("¡Gasto registrado e integrado correctamente!")
                else:
                    st.error("Por favor completa la descripción y un monto superior a 0.")
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("<h4 style='color:#be185d;'>Últimos Gastos Registrados</h4>", unsafe_allow_html=True)
            gastos_df = pd.read_sql("SELECT fecha as Fecha, concepto as Concepto, categoria as Categoria, monto_crc as 'Monto (CRC)', observaciones as Observaciones FROM gastos ORDER BY id DESC LIMIT 10", conn)
            if not gastos_df.empty:
                st.dataframe(gastos_df, use_container_width=True)

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
            st.markdown('<div class="form-card">', unsafe_allow_html=True)
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
            st.markdown("</div>", unsafe_allow_html=True)

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

            st.markdown('<div class="form-card">', unsafe_allow_html=True)
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
                    prod_vr_clean = nombre_prod_vr.strip().capitalize()
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
            st.markdown("</div>", unsafe_allow_html=True)

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
            
            st.markdown('<div class="form-card">', unsafe_allow_html=True)
            st.markdown("<h4 style='color:#be185d;'>➕ Añadir Nuevo Producto al Inventario</h4>", unsafe_allow_html=True)
            
            with st.form("form_agregar_stock"):
                col_i1, col_i2 = st.columns(2)
                with col_i1:
                    new_desc = st.text_input("Descripción del Producto", placeholder="Ej. Blusa elegante")
                    new_precio = st.number_input("Precio (₡ CRC)", min_value=0.0, value=5000.0, step=500.0)
                    new_cant = st.number_input("Stock Inicial", min_value=1, value=1, step=1)
                with col_i2:
                    new_tienda = st.text_input("Tienda / Proveedor", placeholder="Ej. Zara o Local")
                    new_cat = st.text_input("Categoría", placeholder="Ej. Ropa")
                    new_barcode = st.text_input("Código de Barras (Opcional)", placeholder="Escanea o escribe el código")
                
                new_obs = st.text_area("Observaciones", placeholder="Detalles adicionales...", height=60)
                
                btn_add_stock = st.form_submit_button("💾 Guardar Producto en Inventario")
                
                if btn_add_stock:
                    if new_desc.strip():
                        c.execute("""INSERT INTO productos (id_cliente, tienda, categoria, descripcion, precio, moneda, cantidad, estado, observaciones, codigo_barras)
                                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                  (None, new_tienda.strip().capitalize(), new_cat.strip().capitalize(), new_desc.strip().capitalize(), new_precio, "CRC", new_cant, "Disponible en Inventario", new_obs.strip().capitalize() if new_obs else "", new_barcode.strip() if new_barcode else ""))
                        conn.commit()
                        st.success(f"¡Producto '{new_desc}' añadido al inventario exitosamente!")
                        st.rerun()
                    else:
                        st.error("Debes ingresar al menos la descripción del producto.")
            st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown('<div class="form-card">', unsafe_allow_html=True)
            st.markdown("<h4 style='color:#be185d;'>📊 Tabla de Inventario (Stock General)</h4>", unsafe_allow_html=True)
            df_stock_table = pd.read_sql("SELECT id, codigo_barras as 'Cód. Barras', descripcion as Producto, tienda as Tienda, categoria as Categoría, precio as Precio, cantidad as Stock FROM productos WHERE id_cliente IS NULL OR id_cliente = ''", conn)
            st.dataframe(df_stock_table, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown('<div class="form-card">', unsafe_allow_html=True)
            st.markdown("<h4 style='color:#be185d;'>✏️ Editar Producto del Stock</h4>", unsafe_allow_html=True)
            if not df_stock_table.empty:
                prod_edit_dict = {f"ID: {r['id']} - {r['Producto']}": r['id'] for _, r in df_stock_table.iterrows()}
                prod_edit_sel = st.selectbox("Seleccionar producto a editar:", list(prod_edit_dict.keys()), key="inv_edit_sel")
                id_prod_edit = prod_edit_dict[prod_edit_sel]

                c.execute("SELECT descripcion, tienda, categoria, precio, cantidad, codigo_barras FROM productos WHERE id = ?", (id_prod_edit,))
                p_edit_data = c.fetchone()
                if p_edit_data:
                    c_e1, c_e2 = st.columns(2)
                    with c_e1:
                        n_desc = st.text_input("Descripción", value=p_edit_data[0] or "", key="inv_ed_desc")
                        n_precio = st.number_input("Precio", value=float(p_edit_data[3]) if p_edit_data[3] is not None else 0.0, step=500.0, key="inv_ed_pre")
                        n_cant = st.number_input("Stock", value=int(p_edit_data[4]) if p_edit_data[4] is not None else 1, step=1, key="inv_ed_can")
                    with c_e2:
                        n_tienda = st.text_input("Tienda", value=p_edit_data[1] if p_edit_data[1] else "", key="inv_ed_tie")
                        n_cat = st.text_input("Categoría", value=p_edit_data[2] if p_edit_data[2] else "", key="inv_ed_cat")
                        n_barcode = st.text_input("Código de Barras", value=p_edit_data[5] if p_edit_data[5] else "", key="inv_ed_bar")

                    c_b1, c_b2 = st.columns(2)
                    with c_b1:
                        if st.button("💾 Guardar Cambios de Stock", key="btn_inv_save"):
                            c.execute("UPDATE productos SET descripcion=?, tienda=?, categoria=?, precio=?, cantidad=?, codigo_barras=? WHERE id=?", 
                                      (n_desc.strip().capitalize(), n_tienda.strip().capitalize(), n_cat.strip().capitalize(), n_precio, n_cant, n_barcode.strip() if n_barcode else "", id_prod_edit))
                            conn.commit()
                            st.success("¡Producto en stock actualizado!")
                            st.rerun()
                    with c_b2:
                        if st.button("🗑 Eliminar Producto", key="btn_inv_del"):
                            c.execute("DELETE FROM productos WHERE id=?", (id_prod_edit,))
                            conn.commit()
                            st.warning("Producto eliminado del inventario.")
                            st.rerun()
            else:
                st.info("No hay productos en el stock general para editar.")
            st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown('<div class="form-card">', unsafe_allow_html=True)
            st.markdown("<h4 style='color:#be185d;'>📸 Galería de Productos Locales</h4>", unsafe_allow_html=True)
            df_stock = pd.read_sql("SELECT id, descripcion, precio, cantidad, categoria, tienda, foto_path, codigo_barras FROM productos WHERE id_cliente IS NULL OR id_cliente = ''", conn)
            
            if df_stock.empty:
                st.info("No hay productos generales en stock actualmente.")
            else:
                for _, r in df_stock.iterrows():
                    st.markdown('<div class="form-card" style="border: 2px solid #F3B2C9; padding: 16px; border-radius: 14px; margin-bottom: 15px;">', unsafe_allow_html=True)
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
                    st.markdown('</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

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

    col_cli1, col_cli2 = st.columns([3, 1])
    with col_cli1:
        st.markdown(f"<h3 style='color: #be185d;'>🛍️ Hola, {client_name}</h3>", unsafe_allow_html=True)
    with col_cli2:
        if st.button("🚪 Salir", key="btn_salir_client"):
            st.session_state.user_role = None
            st.session_state.current_client = None
            st.rerun()

    st.markdown('<div class="form-card">', unsafe_allow_html=True)
    st.markdown("<h4 style='color:#be185d;'>💰 Estado de Cuenta</h4>", unsafe_allow_html=True)
    if caja_nombre_c:
        st.caption(f"💼 Caja Asignada: **{caja_nombre_c}** | Cuota: **₡{caja_precio_c:,.0f}**")
        
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Total Cargos", f"₡{total_compras:,.0f}")
    col_m2.metric("Total Abonado", f"₡{total_abonos:,.0f}")
    col_m3.metric("Saldo Pendiente", f"₡{max(0.0, saldo_pendiente):,.0f}")
    st.markdown("</div>", unsafe_allow_html=True)

    notifs = pd.read_sql("SELECT titulo, mensaje, fecha FROM notificaciones WHERE id_cliente = ? ORDER BY id DESC LIMIT 5", conn, params=(id_cli,))
    if not notifs.empty:
        st.markdown('<div class="form-card">', unsafe_allow_html=True)
        st.markdown("<h4 style='color:#be185d;'>🔔 Notificaciones Recientes</h4>", unsafe_allow_html=True)
        for _, row in notifs.iterrows():
            st.markdown(f"**{row['titulo']}** ({row['fecha']})")
            st.caption(row["mensaje"])
            st.divider()
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="form-card">', unsafe_allow_html=True)
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
    st.markdown("</div>", unsafe_allow_html=True)
