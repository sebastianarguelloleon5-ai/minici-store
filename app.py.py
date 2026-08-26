import sqlite3
import os
import pandas as pd
import streamlit as st
from datetime import datetime

# Configuración de página
st.set_page_config(
    page_title="Minici Store", 
    page_icon="🛍️", 
    layout="centered", 
    initial_sidebar_state="collapsed"
)

# 1. Base de datos SQLite y Directorios
conn = sqlite3.connect("minici_store.db", check_same_thread=False)
c = conn.cursor()

c.execute('''CREATE TABLE IF NOT EXISTS clientes (
                id_cliente TEXT PRIMARY KEY,
                nombre TEXT NOT NULL,
                telefono TEXT,
                correo TEXT)''')

try:
    c.execute("ALTER TABLE clientes ADD COLUMN correo TEXT")
    conn.commit()
except sqlite3.OperationalError:
    pass

c.execute('''CREATE TABLE IF NOT EXISTS productos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                id_cliente TEXT,
                tienda TEXT,
                categoria TEXT,
                descripcion TEXT,
                precio REAL,
                moneda TEXT,
                cantidad INTEGER,
                estado TEXT,
                id_caja TEXT,
                observaciones TEXT,
                foto_path TEXT,
                FOREIGN KEY(id_cliente) REFERENCES clientes(id_cliente))''')

c.execute('''CREATE TABLE IF NOT EXISTS abonos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                id_cliente TEXT,
                monto_crc REAL,
                fecha TEXT)''')

c.execute('''CREATE TABLE IF NOT EXISTS notificaciones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                id_cliente TEXT,
                titulo TEXT,
                mensaje TEXT,
                leida INTEGER DEFAULT 0,
                fecha TEXT)''')
conn.commit()

if not os.path.exists("fotos_productos"):
    os.makedirs("fotos_productos")

# Datos de prueba iniciales
c.execute("SELECT COUNT(*) FROM clientes")
if c.fetchone()[0] == 0:
    c.execute("INSERT INTO clientes (id_cliente, nombre, telefono, correo) VALUES ('MIN-0001', 'María López', '8888-8888', 'maria@example.com')")
    c.execute("INSERT INTO clientes (id_cliente, nombre, telefono, correo) VALUES ('MIN-0002', 'Ana Pérez', '7777-7777', 'ana@example.com')")
    c.execute("INSERT INTO notificaciones (id_cliente, titulo, mensaje, fecha) VALUES ('MIN-0001', '¡Bienvenida!', 'Tu cuenta ha sido creada con éxito en Minici Store.', '2026-06-06')")
    conn.commit()

# 2. Estilos CSS 100% Compatibles (Compu y Celular)
st.markdown("""
<style>
    .stApp {
        background-color: #f8fafc;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }
    
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    .top-banner {
        background: linear-gradient(135deg, #f472b6 0%, #db2777 100%);
        padding: 20px;
        border-radius: 16px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(219, 39, 119, 0.2);
    }

    .form-card {
        background: #ffffff;
        padding: 20px;
        border-radius: 16px;
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.04);
        margin-bottom: 20px;
        border: 1px solid #f1f5f9;
    }

    .section-title {
        font-size: 13px;
        font-weight: 700;
        color: #475569;
        margin-bottom: 4px;
        margin-top: 10px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    .stTextInput>div>div>input, 
    .stSelectbox>div>div>select, 
    .stNumberInput>div>div>input, 
    .stTextArea>div>div>textarea {
        border-radius: 10px !important;
        border: 1px solid #cbd5e1 !important;
        background-color: #ffffff !important;
        color: #0f172a !important;
        font-size: 15px !important;
        min-height: 42px !important;
    }

    div.stButton > button:first-child {
        background-color: #ec4899;
        color: white;
        border-radius: 12px;
        font-weight: 700;
        border: none;
        padding: 12px 20px;
        width: 100%;
        box-shadow: 0 4px 12px rgba(236, 72, 153, 0.25);
        font-size: 15px;
        min-height: 46px;
        cursor: pointer;
    }
    div.stButton > button:first-child:hover {
        background-color: #db2777;
    }

    div[data-testid="stRadio"] > div {
        flex-direction: row !important;
        gap: 10px !important;
        flex-wrap: wrap !important;
    }
    div[data-testid="stRadio"] label {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        padding: 10px 16px;
        border-radius: 12px;
        font-weight: 600;
        color: #334155;
        cursor: pointer;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }

    @media (max-width: 768px) {
        .row-widget.stHorizontal {
            flex-direction: column !important;
        }
        .stColumn {
            width: 100% !important;
            flex: 100% !important;
            max-width: 100% !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# 3. Control de Sesión
if "user_role" not in st.session_state:
    st.session_state.user_role = None
if "current_client" not in st.session_state:
    st.session_state.current_client = None

# --- PANTALLA DE BIENVENIDA E INICIO DE SESIÓN ---
if st.session_state.user_role is None:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        # Busca el logo con diferentes nombres o extensiones comunes
        logo_path = None
        for posible_nombre in ["logo.jpg", "logo.png", "logo.jpeg", "21237.jpg"]:
            if os.path.exists(posible_nombre):
                logo_path = posible_nombre
                break
        
        if logo_path:
            st.image(logo_path, use_container_width=True)
        else:
            st.markdown("<h1 style='text-align: center; color: #db2777;'>🛍️ Minici Store</h1>", unsafe_allow_html=True)
            st.info("💡 Coloca tu imagen como 'logo.jpg' en la misma carpeta para ver el logo aquí.")

        st.markdown("""
        <div style="text-align: center; margin-bottom: 15px;">
            <h2 style="color: #db2777; font-size: 22px; margin-bottom: 4px;">¡Bienvenidos a Minici Store!</h2>
            <p style="color: #64748b; font-size: 14px;">Tu Personal Shopper de confianza. Ingresa tu código de acceso para continuar.</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown('<div class="form-card" style="text-align: center;">', unsafe_allow_html=True)
        codigo_ingresado = st.text_input("Código de acceso", placeholder="Ej. MIN-0001", label_visibility="collapsed")
        
        st.write("")
        if st.button("🚀 Ingresar al Sistema"):
            codigo_limpio = codigo_ingresado.strip()
            ADMIN_CODE = "Kendra5412"
            
            if codigo_limpio == ADMIN_CODE:
                st.session_state.user_role = "admin"
                st.rerun()
            else:
                codigo_cli = codigo_limpio.upper()
                c.execute("SELECT id_cliente FROM clientes WHERE id_cliente = ?", (codigo_cli,))
                res = c.fetchone()
                if res:
                    st.session_state.user_role = "client"
                    st.session_state.current_client = res[0]
                    st.rerun()
                else:
                    st.error("❌ Código incorrecto o no registrado.")
        st.markdown('</div>', unsafe_allow_html=True)

# --- MÓDULO ADMINISTRADOR ---
elif st.session_state.user_role == "admin":
    col_a, col_b = st.columns([3, 1])
    with col_a:
        st.markdown("### ⚙️ Panel Administrador")
    with col_b:
        if st.button("🚪 Salir"):
            st.session_state.user_role = None
            st.rerun()

    menu_admin = st.radio(
        "Navegación Admin",
        ["📸 Registrar Compra", "👩 Registrar Clienta", "📦 Control de Cajas", "💰 Registrar Abonos"],
        horizontal=True,
        label_visibility="collapsed"
    )
    st.write("")

    if menu_admin == "📸 Registrar Compra":
        st.markdown("""
        <div class="top-banner">
            <div style="font-size: 18px; font-weight: 700;">Registrar nueva compra</div>
            <div style="font-size: 13px; opacity: 0.9;">Agrega un producto para una clienta</div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<div class="form-card">', unsafe_allow_html=True)
        
        st.markdown('<p class="section-title">1. Seleccionar clienta</p>', unsafe_allow_html=True)
        clientes_df = pd.read_sql("SELECT id_cliente || ' — ' || nombre AS display, id_cliente, nombre FROM clientes", conn)
        
        if clientes_df.empty:
            st.warning("No hay clientas registradas.")
        else:
            cli_selected = st.selectbox("Seleccionar clienta", clientes_df['display'], label_visibility="collapsed")
            id_cliente = clientes_df[clientes_df['display'] == cli_selected]['id_cliente'].values[0]
            nombre_clienta = clientes_df[clientes_df['display'] == cli_selected]['nombre'].values[0]
            
            st.markdown(f"""
            <div style="background: #fdf2f8; padding: 10px 14px; border-radius: 12px; border: 1px solid #fbcfe8; margin-top: 6px; font-size: 14px;">
                <b>{nombre_clienta}</b> <span style="background: #fce7f3; color: #be185d; padding: 2px 6px; border-radius: 6px; font-size: 11px; font-weight: bold; float: right;">{id_cliente}</span>
            </div>
            """, unsafe_allow_html=True)

        col_tienda, col_cat = st.columns(2)
        with col_tienda:
            st.markdown('<p class="section-title">2. Tienda</p>', unsafe_allow_html=True)
            tienda = st.selectbox("Tienda", ["Zara", "Guess", "Adidas", "Shein", "Amazon", "Nike", "Victoria's Secret", "Otra"], label_visibility="collapsed")
        with col_cat:
            st.markdown('<p class="section-title">3. Categoría</p>', unsafe_allow_html=True)
            categoria = st.selectbox("Categoría", ["Seleccionar categoría", "Vestido", "Bolso", "Tenis", "Blusa", "Cosméticos", "Accesorios"], label_visibility="collapsed")

        st.markdown('<p class="section-title">4. Producto</p>', unsafe_allow_html=True)
        producto = st.text_input("Producto", placeholder="Ej. Vestido largo estampado", label_visibility="collapsed")

        col_c, col_d = st.columns(2)
        with col_c:
            st.markdown('<p class="section-title">5. Precio ($ USD)</p>', unsafe_allow_html=True)
            precio = st.number_input("Precio", min_value=0.0, value=35.00, step=1.0, label_visibility="collapsed")
        with col_d:
            st.markdown('<p class="section-title">6. Cantidad</p>', unsafe_allow_html=True)
            cantidad = st.number_input("Cantidad", min_value=1, value=1, step=1, label_visibility="collapsed")

        col_est, col_caja = st.columns(2)
        with col_est:
            st.markdown('<p class="section-title">7. Estado del producto</p>', unsafe_allow_html=True)
            estado = st.selectbox("Estado", ["🇺🇸 Comprado en USA", "📦 En tránsito", "🇨🇷 Recibido en CR", "✅ Entregado"], label_visibility="collapsed")
        with col_caja:
            st.markdown('<p class="section-title">8. Asignar a caja</p>', unsafe_allow_html=True)
            caja = st.selectbox("Caja", ["Seleccionar caja", "Caja C01", "Caja C02", "Caja C03", "Caja C04"], label_visibility="collapsed")

        st.markdown('<p class="section-title">9. Fotos del producto</p>', unsafe_allow_html=True)
        foto_file = st.camera_input("Tomar foto") or st.file_uploader("Subir desde galería", type=["jpg", "png", "jpeg"], label_visibility="collapsed")

        st.markdown('<p class="section-title">10. Observaciones</p>', unsafe_allow_html=True)
        observaciones = st.text_area("Observaciones", placeholder="Escribe detalles...", height=80, label_visibility="collapsed")

        st.write("")
        if st.button("💾 Guardar compra"):
            if not producto:
                st.error("Debes ingresar el nombre del producto.")
            else:
                foto_filename = ""
                if foto_file:
                    foto_filename = f"{id_cliente}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
                    filepath = os.path.join("fotos_productos", foto_filename)
                    with open(filepath, "wb") as f:
                        f.write(foto_file.getbuffer())
                
                c.execute("""INSERT INTO productos (id_cliente, tienda, categoria, descripcion, precio, moneda, cantidad, estado, id_caja, observaciones, foto_path)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                          (id_cliente, tienda, categoria, producto, precio, "USD", cantidad, estado, caja, observaciones, foto_filename))
                
                c.execute("""INSERT INTO notificaciones (id_cliente, titulo, mensaje, fecha)
                             VALUES (?, ?, ?, ?)""",
                          (id_cliente, "Nuevo pedido registrado", f"Se agregó '{producto}' ({tienda}) a tus compras con estado: {estado}.", datetime.now().strftime('%Y-%m-%d %H:%M')))
                conn.commit()
                st.success(f"¡Compra de '{producto}' registrada con éxito!")

        st.markdown('</div>', unsafe_allow_html=True)

    elif menu_admin == "👩 Registrar Clienta":
        st.markdown("""
        <div class="top-banner">
            <div style="font-size: 18px; font-weight: 700;">Registrar Nueva Clienta</div>
            <div style="font-size: 13px; opacity: 0.9;">Crea una nueva ficha de clienta</div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown('<div class="form-card">', unsafe_allow_html=True)
        
        c.execute("SELECT id_cliente FROM clientes")
        rows = c.fetchall()
        if not rows:
            nuevo_id = "MIN-0001"
        else:
            numeros = []
            for r in rows:
                try:
                    num = int(r[0].replace("MIN-", ""))
                    numeros.append(num)
                except:
                    pass
            siguiente_num = max(numeros) + 1 if numeros else 1
            nuevo_id = f"MIN-{siguiente_num:04d}"
            
        st.markdown(f"""
        <div style="background: #fdf2f8; padding: 12px 14px; border-radius: 12px; border: 1px solid #fbcfe8; margin-bottom: 15px;">
            <span style="color: #831843; font-size: 13px;">🏷️ Código generado automáticamente: <b>{nuevo_id}</b></span>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown('<p class="section-title">Nombre</p>', unsafe_allow_html=True)
        nombre = st.text_input("Nombre", placeholder="Escribe el nombre de la clienta...", label_visibility="collapsed")
        
        col_tel, col_cor = st.columns(2)
        with col_tel:
            st.markdown('<p class="section-title">Número</p>', unsafe_allow_html=True)
            tel = st.text_input("Número", placeholder="Escribe el número...", label_visibility="collapsed")
        with col_cor:
            st.markdown('<p class="section-title">Correo</p>', unsafe_allow_html=True)
            correo = st.text_input("Correo", placeholder="Escribe el correo...", label_visibility="collapsed")
        
        st.write("")
        if st.button("Guardar Clienta"):
            if nombre:
                c.execute("INSERT OR REPLACE INTO clientes (id_cliente, nombre, telefono, correo) VALUES (?, ?, ?, ?)", 
                          (nuevo_id, nombre, tel, correo))
                conn.commit()
                st.success(f"¡Clienta {nombre} guardada exitosamente ({nuevo_id})!")
            else:
                st.error("Debes ingresar al menos el nombre de la clienta.")
        st.markdown('</div>', unsafe_allow_html=True)

    elif menu_admin == "📦 Control de Cajas":
        st.markdown("""
        <div class="top-banner">
            <div style="font-size: 18px; font-weight: 700;">Control de Cajas</div>
            <div style="font-size: 13px; opacity: 0.9;">Revisa los productos asignados por caja</div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown('<div class="form-card">', unsafe_allow_html=True)
        cajas_df = pd.read_sql("SELECT DISTINCT id_caja FROM productos WHERE id_caja != 'Seleccionar caja'", conn)
        if cajas_df.empty:
            st.info("No hay cajas registradas con productos.")
        else:
            caja_sel = st.selectbox("Seleccionar Caja", cajas_df['id_caja'])
            items_caja = pd.read_sql("SELECT p.id_cliente, c.nombre, p.descripcion, p.tienda, p.estado FROM productos p JOIN clientes c ON p.id_cliente = c.id_cliente WHERE p.id_caja = ?", conn, params=(caja_sel,))
            st.dataframe(items_caja, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    elif menu_admin == "💰 Registrar Abonos":
        st.markdown("""
        <div class="top-banner">
            <div style="font-size: 18px; font-weight: 700;">Registrar Abono</div>
            <div style="font-size: 13px; opacity: 0.9;">Agrega pagos realizados por las clientas</div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown('<div class="form-card">', unsafe_allow_html=True)
        cli_list = pd.read_sql("SELECT id_cliente || ' - ' || nombre AS disp, id_cliente FROM clientes", conn)
        if cli_list.empty:
            st.warning("No hay clientas registradas.")
        else:
            sel = st.selectbox("Clienta", cli_list['disp'])
            id_c = sel.split(" - ")[0]
            monto = st.number_input("Monto abonado (CRC)", min_value=0.0, step=1000.0)
            st.write("")
            if st.button("Guardar Abono"):
                c.execute("INSERT INTO abonos (id_cliente, monto_crc, fecha) VALUES (?, ?, ?)", (id_c, monto, datetime.now().strftime('%Y-%m-%d')))
                conn.commit()
                st.success("Abono registrado con éxito.")
        st.markdown('</div>', unsafe_allow_html=True)

# --- MÓDULO CLIENTA ---
elif st.session_state.user_role == "client":
    id_cli = st.session_state.current_client
    
    c.execute("SELECT nombre FROM clientes WHERE id_cliente = ?", (id_cli,))
    res_cli = c.fetchone()
    client_name = res_cli[0] if res_cli else "Clienta"

    c.execute("SELECT id, titulo, mensaje, fecha FROM notificaciones WHERE id_cliente = ? AND leida = 0 ORDER BY id DESC", (id_cli,))
    notificaciones = c.fetchall()
    num_notis = len(notificaciones)

    st.markdown(f"""
    <div class="top-banner">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
            <span style="font-size: 14px; font-weight: bold;">🛍️ Minici Store</span>
            <span style="background: rgba(255,255,255,0.25); padding: 2px 8px; border-radius: 10px; font-size: 12px;">🔔 {num_notis}</span>
        </div>
        <div style="font-size: 20px; font-weight: bold;">Hola, {client_name} 👋</div>
        <div style="font-size: 12px; opacity: 0.9;">Tu código: <b>{id_cli}</b></div>
    </div>
    """, unsafe_allow_html=True)

    if st.button("🚪 Salir de mi cuenta"):
        st.session_state.user_role = None
        st.session_state.current_client = None
        st.rerun()

    if num_notis > 0:
        with st.expander(f"📬 Tienes {num_notis} notificación(es)", expanded=True):
            for noti in notificaciones:
                st.info(f"**{noti[1]}** ({noti[3]})\n\n{noti[2]}")
                if st.button(f"Marcar como leída #{noti[0]}", key=f"noti_{noti[0]}"):
                    c.execute("UPDATE notificaciones SET leida = 1 WHERE id = ?", (noti[0],))
                    conn.commit()
                    st.rerun()

    prods = pd.read_sql("SELECT descripcion, tienda, precio, moneda, estado, foto_path FROM productos WHERE id_cliente = ?", conn, params=(id_cli,))
    abonos_df = pd.read_sql("SELECT SUM(monto_crc) as total FROM abonos WHERE id_cliente = ?", conn, params=(id_cli,))
    total_abonos = abonos_df.iloc[0]['total'] if abonos_df.iloc[0]['total'] else 0.0
    
    tipo_cambio = 520
    total_compras_usd = prods['precio'].sum() if not prods.empty else 0.0
    total_compras_crc = total_compras_usd * tipo_cambio

    st.markdown("### 📦 Mis Pedidos")
    
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Pedidos", f"₡{total_compras_crc:,.0f}")
    col_m2.metric("Abonado", f"₡{total_abonos:,.0f}")
    col_m3.metric("Pendiente", f"₡{max(0, total_compras_crc - total_abonos):,.0f}")

    if prods.empty:
        st.info("Aún no tienes productos registrados.")
    else:
        for idx, row in prods.iterrows():
            st.markdown(f"""
            <div class="form-card" style="margin-bottom: 10px; padding: 12px;">
                <div style="font-weight: bold; font-size: 15px;">{row['descripcion']}</div>
                <div style="font-size: 12px; color: #64748b; margin-top: 2px;">Tienda: <b>{row['tienda']}</b> | Precio: ${row['precio']} USD</div>
                <div style="margin-top: 6px;"><span style="background: #fdf2f8; color: #db2777; padding: 2px 6px; border-radius: 6px; font-size: 11px; font-weight: bold;">{row['estado']}</span></div>
            </div>
            """, unsafe_allow_html=True)
            if row['foto_path'] and os.path.exists(os.path.join("fotos_productos", row['foto_path'])):
                st.image(os.path.join("fotos_productos", row['foto_path']), width=120)