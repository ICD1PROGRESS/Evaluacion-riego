import textwrap
import streamlit as st
from datetime import date, timedelta

from ..config import TRIAL_DIAS, TRIAL_MAX_PROYECTOS, PLANES
from ..models import init_db, seed_planes, seed_super_admin, Plan, engine
from ..auth import (
    autenticar_usuario,
    registrar_usuario_trial,
    iniciar_sesion_streamlit,
    obtener_usuario_actual,
    verificar_suscripcion_activa,
)
from ..limites import resumen_limites_sidebar

# ======================================================================
# 1. IDENTIDAD DE MARCA — DIGITAL PROGRESS
# ======================================================================

BRAND = {
    "nombre": "Digital Progress",
    "descriptor": "Ingeniería Inteligente para el Territorio",
    "eslogan": "Datos · Territorio · Decisiones",
    "claim": (
        "Ayudamos a organizaciones, productores y empresas de ingeniería a tomar "
        "decisiones más rápidas, precisas y rentables mediante geotecnologías, "
        "análisis espacial e inteligencia de datos."
    ),
    "email": "digitalprogress.org@gmail.com",
}

UNIDADES_NEGOCIO = [
    ("💧", "Ingeniería del Agua"),
    ("🌱", "Agricultura Digital"),
    ("🏛️", "Gestión Territorial y Municipal"),
    ("🤖", "Inteligencia Geoespacial y Automatización"),
]

FEATURES = [
    ("📊", "Evaluación VIPFE completa",
     "VAN, TIR, relación beneficio/costo y análisis de sensibilidad, calculados en minutos."),
    ("🌾", "Costos de producción por cultivo",
     "Presupuestos agrícolas precisos para cada alternativa productiva de riego."),
    ("💰", "Análisis de inversión detallado",
     "Flujos de caja, financiamiento y rentabilidad con trazabilidad total."),
    ("📑", "Reportes exportables",
     "Informes técnicos listos para financiadores, GAM y etapas de preinversión."),
]

PLAN_COLORES = [
    ("#1F6FB2", "#0FA3B1"),
    ("#0FA3B1", "#2E9E5B"),
    ("#2E9E5B", "#0A2E4E"),
    ("#0A2E4E", "#1F6FB2"),
]
PLAN_ICONOS_DEFAULT = ["🌱", "🚀", "🛰️", "🏛️"]

_NOMBRES_SLUG = {
    "trial": "Trial", "free": "Gratis", "gratis": "Gratis",
    "mensual": "Mensual", "anual": "Anual", "bianual": "Bianual", "bienal": "Bianual",
    "basico": "Básico", "básico": "Básico", "inicial": "Inicial", "starter": "Starter",
    "estandar": "Estándar", "estándar": "Estándar",
    "profesional": "Profesional", "pro": "Pro", "avanzado": "Avanzado",
    "premium": "Premium", "plus": "Plus",
    "municipal": "Municipal", "gobierno": "Gobierno",
    "empresarial": "Empresarial", "enterprise": "Enterprise", "corporativo": "Corporativo",
}

_ICONOS_SLUG = {
    "trial": "🎁", "free": "🎁", "gratis": "🎁",
    "mensual": "📆", "anual": "📅", "bianual": "🏆", "bienal": "🏆",
    "basico": "🌱", "inicial": "🌱",
    "profesional": "🚀", "pro": "🚀", "avanzado": "🚀",
    "premium": "⭐", "municipal": "🏛️", "enterprise": "🏢",
}

# ======================================================================
# 2. TEMA VISUAL (CSS)
# ======================================================================

BASE_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@600;700;800&family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family:'Inter',sans-serif; }
h1,h2,h3,h4 { font-family:'Sora',sans-serif; color:#0A2E4E; }
p { color:#33475B; }
.block-container { padding-top:1.4rem; padding-bottom:2.5rem; max-width:1220px; }
.dp-user{
background:linear-gradient(135deg,#0A2E4E 0%,#1F6FB2 100%);
border-radius:16px;padding:18px;margin-bottom:14px;color:#fff;
box-shadow:0 8px 22px rgba(10,46,78,.25);text-align:center;
}
.dp-user .u-avatar{font-size:30px;}
.dp-user .u-name{font-weight:700;font-size:15px;margin-top:6px;color:#fff;}
.dp-user .u-email{font-size:12px;opacity:.85;margin-top:2px;word-break:break-all;color:#EAF4FB;}
.dp-user .u-rol{
display:inline-block;margin-top:10px;padding:3px 12px;border-radius:999px;
background:rgba(255,255,255,.18);border:1px solid rgba(255,255,255,.30);
font-size:10.5px;letter-spacing:.8px;text-transform:uppercase;color:#fff;
}
.stButton > button, .stFormSubmitButton > button{
border-radius:12px;font-weight:600;transition:all .18s ease;
}
button[kind="primary"]{
background:linear-gradient(135deg,#1F6FB2 0%,#0FA3B1 100%) !important;
border:none !important;color:#fff !important;
box-shadow:0 8px 20px rgba(31,111,178,.30);
}
button[kind="primary"]:hover{
transform:translateY(-1px);color:#fff !important;
box-shadow:0 12px 26px rgba(15,163,177,.40);
}
div[data-baseweb="input"], div[data-baseweb="textarea"], div[data-baseweb="select"]{border-radius:10px;}
div[data-baseweb="input"]:focus-within{box-shadow:0 0 0 3px rgba(15,163,177,.16) !important;}
div[data-testid="stAlert"]{border-radius:12px;}
div[data-testid="stMetric"]{
background:#fff;border:1px solid #E3EAF2;border-radius:14px;
padding:14px 18px;box-shadow:0 3px 12px rgba(15,45,78,.05);
}
</style>
"""

PUBLIC_CSS = """
<style>
#MainMenu {visibility:hidden;}
footer {visibility:hidden;}
header[data-testid="stHeader"]{background:transparent;}
.dp-hero{
position:relative;overflow:hidden;color:#fff;
background:linear-gradient(135deg,#0A2E4E 0%,#1F6FB2 48%,#0FA3B1 100%);
border-radius:24px;padding:44px 44px 38px;
box-shadow:0 18px 45px rgba(10,46,78,.28);margin-bottom:20px;
}
.dp-hero::before{
content:"";position:absolute;inset:0;pointer-events:none;
background:
radial-gradient(circle at 88% 12%, rgba(255,255,255,.16) 0, transparent 30%),
radial-gradient(circle at 8% 95%, rgba(46,158,91,.40) 0, transparent 42%);
}
.dp-hero > *{position:relative;}
.dp-hero h1{font-size:2.8rem;font-weight:800;margin:0;line-height:1.05;color:#fff;}
.dp-hero h1 span{color:#7FF3D4;}
.dp-badge{
display:inline-block;background:rgba(255,255,255,.14);
border:1px solid rgba(255,255,255,.35);padding:6px 14px;border-radius:999px;
font-size:11.5px;font-weight:600;letter-spacing:1.2px;margin-bottom:16px;color:#fff;
}
.dp-hero .dp-desc{font-size:1.15rem;margin:12px 0 2px;opacity:.95;font-weight:600;color:#fff;}
.dp-hero .dp-sub{font-size:.95rem;margin:0;opacity:.82;color:#fff;}
.dp-hero .dp-slogan{
margin:18px 0 0;font-size:.8rem;letter-spacing:3px;text-transform:uppercase;
color:#9ED9FF;font-weight:700;
}
.dp-units{display:flex;flex-wrap:wrap;gap:10px;margin:4px 0 24px;}
.dp-unit{
background:#fff;border:1px solid #E3EAF2;border-radius:999px;padding:8px 16px;
font-size:12.5px;font-weight:600;color:#0A2E4E;
box-shadow:0 2px 8px rgba(15,45,78,.06);
}
.dp-card{
background:#fff;border:1px solid #E3EAF2;border-radius:16px;
padding:22px 24px;box-shadow:0 6px 22px rgba(15,45,78,.07);margin-bottom:18px;
}
.dp-card-title{font-family:'Sora',sans-serif;font-weight:700;font-size:1.05rem;color:#0A2E4E;margin-bottom:8px;}
.dp-claim{margin:0;color:#33475B;line-height:1.65;font-size:.95rem;}
.dp-feat{display:flex;gap:14px;margin-bottom:14px;align-items:flex-start;}
.dp-feat-ico{
flex:0 0 42px;height:42px;display:flex;align-items:center;justify-content:center;
font-size:20px;background:#EAF6F4;border:1px solid #D2EBE6;border-radius:12px;
}
.dp-feat-title{font-weight:700;color:#0A2E4E;font-size:.95rem;}
.dp-feat-desc{color:#5A6B7B;font-size:.86rem;line-height:1.5;}
.dp-quote{
background:#F0F9FF;border-left:4px solid #0FA3B1;border-radius:0 12px 12px 0;
padding:14px 18px;color:#33475B;font-size:.9rem;margin:6px 0 18px;
}
.dp-card-cta{
background:linear-gradient(135deg,#0FA3B1 0%,#2E9E5B 100%);
color:#fff;border-radius:16px;padding:22px 24px;
box-shadow:0 10px 26px rgba(46,158,91,.30);
}
.dp-card-cta h4, .dp-card-cta p{color:#fff;}
div[data-testid="stVerticalBlockBorderWrapper"]{
background:#fff !important;border:1px solid #E3EAF2 !important;
border-radius:20px !important;box-shadow:0 14px 40px rgba(10,46,78,.10) !important;
}
div[data-testid="stForm"]{border:none !important;background:transparent !important;}
.dp-auth-head{
font-size:1.02rem;color:#0A2E4E;background:#F4F7FA;
border:1px solid #E3EAF2;border-radius:12px;padding:10px 16px;margin-bottom:14px;
}
.dp-mini-note{
font-size:.85rem;color:#33475B;background:#F0F9FF;border:1px solid #D6EEF5;
border-radius:10px;padding:10px 14px;margin-bottom:14px;
}
.dp-trust{text-align:center;color:#8598AA;font-size:.78rem;margin-top:12px;}
.stTabs [data-baseweb="tab-list"]{
background:#F4F7FA;border:1px solid #E3EAF2;border-radius:14px;padding:6px;gap:6px;
}
.stTabs [data-baseweb="tab"]{
border-radius:10px;padding:10px 14px;font-weight:600;color:#5A6B7B;background:transparent;
}
.stTabs [data-baseweb="tab"][aria-selected="true"]{
background:#fff;color:#0A2E4E;box-shadow:0 3px 10px rgba(15,45,78,.12);
}
.stTabs [data-baseweb="tab-highlight"]{
height:3px;background:linear-gradient(90deg,#0FA3B1,#2E9E5B);border-radius:3px;
}
.dp-section-head{margin:38px 0 20px;}
.dp-section-kicker{color:#0FA3B1;font-weight:700;letter-spacing:2.5px;font-size:11px;}
.dp-plan{
position:relative;background:#fff;border:1px solid #E3EAF2;border-radius:18px;
padding:30px 20px 24px;text-align:center;height:100%;
box-shadow:0 8px 24px rgba(15,45,78,.09);
transition:transform .18s ease, box-shadow .18s ease;
}
.dp-plan:hover{transform:translateY(-6px);box-shadow:0 18px 38px rgba(15,45,78,.16);}
.dp-plan.featured{border:2px solid #0FA3B1;box-shadow:0 12px 32px rgba(15,163,177,.24);}
.dp-plan .tag{
position:absolute;top:-12px;left:50%;transform:translateX(-50%);
background:linear-gradient(135deg,#0FA3B1,#2E9E5B);color:#fff;
font-size:10.5px;font-weight:700;padding:4px 14px;border-radius:999px;
letter-spacing:1px;white-space:nowrap;
}
.dp-plan .dp-icon{font-size:30px;margin-bottom:8px;}
.dp-plan h4{margin:0 0 10px;color:#0A2E4E;}
.dp-plan .dp-bar{height:5px;width:64px;border-radius:99px;margin:0 auto 14px;}
.dp-plan .price{font-family:'Sora',sans-serif;font-size:1.9rem;font-weight:800;color:#1F6FB2;}
.dp-plan .price small{font-size:.8rem;color:#5A6B7B;font-weight:500;}
.dp-plan-meta{margin-top:12px;}
.dp-plan-meta span{display:block;color:#5A6B7B;font-size:.86rem;margin:5px 0;}
.dp-footer{text-align:center;padding:32px 0 8px;color:#5A6B7B;font-size:13px;}
.dp-footer .brand{color:#0A2E4E;font-weight:700;}
.dp-footer .slogan{
display:block;margin-top:6px;letter-spacing:2.5px;text-transform:uppercase;
font-size:10.5px;color:#0FA3B1;font-weight:700;
}
[data-testid="stSidebar"]{
background:linear-gradient(180deg,#08243D 0%,#0E3A5F 55%,#11506B 100%);
}
.dp-sb-brand{
background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.15);
border-radius:16px;padding:18px;text-align:center;margin-bottom:16px;
}
.dp-sb-logo{font-size:34px;}
.dp-sb-name{font-family:'Sora',sans-serif;font-weight:800;font-size:19px;color:#fff;margin-top:6px;}
.dp-sb-tag{font-size:10.5px;color:#9ED9C8;letter-spacing:.8px;text-transform:uppercase;margin-top:5px;}
.dp-sb-item{
color:#DCEBF7;font-size:13px;padding:9px 12px;border-radius:10px;margin-bottom:6px;
background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.08);
}
.dp-sb-contact{
margin-top:16px;color:#B9D4E8;font-size:12px;line-height:1.8;
background:rgba(255,255,255,.05);border-radius:12px;padding:14px;
}
.dp-sb-slogan{color:#7FF3D4;letter-spacing:1.5px;font-size:10px;text-transform:uppercase;font-weight:700;}
</style>
"""
def _css(publico: bool = True) -> str:
    """Devuelve el tema CSS: base (+ estilos de pantalla pública si aplica)."""
    return BASE_CSS + (PUBLIC_CSS if publico else "")

def _md(html: str):

    st.markdown(textwrap.dedent(html), unsafe_allow_html=True)

# ======================================================================
# 3. HELPERS DE PLANES
# ======================================================================
def _plan_get(plan: dict, clave: str, default=None):
    """Accede a una clave del plan tolerando espacios invisibles (ej. 'nombre ')."""
    if clave in plan:
        return plan[clave]
    for k, v in plan.items():
        if str(k).strip().lower() == clave.lower():
            return v
    return default

def _formatear_precio(precio) -> str:
    """Muestra $99 en vez de $99.0; conserva decimales solo si existen."""
    try:
        p = float(precio)
    except (TypeError, ValueError):
        return "—"
    return f"${p:,.0f}" if p.is_integer() else f"${p:,.2f}"


def _nombre_amigable(slug: str, plan: dict) -> str:
    """Devuelve un nombre legible para el plan (nunca códigos/slugs crudos)."""
    nombre = str(_plan_get(plan, "nombre") or "").strip()
    nombre = " ".join(nombre.split())
    if nombre and "_" not in nombre and not (nombre.isupper() and any(c.isdigit() for c in nombre)):
        return nombre
    clave = str(slug).strip().lower().replace("plan_", "").replace("plan-", "")
    clave = " ".join(clave.replace("_", " ").replace("-", " ").split())
    bonito = " ".join(_NOMBRES_SLUG.get(p, p) for p in clave.split())
    return bonito.title() if bonito else "Plan"


def _icono_plan(slug: str, indice: int) -> str:
    """Ícono coherente según el tipo de plan (anual→📅, bianual→🏆, etc.)."""
    s = str(slug).strip().lower()
    for clave, icono in _ICONOS_SLUG.items():
        if clave in s:
            return icono
    return PLAN_ICONOS_DEFAULT[indice % len(PLAN_ICONOS_DEFAULT)]

def _plan_card(slug: str, plan: dict, indice: int = 0, destacado: bool = False) -> str:
    """Card moderna para cada plan — precio limpio, ícono y valor mensual."""
    nombre = _nombre_amigable(slug, plan)
    icono = _icono_plan(slug, indice)
    c1, c2 = PLAN_COLORES[indice % len(PLAN_COLORES)]

    precio = _plan_get(plan, "precio_usd", 0) or 0
    meses = _plan_get(plan, "duracion_meses", None)
    max_proy = _plan_get(plan, "max_proyectos", "—")

    try:
        precio_num = float(precio)
    except (TypeError, ValueError):
        precio_num = 0.0

    if precio_num <= 0:
        precio_html, sufijo = "Gratis", ""
    else:
        precio_html, sufijo = _formatear_precio(precio_num), "<small> Bs</small>"

    mensual_html = ""
    if precio_num > 0 and meses:
        try:
            por_mes = precio_num / int(meses)
            mensual_html = f"<span>💡 Equivale a <b>{_formatear_precio(por_mes)}</b>/mes</span>"
        except Exception:
            pass

    clase = "dp-plan featured" if destacado else "dp-plan"
    tag = '<span class="tag">⭐ MÁS POPULAR</span>' if destacado else ""

    # IMPORTANTE: HTML a columna 0 y sin líneas en blanco internas
    return f"""
<div class="{clase}" style="border-top:5px solid {c1};">{tag}
<div class="dp-icon">{icono}</div>
<h4>{nombre}</h4>
<div class="dp-bar" style="background:linear-gradient(90deg,{c1},{c2});"></div>
<div class="price">{precio_html}{sufijo}</div>
<div class="dp-plan-meta">
<span>🗓️ Vigencia: <b>{meses if meses else '—'}</b> meses</span>
<span>📁 Hasta <b>{max_proy}</b> proyectos</span>{mensual_html}
</div>
</div>
"""


def _render_sidebar_marca():
    """Sidebar institucional en la pantalla pública (no autenticado)."""
    with st.sidebar:
        _md(f"""
<div class="dp-sb-brand">
<div class="dp-sb-logo">🛰️</div>
<div class="dp-sb-name">{BRAND['nombre']}</div>
<div class="dp-sb-tag">{BRAND['descriptor']}</div>
</div>
""")
        for icon, nombre in UNIDADES_NEGOCIO:
            _md(f'<div class="dp-sb-item">{icon}&nbsp;&nbsp;{nombre}</div>')
        _md(f"""
<div class="dp-sb-contact">
📩 {BRAND['email']}<br>
<span class="dp-sb-slogan">{BRAND['eslogan']}</span>
</div>
""")

# ======================================================================
# 4. INICIALIZACIÓN SAAS
# ======================================================================
def init_saas():
    """Inicializa BD y datos base (planes, admin). Idempotente."""
    init_db()
    seed_planes()
    seed_super_admin()


# ======================================================================
# 5. PANTALLA DE AUTENTICACIÓN
# ======================================================================
def render_auth_screen():

    # Configuración de página (eliminar si la app principal ya la define)
    try:
        st.set_page_config(
            page_title="Digital Progress · Evaluación Económica de Riego",
            page_icon="🛰️",
            layout="wide",
            initial_sidebar_state="expanded",
        )
    except Exception:
        pass

    init_saas()

    # ------------------------------------------------------------------
    # USUARIO YA AUTENTICADO
    # ------------------------------------------------------------------
    if st.session_state.get("saas_autenticado"):
        usuario = obtener_usuario_actual()
        if usuario:
            estado = verificar_suscripcion_activa(usuario.id)
            _md(_css(publico=False))

            with st.sidebar:
                _md(f"""
<div class="dp-user">
<div class="u-avatar">👤</div>
<div class="u-name">{usuario.nombre}</div>
<div class="u-email">{usuario.email}</div>
<span class="u-rol">{usuario.rol.value.replace('_', ' ')}</span>
</div>
""")
                _md(resumen_limites_sidebar(usuario.id))
                if st.button("🔒 Cerrar Sesión", use_container_width=True):
                    from ..auth import cerrar_sesion_streamlit
                    cerrar_sesion_streamlit()

            if not estado["activa"] and usuario.rol.value != "super_admin":
                st.error("⚠️ Tu suscripción ha vencido o no está activa. Contacta al administrador para renovar.")
                st.info("Los super administradores pueden gestionar suscripciones desde el panel de admin.")
                st.stop()

            return True
        else:
            for k in list(st.session_state.keys()):
                if k.startswith("saas_"):
                    del st.session_state[k]
            st.rerun()

    # ------------------------------------------------------------------
    # PANTALLA DE BIENVENIDA (no autenticado)
    # ------------------------------------------------------------------
    _md(_css(publico=True))
    _render_sidebar_marca()

    # ----- HERO DE MARCA -----
    _md(f"""
<div class="dp-hero">
<span class="dp-badge">🛰️ INGENIERÍA · GEOTECNOLOGÍA · INTELIGENCIA TERRITORIAL</span>
<h1>Digital <span>Progress</span></h1>
<p class="dp-desc">{BRAND['descriptor']}</p>
<p class="dp-sub">💧 Sistema de Evaluación Económica · Gestión integral de proyectos de riego en Pre-inversión (VIPFE)</p>
<p class="dp-slogan">{BRAND['eslogan']}</p>
</div>
""")

    # ----- UNIDADES DE NEGOCIO (pills) -----
    pills = "".join(
        f'<span class="dp-unit">{icon}&nbsp;&nbsp;{nombre}</span>'
        for icon, nombre in UNIDADES_NEGOCIO
    )
    _md(f'<div class="dp-units">{pills}</div>')

    col_brand, col_auth = st.columns([1.08, 1])
    prod_label = "proyecto" if TRIAL_MAX_PROYECTOS == 1 else "proyectos"

    # ----- COLUMNA IZQUIERDA: propuesta de valor -----
    with col_brand:
        _md(f"""
<div class="dp-card">
<div class="dp-card-title">🎯 ¿Por qué {BRAND['nombre']}?</div>
<p class="dp-claim">{BRAND['claim']}</p>
</div>
""")
        for icon, title, desc in FEATURES:
            _md(f"""
<div class="dp-feat">
<div class="dp-feat-ico">{icon}</div>
<div>
<div class="dp-feat-title">{title}</div>
<div class="dp-feat-desc">{desc}</div>
</div>
</div>
""")
        _md("""
<div class="dp-quote">
⚡ Reducimos hasta un <b>70%</b> el tiempo de elaboración de informes técnicos y mejoramos la trazabilidad de tu información.
</div>
""")
        _md(f"""
<div class="dp-card-cta">
<h4 style="margin:0 0 8px;">🚀 Prueba Gratis por {TRIAL_DIAS} Días</h4>
<p style="margin:0 0 6px;">Regístrate ahora y gestiona hasta <b>{TRIAL_MAX_PROYECTOS} {prod_label}</b> sin costo.</p>
<p style="margin:0;opacity:.88;font-size:.9rem;">✔ Sin tarjeta de crédito &nbsp;·&nbsp; ✔ Sin compromiso &nbsp;·&nbsp; ✔ Cancela cuando quieras</p>
</div>
""")

    # ----- COLUMNA DERECHA: tarjeta de autenticación -----
    with col_auth:
        try:
            auth_card = st.container(border=True)
        except Exception:  # Streamlit < 1.35 sin parámetro border
            auth_card = st.container()

        with auth_card:
            _md('<div class="dp-auth-head">🔐 <b>Acceso a la plataforma</b></div>')
            tab_login, tab_registro = st.tabs(["🔑 Iniciar Sesión", "🚀 Cuenta TRIAL"])

            # ---------------- LOGIN ----------------
            with tab_login:
                with st.form("form_login", clear_on_submit=False):
                    st.markdown("#### Accede a tu cuenta")
                    email = st.text_input("Correo electrónico", placeholder="tu@email.com")
                    password = st.text_input("Contraseña", type="password", placeholder="••••••••")
                    submit = st.form_submit_button("🔓 Ingresar", use_container_width=True, type="primary")

                    if submit:
                        if not email or not password:
                            st.error("Completa todos los campos.")
                        else:
                            usuario = autenticar_usuario(email, password)
                            if usuario:
                                if not usuario.activo:
                                    st.error("Tu cuenta está desactivada.")
                                else:
                                    iniciar_sesion_streamlit(usuario)
                                    st.success(f"¡Bienvenido, {usuario.nombre}!")
                                    st.rerun()
                            else:
                                st.error("Correo o contraseña incorrectos.")
                st.caption(f"¿Problemas para ingresar? Escríbenos a {BRAND['email']}")

            # ---------------- REGISTRO TRIAL ----------------
            with tab_registro:
                with st.form("form_registro", clear_on_submit=False):
                    st.markdown("#### Crea tu cuenta de prueba")
                    _md(f"""
<div class="dp-mini-note">
🎁 Incluye <b>{TRIAL_DIAS} días</b> de acceso completo con <b>{TRIAL_MAX_PROYECTOS} {prod_label}</b>. Sin tarjeta de crédito.
</div>
""")
                    nombre = st.text_input("Nombre completo", placeholder="Juan Pérez")
                    email = st.text_input("Correo electrónico", placeholder="juan@email.com", key="reg_email")
                    password = st.text_input("Contraseña", type="password", placeholder="Mínimo 6 caracteres", key="reg_pass")
                    password2 = st.text_input("Confirmar contraseña", type="password", placeholder="Repite tu contraseña")
                    acepta = st.checkbox("Acepto los términos y condiciones")
                    submit_reg = st.form_submit_button("🚀 Comenzar Prueba Gratis", use_container_width=True, type="primary")

                    if submit_reg:
                        if not all([nombre, email, password, password2]):
                            st.error("Completa todos los campos.")
                        elif password != password2:
                            st.error("Las contraseñas no coinciden.")
                        elif len(password) < 6:
                            st.error("La contraseña debe tener al menos 6 caracteres.")
                        elif not acepta:
                            st.error("Debes aceptar los términos y condiciones.")
                        else:
                            ok, msg = registrar_usuario_trial(email, nombre, password)
                            if ok:
                                st.success(msg)
                                usuario = autenticar_usuario(email, password)
                                if usuario:
                                    iniciar_sesion_streamlit(usuario)
                                    st.balloons()
                                    st.rerun()
                            else:
                                st.error(msg)

            _md('<div class="dp-trust">🛡️ Tus datos están protegidos · Información con trazabilidad y transparencia</div>')

    # ----- PLANES Y SUSCRIPCIONES -----
    _md("""
<div class="dp-section-head">
<div class="dp-section-kicker">PLANES Y SUSCRIPCIONES</div>
<h3 style="margin:6px 0 0;">Elige el plan que acompaña el crecimiento de tus proyectos</h3>
</div>
""")

    cols = st.columns(len(PLANES))
    total_planes = len(PLANES)
    destacado_idx = min(1, total_planes - 1) if total_planes >= 2 else None
    for idx, (slug, plan) in enumerate(PLANES.items()):
        with cols[idx]:
            _md(_plan_card(slug, plan, indice=idx, destacado=(idx == destacado_idx)))

    # ----- FOOTER DE MARCA -----
    _md(f"""
<div class="dp-footer">
<span class="brand">{BRAND['nombre']}</span> · {BRAND['descriptor']} ·
<a href="mailto:{BRAND['email']}" style="color:#1F6FB2;text-decoration:none;">{BRAND['email']}</a>
<span class="slogan">{BRAND['eslogan']}</span>
</div>
""")

    st.stop()

# ======================================================================
# 6. PANEL DE ADMINISTRACIÓN
# ======================================================================

def render_admin_panel():
    """Panel exclusivo para super_admin."""
    from ..auth import requerir_rol
    from sqlmodel import Session, select
    from ..models import Usuario, Suscripcion, EstadoSuscripcion
    from ..config import TRIAL_MAX_PROYECTOS

    requerir_rol(["super_admin"])
    _md(_css(publico=False))
    _md(f"""
<div class="dp-card">
<div class="dp-section-kicker">{BRAND['nombre'].upper()} · ADMINISTRACIÓN</div>
<h3 style="margin:6px 0 2px;">🛠️ Panel de Administración SaaS</h3>
<p style="margin:0;">Gestión de usuarios, suscripciones y planes.</p>
</div>
""")

    with Session(engine) as session:
        total_usuarios = len(session.exec(select(Usuario)).all())
        total_trial = len(session.exec(select(Suscripcion).where(Suscripcion.estado == EstadoSuscripcion.TRIAL)).all())
        total_activas = len(session.exec(select(Suscripcion).where(Suscripcion.estado == EstadoSuscripcion.ACTIVA)).all())
        total_vencidas = len(session.exec(select(Suscripcion).where(Suscripcion.estado == EstadoSuscripcion.VENCIDA)).all())

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("👥 Usuarios", total_usuarios)
        c2.metric("🚀 TRIAL", total_trial)
        c3.metric("✅ Activas", total_activas)
        c4.metric("⚠️ Vencidas", total_vencidas)

        st.divider()
        st.subheader("Usuarios Registrados")

        usuarios = session.exec(select(Usuario)).all()
        data = []
        for u in usuarios:
            sus = session.exec(select(Suscripcion).where(Suscripcion.usuario_id == u.id)).first()
            if sus and sus.plan:
                proyectos_str = f"{sus.proyectos_usados}/{sus.plan.max_proyectos}"
                plan_nombre = sus.plan.nombre
            elif sus:
                proyectos_str = f"{sus.proyectos_usados}/{TRIAL_MAX_PROYECTOS}"
                plan_nombre = "TRIAL"
            else:
                proyectos_str = "N/A"
                plan_nombre = "Sin Plan"

            data.append({
                "ID": u.id,
                "Nombre": u.nombre,
                "Email": u.email,
                "Rol": u.rol.value,
                "Activo": "✅" if u.activo else "❌",
                "Suscripción": sus.estado.value if sus else "N/A",
                "Plan": plan_nombre,
                "Proyectos": proyectos_str,
                "Vence": str(sus.fecha_fin) if sus else "N/A",
            })

        st.dataframe(data, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("🔧 Gestión Manual de Suscripción")

        col1, col2, col3 = st.columns(3)
        with col1:
            user_id_edit = st.number_input("ID Usuario", min_value=1, step=1)
        with col2:
            nuevo_estado = st.selectbox("Nuevo Estado", ["TRIAL", "ACTIVA", "VENCIDA"])
        with col3:
            dias_extra = st.number_input("Días de vigencia", min_value=1, value=30)

        if st.button("Actualizar Suscripción", type="primary"):
            sus = session.exec(select(Suscripcion).where(Suscripcion.usuario_id == user_id_edit)).first()
            if sus:
                sus.estado = EstadoSuscripcion(nuevo_estado)
                sus.fecha_fin = date.today() + timedelta(days=dias_extra)
                session.add(sus)
                session.commit()
                st.success("Suscripción actualizada.")
                st.rerun()
            else:
                st.error("Usuario no tiene suscripción.")
