"""
Configuración SaaS para la aplicación de Evaluación Económica de Riego.
Detecta automáticamente PostgreSQL (producción) vs SQLite (desarrollo).
Lee variables desde st.secrets (Streamlit Cloud) o variables de entorno (local).
"""
import os
import streamlit as st
from pathlib import Path
from dotenv import load_dotenv

# Cargar .env si existe (para desarrollo local)
load_dotenv()

def get_secret(key, default=None):
    try:
        # st.secrets solo está disponible cuando se ejecuta con streamlit run
        return st.secrets.get(key, default)
    except Exception:
        # Fallback a variables de entorno
        return os.getenv(key, default)

# --- Rutas base ---
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
SAAS_DATA_DIR = DATA_DIR / "saas"  # Datos exclusivos del módulo SaaS

DATA_DIR.mkdir(parents=True, exist_ok=True)
SAAS_DATA_DIR.mkdir(parents=True, exist_ok=True)

# --- Base de Datos (DETECCIÓN AUTOMÁTICA) ---
_env_db = get_secret("DATABASE_URL", "").strip()

if _env_db.startswith("postgresql"):
    DATABASE_URL = _env_db
    IS_PRODUCTION = True
    DB_TYPE = "postgresql"
elif _env_db.startswith("sqlite"):
    DATABASE_URL = _env_db
    IS_PRODUCTION = False
    DB_TYPE = "sqlite"
else:
    # Fallback SQLite local exclusivo para SaaS
    DATABASE_URL = f"sqlite:///{SAAS_DATA_DIR.as_posix()}/saas_auth.db"
    IS_PRODUCTION = False
    DB_TYPE = "sqlite"

# --- Base de Datos Global (opcional) ---
GLOBAL_DATABASE_URL = get_secret("GLOBAL_DATABASE_URL", "").strip()
if not GLOBAL_DATABASE_URL:
    # Si no se define, lo dejamos como None para que no se use
    GLOBAL_DATABASE_URL = None

# --- Seguridad ---
SECRET_KEY = get_secret("SECRET_KEY", "riego-saas-clave-secreta-por-defecto-2026")
JWT_ALGORITHM = get_secret("JWT_ALGORITHM", "HS256")

try:
    ACCESS_TOKEN_EXPIRE_MINUTES = int(get_secret("ACCESS_TOKEN_EXPIRE_MINUTES", "480"))
except ValueError:
    ACCESS_TOKEN_EXPIRE_MINUTES = 480

# --- Configuración de Suscripciones ---
TRIAL_DIAS = int(get_secret("TRIAL_DIAS", "10"))
TRIAL_MAX_PROYECTOS = int(get_secret("TRIAL_MAX_PROYECTOS", "1"))

# Planes predefinidos (se insertan vía seed si no existen)
PLANES = {
    "anual": {
        "nombre": "Plan Anual",
        "duracion_meses": 12,
        "max_proyectos": 10,
        "precio_usd": float(get_secret("PLAN_ANUAL_PRECIO", "99.0")),
    },
    "bianual": {
        "nombre": "Plan Bianual",
        "duracion_meses": 24,
        "max_proyectos": 35,
        "precio_usd": float(get_secret("PLAN_BIANUAL_PRECIO", "179.0")),
    },
}

# --- Streamlit / Servidor ---
STREAMLIT_SERVER_PORT = int(get_secret("STREAMLIT_SERVER_PORT", "8501"))
STREAMLIT_SERVER_ADDRESS = get_secret("STREAMLIT_SERVER_ADDRESS", "0.0.0.0")

# --- Modo Debug ---
DEBUG = get_secret("DEBUG", "false").lower() in ("true", "1", "yes", "on")

# --- Validación de seguridad en producción ---
if IS_PRODUCTION and SECRET_KEY == "riego-saas-clave-secreta-por-defecto-2026":
    raise RuntimeError(
        "ERROR DE SEGURIDAD: Debes definir SECRET_KEY en producción (secrets de Streamlit o variable de entorno). "
        "No uses la clave por defecto."
    )