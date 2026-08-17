"""
Autenticación y autorización SIN JWT.
Usa únicamente st.session_state + consultas a BD.
"""
import hashlib
import secrets
from datetime import datetime, timedelta, date
from typing import Optional
from sqlmodel import Session, select

import streamlit as st

from .config import TRIAL_DIAS, TRIAL_MAX_PROYECTOS
from .models import engine, Usuario, Suscripcion, Plan, EstadoSuscripcion, RolUsuario

# ------------------------------------------------------------------
# Hash de contraseñas (SHA-256 + salt)
# ------------------------------------------------------------------
def hash_password(password: str) -> str:
    salt = secrets.token_hex(16)
    pwdhash = hashlib.sha256((password + salt).encode()).hexdigest()
    return f"{salt}${pwdhash}"

def verify_password(password: str, hashed: str) -> bool:
    try:
        salt, stored_hash = hashed.split("$")
        pwdhash = hashlib.sha256((password + salt).encode()).hexdigest()
        return pwdhash == stored_hash
    except Exception:
        return False

# ------------------------------------------------------------------
# Gestión de sesión en Streamlit (sin JWT)
# ------------------------------------------------------------------
def iniciar_sesion_streamlit(usuario: Usuario):
    """Guarda datos esenciales en st.session_state."""
    st.session_state["saas_user_id"] = usuario.id
    st.session_state["saas_user_email"] = usuario.email
    st.session_state["saas_user_nombre"] = usuario.nombre
    st.session_state["saas_user_rol"] = usuario.rol.value
    st.session_state["saas_autenticado"] = True
    st.session_state["saas_login_time"] = datetime.utcnow().isoformat()

def cerrar_sesion_streamlit():
    """Limpia la sesión SaaS."""
    for key in list(st.session_state.keys()):
        if key.startswith("saas_"):
            del st.session_state[key]
    st.rerun()

def obtener_usuario_actual() -> Optional[Usuario]:
    if not st.session_state.get("saas_autenticado"):
        return None
    user_id = st.session_state.get("saas_user_id")
    if not user_id:
        return None
    with Session(engine) as session:
        return session.get(Usuario, user_id)

def requerir_rol(roles_permitidos: list[str]):
    """Verificación de roles. Aborta con error si no cumple."""
    rol_actual = st.session_state.get("saas_user_rol", "")
    if rol_actual not in roles_permitidos:
        st.error("⛔ No tienes permisos para acceder a esta sección.")
        st.stop()

# ------------------------------------------------------------------
# Verificación de suscripción
# ------------------------------------------------------------------
def verificar_suscripcion_activa(usuario_id: int) -> dict:
    """
    Retorna estado detallado de la suscripción del usuario.
    SIEMPRE incluye las claves: plan_nombre, suscripcion_id, estado, etc.
    """
    with Session(engine) as session:
        suscripcion = session.exec(
            select(Suscripcion).where(Suscripcion.usuario_id == usuario_id)
        ).first()

        # Caso por defecto: SIN suscripción (todas las claves presentes)
        if not suscripcion:
            return {
                "activa": False,
                "estado": None,
                "dias_restantes": 0,
                "max_proyectos": 0,
                "proyectos_usados": 0,
                "puede_crear_proyecto": False,
                "mensaje": "No tienes una suscripción registrada.",
                "suscripcion_id": None,
                "plan_nombre": "Sin Plan",
            }

        hoy = date.today()
        dias_restantes = (suscripcion.fecha_fin - hoy).days

        # Auto-vencer si ya pasó la fecha
        if dias_restantes < 0 and suscripcion.estado != EstadoSuscripcion.VENCIDA:
            suscripcion.estado = EstadoSuscripcion.VENCIDA
            session.add(suscripcion)
            session.commit()

        max_proyectos = suscripcion.plan.max_proyectos if suscripcion.plan else TRIAL_MAX_PROYECTOS

        activa = suscripcion.estado in (EstadoSuscripcion.TRIAL, EstadoSuscripcion.ACTIVA)
        puede_crear = activa and (suscripcion.proyectos_usados < max_proyectos)

        mensaje = ""
        if suscripcion.estado == EstadoSuscripcion.TRIAL:
            mensaje = f"🚀 Modo TRIAL: {dias_restantes} días restantes."
        elif suscripcion.estado == EstadoSuscripcion.ACTIVA:
            mensaje = f"✅ Suscripción activa. Expira en {dias_restantes} días."
        elif suscripcion.estado == EstadoSuscripcion.VENCIDA:
            mensaje = "⚠️ Tu suscripción ha vencido. Renueva para continuar."

        return {
            "activa": activa,
            "estado": suscripcion.estado,
            "dias_restantes": max(dias_restantes, 0),
            "max_proyectos": max_proyectos,
            "proyectos_usados": suscripcion.proyectos_usados,
            "puede_crear_proyecto": puede_crear,
            "mensaje": mensaje,
            "suscripcion_id": suscripcion.id,
            "plan_nombre": suscripcion.plan.nombre if suscripcion.plan else "TRIAL"
        }

def registrar_usuario_trial(email: str, nombre: str, password: str) -> tuple[bool, str]:
    """Registra un nuevo usuario con suscripción TRIAL."""
    with Session(engine) as session:
        existente = session.exec(select(Usuario).where(Usuario.email == email)).first()
        if existente:
            return False, "El correo ya está registrado."

        usuario = Usuario(
            email=email,
            nombre=nombre,
            password_hash=hash_password(password),
            rol=RolUsuario.usuario,
            activo=True,
        )
        session.add(usuario)
        session.flush()

        hoy = date.today()
        trial = Suscripcion(
            usuario_id=usuario.id,
            plan_id=None,
            estado=EstadoSuscripcion.TRIAL,
            fecha_inicio=hoy,
            fecha_fin=hoy + timedelta(days=TRIAL_DIAS),
            proyectos_usados=0,
        )
        session.add(trial)
        session.commit()

        return True, "Registro exitoso. Bienvenido a tu período de prueba."

def autenticar_usuario(email: str, password: str) -> Optional[Usuario]:
    """Login tradicional. Retorna usuario si credenciales válidas."""
    with Session(engine) as session:
        usuario = session.exec(select(Usuario).where(Usuario.email == email)).first()
        if usuario and verify_password(password, usuario.password_hash):
            return usuario
        return None

def incrementar_proyectos_usados(usuario_id: int, cantidad: int = 1):
    with Session(engine) as session:
        suscripcion = session.exec(
            select(Suscripcion).where(Suscripcion.usuario_id == usuario_id)
        ).first()
        if suscripcion:
            suscripcion.proyectos_usados += cantidad
            session.add(suscripcion)
            session.commit()

def decrementar_proyectos_usados(usuario_id: int, cantidad: int = 1):
    with Session(engine) as session:
        suscripcion = session.exec(
            select(Suscripcion).where(Suscripcion.usuario_id == usuario_id)
        ).first()
        if suscripcion:
            suscripcion.proyectos_usados = max(0, suscripcion.proyectos_usados - cantidad)
            session.add(suscripcion)
            session.commit()
