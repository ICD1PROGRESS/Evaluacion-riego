"""
Pantalla inicial de autenticación y suscripción.
Diseño atractivo con tabs: Login | Registro TRIAL.
Se muestra ANTES de cargar la aplicación principal.
"""
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

def init_saas():
    """Inicializa BD y datos base (planes, admin). Idempotente."""
    init_db()
    seed_planes()
    seed_super_admin()

def render_auth_screen():
    """
    Renderiza la pantalla de login/registro.
    Si el usuario ya está autenticado, muestra resumen y botón de cerrar sesión.
    Retorna True si la app principal debe continuar.
    """
    # Inicializar SaaS siempre
    init_saas()

    # Si ya está autenticado, mostrar info y permitir continuar
    if st.session_state.get("saas_autenticado"):
        usuario = obtener_usuario_actual()
        if usuario:
            estado = verificar_suscripcion_activa(usuario.id)

            # Sidebar info
            with st.sidebar:
                st.markdown("---")
                st.markdown(f"**👤 {usuario.nombre}**")
                st.markdown(f"<small>{usuario.email}</small>", unsafe_allow_html=True)
                st.markdown(resumen_limites_sidebar(usuario.id), unsafe_allow_html=True)
                if st.button("🔒 Cerrar Sesión", use_container_width=True):
                    from ..auth import cerrar_sesion_streamlit
                    cerrar_sesion_streamlit()

            # Si la suscripción venció o no existe, bloquear excepto admin
            if not estado["activa"] and usuario.rol.value != "super_admin":
                st.error("⚠️ Tu suscripción ha vencido o no está activa. Contacta al administrador para renovar.")
                st.info("Los super administradores pueden gestionar suscripciones desde el panel de admin.")
                st.stop()

            return True  # Continuar a la app principal
        else:
            # Sesión corrupta, limpiar
            for k in list(st.session_state.keys()):
                if k.startswith("saas_"):
                    del st.session_state[k]
            st.rerun()

    # ------------------------------------------------------------------
    # PANTALLA DE BIENVENIDA (no autenticado)
    # ------------------------------------------------------------------
    col_left, col_right = st.columns([1, 1])

    with col_left:
        st.markdown("# 💧 Sistema de Evaluación Económica")
        st.markdown("### Gestión integral de proyectos de riego")
        st.markdown("""
        <div style="margin-top:20px;">
            <p>✅ <b>Evaluación VIPFE</b> completa</p>
            <p>✅ <b>Costos de producción</b> por cultivo</p>
            <p>✅ <b>Análisis de inversión</b> detallado</p>
            <p>✅ <b>Reportes</b> exportables</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                    padding: 20px; border-radius: 15px; color: white;">
            <h4>🚀 Prueba Gratis por {TRIAL_DIAS} Días</h4>
            <p>Regístrate ahora y crea hasta <b>{TRIAL_MAX_PROYECTOS} proyecto</b> sin costo.</p>
            <p>Sin tarjeta de crédito. Sin compromiso.</p>
        </div>
        """, unsafe_allow_html=True)

    with col_right:
        tab_login, tab_registro = st.tabs(["🔑 Iniciar Sesión", "📝 Crear Cuenta TRIAL"])

        # ---------------- LOGIN ----------------
        with tab_login:
            with st.form("form_login", clear_on_submit=False):
                st.markdown("#### Accede a tu cuenta")
                email = st.text_input("Correo electrónico", placeholder="tu@email.com")
                password = st.text_input("Contraseña", type="password", placeholder="••••••••")
                submit = st.form_submit_button("Ingresar", use_container_width=True, type="primary")

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

            st.caption("¿Eres administrador? Usa las credenciales por defecto solo para demo.")

        # ---------------- REGISTRO TRIAL ----------------
        with tab_registro:
            with st.form("form_registro", clear_on_submit=False):
                st.markdown("#### Crea tu cuenta de prueba")
                st.info(f"Incluye {TRIAL_DIAS} días de acceso completo a {TRIAL_MAX_PROYECTOS} proyecto.")

                nombre = st.text_input("Nombre completo", placeholder="Juan Pérez")
                email = st.text_input("Correo electrónico", placeholder="juan@email.com", key="reg_email")
                password = st.text_input("Contraseña", type="password", placeholder="Mínimo 6 caracteres", key="reg_pass")
                password2 = st.text_input("Confirmar contraseña", type="password", placeholder="Repite tu contraseña")

                acepta = st.checkbox("Acepto los términos y condiciones")
                submit_reg = st.form_submit_button("Comenzar Prueba Gratis", use_container_width=True, type="primary")

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
                            # Auto-login
                            usuario = autenticar_usuario(email, password)
                            if usuario:
                                iniciar_sesion_streamlit(usuario)
                                st.balloons()
                                st.rerun()
                        else:
                            st.error(msg)

    # Footer
    st.markdown("---")
    cols = st.columns(len(PLANES))
    for idx, (slug, plan) in enumerate(PLANES.items()):
        with cols[idx]:
            st.markdown(f"""
            <div style="border:1px solid #e5e7eb; border-radius:10px; padding:15px; text-align:center;">
                <h4>{plan['nombre']}</h4>
                <p style="font-size:24px; color:#1F4E78;"><b>${plan['precio_usd']}</b></p>
                <p>Hasta <b>{plan['max_proyectos']}</b> proyectos</p>
                <p>Duración: <b>{plan['duracion_meses']} meses</b></p>
            </div>
            """, unsafe_allow_html=True)

    st.stop()  # Detener ejecución de la app principal hasta autenticarse

def render_admin_panel():
    """
    Panel exclusivo para super_admin.
    Permite ver usuarios, suscripciones y gestionar estados.
    """
    from ..auth import requerir_rol
    from sqlmodel import Session, select
    from ..models import Usuario, Suscripcion, EstadoSuscripcion
    from ..config import TRIAL_MAX_PROYECTOS

    requerir_rol(["super_admin"])

    st.header("🛠️ Panel de Administración SaaS")
    st.caption("Gestión de usuarios, suscripciones y planes.")

    with Session(engine) as session:
        # Estadísticas
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

        # Listado de usuarios
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

        # Gestión manual de suscripción
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
