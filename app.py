import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import streamlit as st
from core.config import ConfiguracionProyecto
from core.project_manager import ProjectManager

from saas.ui.auth_screen import render_auth_screen, render_admin_panel   # SOLO PARA DESPLIEGUE
render_auth_screen()    # SOLO PARA DESPLIEGUE
# ============================================================
# CONFIGURACIÓN DE PÁGINA (común para toda la aplicación)
# ============================================================
st.set_page_config(
    page_title="Sistema de Evaluación Económica - Riego",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# ESTILOS CSS PERSONALIZADOS
# ============================================================
st.markdown("""
<style>
    .step-container {
        display: flex;
        justify-content: space-between;
        margin-bottom: 2rem;
        padding: 1rem;
        background-color: #f0f2f6;
        border-radius: 10px;
    }
    .step {
        text-align: center;
        padding: 0.5rem 1rem;
        border-radius: 20px;
        font-weight: bold;
    }
    .step-active {
        background-color: #1F4E78;
        color: white;
    }
    .step-completed {
        background-color: #70AD47;
        color: white;
    }
    .step-pending {
        background-color: #d1d5db;
        color: #6b7280;
    }
    /* Estilo para la navegación activa */
    [data-testid="stSidebarNav"] {
        background-color: #f8f9fa;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# GESTIÓN DE PROYECTOS (Sidebar común)
# ============================================================

def render_gestion_proyectos():
    """Sidebar: gestión de proyectos múltiples."""
    pm = ProjectManager()
    st.sidebar.caption("📁 Proyectos")

    # Proyecto activo
    activo = pm.get_rutas_activas()
    if activo:
        st.sidebar.success(f"✅ Activo: **{activo['nombre']}**")
    else:
        st.sidebar.warning("⚠️ Sin proyecto activo")

    # Crear nuevo
    with st.sidebar.expander("➕ Nuevo Proyecto"):
        nombre_nuevo = st.text_input("Nombre del proyecto", key="nuevo_proy")
        if st.button("Crear", key="btn_crear_proy"):
            if nombre_nuevo.strip():
                ok, slug = pm.crear_proyecto(nombre_nuevo.strip())
                if ok:
                    st.sidebar.success(f"Proyecto '{nombre_nuevo}' creado")
                    st.rerun()
                else:
                    st.sidebar.error(f"Error: {slug}")
            else:
                st.sidebar.error("Ingrese un nombre")

    # Cargar existente
    proyectos = pm.listar_proyectos()
    if proyectos:
        opciones = {f"{p['nombre']} ({p['slug']})": p['slug'] for p in proyectos}
        seleccion = st.sidebar.selectbox("Cargar proyecto", list(opciones.keys()), key="sel_proy")
        if st.sidebar.button("📂 Cargar seleccionado", key="btn_cargar"):
            slug = opciones[seleccion]
            ok, msg = pm.cargar_proyecto(slug)
            if ok:
                st.sidebar.success(f"Proyecto '{seleccion}' activado")
                st.rerun()
            else:
                st.sidebar.error(msg)

        # Duplicar
        with st.sidebar.expander("📋 Duplicar proyecto"):
            nombre_dup = st.text_input("Nombre para la copia", key="dup_proy")
            if st.button("Duplicar", key="btn_dup"):
                if nombre_dup.strip():
                    slug_origen = opciones[seleccion]
                    ok, slug_nuevo = pm.duplicar_proyecto(slug_origen, nombre_dup.strip())
                    if ok:
                        st.sidebar.success(f"Duplicado como '{nombre_dup}'")
                        st.rerun()
                    else:
                        st.sidebar.error(slug_nuevo)

    # Eliminar
    if proyectos and st.sidebar.checkbox("🗑️ Mostrar eliminar", key="chk_del"):
        st.sidebar.error("⚠️ Esta acción no se puede deshacer")
        if st.sidebar.button("Eliminar proyecto seleccionado", type="primary"):
            slug = opciones[seleccion]
            if pm.eliminar_proyecto(slug):
                st.sidebar.success("Proyecto eliminado")
                st.rerun()
            else:
                st.sidebar.error("Error al eliminar")

    # Ruta del proyecto activo (para referencia)
    if activo:
        st.sidebar.caption(f"📂 `{activo['carpeta']}`")

    return pm
# ============================================================
# DICCIONARIO OFICIAL DE TASAS POR DEPARTAMENTO
# ============================================================
TASAS_DEPARTAMENTO = {
    'Pando': 1.42, 'Oruro': 1.21, 'Cochabamba': 1.13, 'Beni': 1.03,
    'Santa Cruz': 1.33, 'La Paz': 0.91, 'Tarija': 0.83,
    'Chuquisaca': 0.35, 'Potosí': 0.33,
}

# ============================================================
# PÁGINA 1: CONFIGURACIÓN DEL PROYECTO (ÚNICA FUENTE DE VERDAD)
# ============================================================
def pagina_configuracion():
    """Paso 1: Configuración centralizada del proyecto. Todos los módulos consumen este JSON."""
    st.header("⚙️ Configuración del Proyecto")
    st.caption("Complete todos los parámetros. Los demás módulos leerán automáticamente esta configuración.")

    pm = ProjectManager()
    rutas = pm.get_rutas_activas()
    if not rutas:
        st.error("❌ No hay un proyecto activo. Cree o seleccione uno desde el panel lateral.")
        return

    ruta_config = rutas["config"]
    config = ConfiguracionProyecto.cargar(ruta_config)

    # --- CALLBACK: actualizar tasa al cambiar departamento ---
    def actualizar_tasa_depto():
        depto_sel = st.session_state.get('cfg_depto', '')
        if depto_sel in TASAS_DEPARTAMENTO:
            st.session_state.cfg_tasa_crec = TASAS_DEPARTAMENTO[depto_sel]
            
    # ============================================================
    # SECCIÓN 1: IDENTIFICACIÓN Y UBICACIÓN
    # ============================================================
    st.subheader("🏷️ Identificación y Ubicación")
    col_id1, col_id2 = st.columns(2)
    with col_id1:
        nombre = st.text_input("Nombre del Proyecto", config.nombre, key="cfg_nombre")
        codigo = st.text_input("Código del Proyecto", config.codigo, key="cfg_codigo")
    with col_id2:
        lista_deptos = [''] + sorted(TASAS_DEPARTAMENTO.keys())
        depto_index = lista_deptos.index(config.depto) if config.depto in lista_deptos else 0
        depto = st.selectbox(
            "Departamento",
            options=lista_deptos,
            index=depto_index,
            key="cfg_depto",
            on_change=actualizar_tasa_depto,
            help="Seleccione el departamento para cargar automáticamente la tasa de crecimiento poblacional oficial."
        )
        municipio = st.text_input("Municipio", config.municipio, key="cfg_municipio")

    # ============================================================
    # SECCIÓN 2: TEMPORAL Y DEMografía
    # ============================================================
    st.divider()
    col_t1, col_t2, col_t3 = st.columns(3)
    with col_t1:
        st.markdown("**📅 Temporal**")
        anio_inicio = st.number_input("Año Inicio", 2015, 2050, config.anio_inicio, key="cfg_anio_inicio")
        periodo = st.number_input("Período Diseño (años)", 1, 30, config.periodo_diseno, key="cfg_periodo")
        dur_inv = st.number_input("Duración Inversión (años)", 1, 5, config.duracion_inversion, key="cfg_dur_inv")
    with col_t2:
        st.markdown("**👥 Demografía**")
        pob_base = st.number_input("Población Base", 0, 50000, config.poblacion_base, key="cfg_pob_base")

        # Inicializar tasa en session_state si no existe
        tasa_default = config.tasa_crecimiento * 100
        if config.depto in TASAS_DEPARTAMENTO and 'cfg_tasa_crec' not in st.session_state:
            st.session_state.cfg_tasa_crec = TASAS_DEPARTAMENTO[config.depto]
        elif 'cfg_tasa_crec' not in st.session_state:
            st.session_state.cfg_tasa_crec = tasa_default

        tasa_crec = st.number_input(
            "Tasa Crec. Poblacional (%)", 0.0, 20.0,
            value=st.session_state.cfg_tasa_crec,
            format="%.4f", key="cfg_tasa_crec"   # ← key unificada con session_state
        )

        personas_fam = st.number_input("Pers. por Familia", 1, 20, max(1, config.personas_por_familia), key="cfg_personas_fam")
    with col_t3:
        st.markdown("**🌱 Superficie**")
        sup_actual = st.number_input("Actual con Riego (Ha)", 0.0, 5000.0, config.superficie_actual, key="cfg_sup_actual")
        sup_proy = st.number_input("Con Proyecto (Ha)", 0.0, 10000.0, config.superficie_proyecto, key="cfg_sup_proy")
        indice_impacto = st.slider("Índice Impacto Primer Año", 0.0, 1.0, config.indice_impacto, 0.05, key="cfg_indice_impacto")

    # ============================================================
    # SECCIÓN 3: TIPO DE CAMBIO, RPC Y TASAS
    # ============================================================
    st.divider()
    col_rpc1, col_rpc2, col_rpc3 = st.columns(3)
    with col_rpc1:
        st.markdown("**💱 Tipo de Cambio**")
        tipo_cambio = st.number_input(
            "Tipo de Cambio (Bs/USD)", 1.0, 20.0, config.tipo_cambio, 0.01,
            key="cfg_tipo_cambio",
            help="Tipo de cambio oficial aplicado a insumos transables en todos los módulos."
        )
        st.markdown("**🌾 Producción**")
        pct_transable = st.slider(
            "% Producción Transable", 0.0, 100.0,
            config.pct_produccion_transable * 100, 5.0, key="cfg_pct_transable"
        ) / 100
    with col_rpc2:
        st.markdown("**💰 Tasas VIPFE**")
        tasa_soc = st.number_input("Tasa Social Descuento (%)", 0.0, 50.0, config.tasa_social_descuento * 100, key="cfg_tasa_soc")
        tasa_priv = st.number_input("Tasa Privada Descuento (%)", 0.0, 50.0, config.tasa_privada_descuento * 100, key="cfg_tasa_priv")
    with col_rpc3:
        st.markdown("**🔄 RPC (Precios de Cuenta)**")
        rpc_div = st.number_input("RPC Divisa", 0.5, 3.0, config.rpc['divisa'], key="cfg_rpc_div")
        rpc_mcal = st.number_input("RPC MO Calificada", 0.1, 2.0, config.rpc['mo_calificada'], key="cfg_rpc_mcal")
        rpc_msemi = st.number_input("RPC MO Semi", 0.1, 2.0, config.rpc['mo_semicalificada'], key="cfg_rpc_msemi")
        rpc_mncu = st.number_input("RPC MO No Calif Urb", 0.1, 2.0, config.rpc['mo_no_calif_urbana'], key="cfg_rpc_mncu")
        rpc_mncr = st.number_input("RPC MO No Calif Rur", 0.1, 2.0, config.rpc['mo_no_calif_rural'], key="cfg_rpc_mncr")

    # ============================================================
    # SECCIÓN 4: PARÁMETROS FINANCIEROS ADICIONALES
    # ============================================================
    st.divider()
    with st.expander("💳 Parámetros Financieros Adicionales"):
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            valor_salvamento = st.number_input("Valor de Salvamento (último año)", value=config.valor_salvamento, step=100.0, key="cfg_salvamento")
            costo_financiero = st.number_input("Costos Financieros (Intereses) anual", value=config.costo_financiero, step=100.0, key="cfg_costo_fin")
            depreciacion_pos = st.number_input("Depreciación (+) anual", value=config.depreciacion_pos, step=100.0, key="cfg_dep_pos")
        with col_f2:
            depreciacion_neg = st.number_input("Depreciación (-) anual", value=config.depreciacion_neg, step=100.0, key="cfg_dep_neg")
            amortizacion = st.number_input("Amortización (+) anual", value=config.amortizacion, step=100.0, key="cfg_amort")
            prestamo = st.number_input("Préstamo (negativo)", value=config.prestamo, step=100.0, key="cfg_prestamo", help="Ingrese en negativo, ej. -100000")

    # ============================================================
    # GUARDAR CONFIGURACIÓN
    # ============================================================
    st.divider()
    col_guardar, _ = st.columns([1, 3])
    with col_guardar:
        if st.button("💾 Guardar Configuración", type="primary", use_container_width=True):
            config_actualizada = ConfiguracionProyecto(
                nombre=nombre,
                codigo=codigo,
                depto=depto,
                municipio=municipio,
                anio_inicio=int(anio_inicio),
                periodo_diseno=int(periodo),
                duracion_inversion=int(dur_inv),
                poblacion_base=int(pob_base),
                tasa_crecimiento=tasa_crec / 100,
                personas_por_familia=int(personas_fam),
                superficie_actual=sup_actual,
                superficie_proyecto=sup_proy,
                indice_impacto=indice_impacto,
                tasa_social_descuento=tasa_soc / 100,
                tasa_privada_descuento=tasa_priv / 100,
                tipo_cambio=tipo_cambio,
                pct_produccion_transable=pct_transable,
                valor_salvamento=valor_salvamento,
                costo_financiero=costo_financiero,
                depreciacion_pos=depreciacion_pos,
                depreciacion_neg=depreciacion_neg,
                amortizacion=amortizacion,
                prestamo=prestamo,
                rpc={
                    'divisa': rpc_div,
                    'mo_calificada': rpc_mcal,
                    'mo_semicalificada': rpc_msemi,
                    'mo_no_calif_urbana': rpc_mncu,
                    'mo_no_calif_rural': rpc_mncr,
                    'bienes_no_transables': 1.0
                }
            )
            config_actualizada.guardar(ruta_config)
            st.success(f"✅ Configuración guardada en '{rutas['nombre']}'")
            st.info("📌 Ahora puede navegar a los demás módulos. La evaluación leerá automáticamente estos parámetros.")

    # --- Vista previa de configuración guardada ---
    with st.expander("👁️ Vista previa de configuración actual"):
        st.json(config.to_dict())

# ============================================================
# PÁGINA 2: COSTOS DE PRODUCCIÓN
# ============================================================
def pagina_costos():

    # Usamos importlib para evitar conflictos de nombres
    import importlib
    import app_costos_v3
    importlib.reload(app_costos_v3)

    # Inicializar session_state del módulo de costos
    #if 'gestor_conceptos' not in st.session_state:    #esta linea se puede eliminar
    app_costos_v3.inicializar_session_state()    # Si hay if arriba, identar  espacio a la derecha

    # Renderizar las pestañas del módulo de costos
    st.title("🌾 Sistema de Costos de Produccion - Gestión de Conceptos y Cultivos")
    tabs = st.tabs([
        "📋 Conceptos de Costo",
        "🌾 Cultivos Referencia",
        "⚙️ Asignación y Cálculo",
        "🔄 Exportar a Proyecto"
    ])

    with tabs[0]:
        app_costos_v3.render_gestion_conceptos()
    with tabs[1]:
        app_costos_v3.render_gestion_referencia()
    with tabs[2]:
        app_costos_v3.render_asignacion_calculo()
    with tabs[3]:
        app_costos_v3.render_exportar_proyecto()

# ============================================================
# PÁGINA 3: INVERSIÓN
# ============================================================
def pagina_inversion():
    """Página de Inversión - integra app_inversion.py"""
    import importlib
    import app_inversion

    importlib.reload(app_inversion)

    # Inicializar session_state del módulo de inversión
    if 'obras' not in st.session_state:
        app_inversion.inicializar_session_state()

    # Crear y renderizar la UI de inversión
    ui = app_inversion.UIInversionCompleta()
    ui.render()

# ============================================================
# PÁGINA 4: EVALUACIÓN ECONÓMICA
# ============================================================
def pagina_evaluacion():
    """Página de Evaluación - integra app_evaluacion_mon.py"""
    import importlib
    import app_evaluacion

    importlib.reload(app_evaluacion)

    # Ejecutar el main del módulo de evaluación
    # Este módulo tiene toda su UI en main()
    app_evaluacion.main()

# ============================================================
# PÁGINA 5: REPORTES FINALES
# ============================================================
# ============================================================
# EJECUCIÓN PRINCIPAL - NAVEGACIÓN MULTIPAGE
# ============================================================

def main():
    # Renderizar gestión de proyectos en sidebar (común a todas las páginas)
    pm = render_gestion_proyectos()

    # Verificar que haya proyecto activo
    if not pm.existe_activo():
        st.title("💧 Sistema Integral de Evaluación Económica")
        st.warning("⚠️ No hay un proyecto activo. Cree o seleccione un proyecto desde el panel lateral.")

        # Mostrar solo la página de configuración como default
        pages = {
            "Proyecto": [
                st.Page(pagina_configuracion, title="⚙️ Configuración", default=True),
            ]
        }
        pg = st.navigation(pages)
        pg.run()
        return

    # Cargar configuración del proyecto activo para mostrar en header
    rutas = pm.get_rutas_activas()
    config = ConfiguracionProyecto.cargar(rutas["config"])

    # Header común en el interfaz principal
    st.caption(f"📁 Proyecto: {rutas['nombre']}")
    
    # Definir todas las páginas de navegación
    from saas.ui.auth_screen import render_admin_panel     # SOLO PARA DESPLIEGUE
    pages = {
        "Paginas de Evaluación Económica": [
            st.Page(pagina_configuracion, title="Configuración", icon="⚙️", default=True),
            st.Page(pagina_costos, title="Costos", icon="🌾"),
            st.Page(pagina_inversion, title="Inversión", icon="🏗️"),
            st.Page(pagina_evaluacion, title="Evaluación", icon="📊"),
        ],
        "Administración": [
            st.Page(render_admin_panel, title="Panel SaaS", icon="🛠️"),
        ]
    }
    # Crear navegación
    pg = st.navigation(pages, position="sidebar", expanded=True)
    pg.run()

if __name__ == "__main__":
    main()
