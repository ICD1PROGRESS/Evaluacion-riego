"""
Lógica de negocio: límites por plan y suscripción.
Solo controla cantidad de proyectos (no usuarios).
"""
from sqlmodel import Session, select
from .models import engine, Suscripcion, EstadoSuscripcion
from .auth import verificar_suscripcion_activa

class LimiteExcedidoError(Exception):
    """Excepción lanzada cuando se supera el límite del plan."""
    pass

class SuscripcionVencidaError(Exception):
    """Excepción lanzada cuando la suscripción no está activa."""
    pass

def verificar_puede_crear_proyecto(usuario_id: int) -> dict:
    """
    Verifica si el usuario puede crear un nuevo proyecto.
    Lanza excepción si no está permitido.
    Retorna dict con info del estado.
    """
    estado = verificar_suscripcion_activa(usuario_id)

    if not estado["activa"]:
        raise SuscripcionVencidaError(estado["mensaje"])

    if not estado["puede_crear_proyecto"]:
        raise LimiteExcedidoError(
            f"Has alcanzado el límite de {estado['max_proyectos']} proyectos "
            f"de tu plan {estado['plan_nombre']}. "
            f"Elimina proyectos o actualiza tu suscripción."
        )

    return estado

def resumen_limites_sidebar(usuario_id: int) -> str:
    """Genera un texto HTML/Markdown para mostrar en el sidebar."""
    estado = verificar_suscripcion_activa(usuario_id)

    plan = estado.get("plan_nombre", "Sin Plan")
    usados = estado.get("proyectos_usados", 0)
    max_p = estado.get("max_proyectos", 0)
    dias = estado.get("dias_restantes", 0)
    estado_enum = estado.get("estado")
    mensaje = estado.get("mensaje", "Estado desconocido")

    barra = "🟩" * usados + "⬜" * max(0, max_p - usados)

    if estado_enum == EstadoSuscripcion.TRIAL:
        color = "#F59E0B"
        icono = "🚀"
    elif estado_enum == EstadoSuscripcion.ACTIVA:
        color = "#10B981"
        icono = "✅"
    elif estado_enum == EstadoSuscripcion.VENCIDA:
        color = "#EF4444"
        icono = "⚠️"
    else:
        # Sin suscripción o estado desconocido
        color = "#6B7280"
        icono = "❓"
        barra = ""

    return f"""
    <div style="padding:10px;border-radius:8px;background-color:{color}15;border-left:4px solid {color};">
        <b>{icono} {plan}</b><br>
        <small>{mensaje}</small><br>
        <small>Proyectos: {usados}/{max_p} {barra}</small>
    </div>
    """
