"""
Modelos SQLModel para el módulo SaaS.
Tablas: Plan, Usuario, Suscripcion
Incluye extend_existing=True para evitar errores en hot-reload de Streamlit.
"""
from datetime import datetime, date
from typing import Optional
from sqlmodel import SQLModel, Field, Relationship, create_engine, Session, select
from sqlalchemy import Column, String, DateTime, Integer, Float, ForeignKey, Enum as SAEnum
import enum

from .config import DATABASE_URL

# ------------------------------------------------------------------
# Enums
# ------------------------------------------------------------------
class EstadoSuscripcion(str, enum.Enum):
    TRIAL = "TRIAL"
    ACTIVA = "ACTIVA"
    VENCIDA = "VENCIDA"

class RolUsuario(str, enum.Enum):
    super_admin = "super_admin"
    usuario = "usuario"

# ------------------------------------------------------------------
# Plan
# ------------------------------------------------------------------
class Plan(SQLModel, table=True):
    __tablename__ = "planes"
    __table_args__ = {"extend_existing": True}

    id: Optional[int] = Field(default=None, primary_key=True)
    slug: str = Field(index=True, unique=True)
    nombre: str
    duracion_meses: int
    max_proyectos: int
    precio_usd: float
    descripcion: Optional[str] = None
    activo: bool = Field(default=True)

    suscripciones: list["Suscripcion"] = Relationship(back_populates="plan")

# ------------------------------------------------------------------
# Usuario
# ------------------------------------------------------------------
class Usuario(SQLModel, table=True):
    __tablename__ = "usuarios"
    __table_args__ = {"extend_existing": True}

    id: Optional[int] = Field(default=None, primary_key=True)
    email: str = Field(index=True, unique=True)
    nombre: str
    password_hash: str
    rol: RolUsuario = Field(sa_column=Column(SAEnum(RolUsuario), default=RolUsuario.usuario))
    activo: bool = Field(default=True)
    creado_en: datetime = Field(default_factory=datetime.utcnow, sa_column=Column(DateTime))

    suscripcion: Optional["Suscripcion"] = Relationship(back_populates="usuario")

# ------------------------------------------------------------------
# Suscripcion
# ------------------------------------------------------------------
class Suscripcion(SQLModel, table=True):
    __tablename__ = "suscripciones"
    __table_args__ = {"extend_existing": True}

    id: Optional[int] = Field(default=None, primary_key=True)
    usuario_id: int = Field(foreign_key="usuarios.id", unique=True)
    plan_id: Optional[int] = Field(default=None, foreign_key="planes.id")
    estado: EstadoSuscripcion = Field(sa_column=Column(SAEnum(EstadoSuscripcion), default=EstadoSuscripcion.TRIAL))
    fecha_inicio: date
    fecha_fin: date
    proyectos_usados: int = Field(default=0)
    creado_en: datetime = Field(default_factory=datetime.utcnow, sa_column=Column(DateTime))

    usuario: Usuario = Relationship(back_populates="suscripcion")
    plan: Optional[Plan] = Relationship(back_populates="suscripciones")

# ------------------------------------------------------------------
# Engine y helpers
# ------------------------------------------------------------------
engine = create_engine(DATABASE_URL, echo=False)

def init_db():
    """Crea todas las tablas si no existen."""
    SQLModel.metadata.create_all(engine, checkfirst=True)

def get_session():
    """Context manager para sesiones de BD."""
    with Session(engine) as session:
        yield session

def seed_planes():
    """Inserta planes por defecto si la tabla está vacía."""
    from .config import PLANES
    with Session(engine) as session:
        for slug, data in PLANES.items():
            existing = session.exec(select(Plan).where(Plan.slug == slug)).first()
            if not existing:
                plan = Plan(
                    slug=slug,
                    nombre=data["nombre"],
                    duracion_meses=data["duracion_meses"],
                    max_proyectos=data["max_proyectos"],
                    precio_usd=data["precio_usd"],
                    descripcion=f"Incluye hasta {data['max_proyectos']} proyectos.",
                )
                session.add(plan)
        session.commit()

def seed_super_admin():
    """Crea super_admin por defecto si no existe."""
    from .auth import hash_password
    with Session(engine) as session:
        admin = session.exec(
            select(Usuario).where(Usuario.rol == RolUsuario.super_admin)
        ).first()
        if not admin:
            admin = Usuario(
                email="admin@riego.saas",
                nombre="Super Administrador",
                password_hash=hash_password("admin123"),
                rol=RolUsuario.super_admin,
                activo=True,
            )
            session.add(admin)
            session.commit()
