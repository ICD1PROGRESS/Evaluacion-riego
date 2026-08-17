# core/config.py
import json
import os
import math
from dataclasses import dataclass, field, asdict
from typing import Dict, List


@dataclass
class ConfiguracionProyecto:
    """
    Configuración única del proyecto, compartida por todos los módulos.
    """
    
    # Identificación
    nombre: str = ""
    codigo: str = ""
    depto: str = ""
    municipio: str = ""
    
    # Temporal
    anio_inicio: int = 2026
    periodo_diseno: int = 10
    duracion_inversion: int = 1
    
    # Demografía
    poblacion_base: int = 0
    tasa_crecimiento: float = 0.06
    personas_por_familia: int = 1
    
    # Superficie
    superficie_actual: float = 0.0
    superficie_proyecto: float = 0.0
    indice_impacto: float = 0.5
    
    # Financieras
    tasa_social_descuento: float = 0.0842
    tasa_privada_descuento: float = 0.0594
    
    # Producción
    pct_produccion_transable: float = 0.8
    
    # Tipo de cambio oficial (Bs/USD) — afecta a todos los módulos con insumos transables
    tipo_cambio: float = 10.12
    
    # RPC (Precios de Cuenta)
    rpc: Dict[str, float] = field(default_factory=lambda: {
        'divisa': 1.21,
        'mo_calificada': 0.43,
        'mo_semicalificada': 0.48,
        'mo_no_calif_urbana': 0.65,
        'mo_no_calif_rural': 0.63,
        'bienes_no_transables': 1.0
    })
    
    # Parámetros financieros adicionales
    valor_salvamento: float = 0.0
    costo_financiero: float = 0.0
    depreciacion_pos: float = 0.0
    depreciacion_neg: float = 0.0
    amortizacion: float = 0.0
    prestamo: float = 0.0
    
    # Nuevos parámetros para costos
    aplicar_gastos_administrativos: bool = True
    porcentaje_gastos_administrativos: float = 0.075
    
    # Factores de incremento para situación CON proyecto
    factor_incremento_riego: float = 1.2
    factor_incremento_mo: float = 1.1

    # Propiedades calculadas
    @property
    def area_incremental(self) -> float:
        return self.superficie_proyecto - self.superficie_actual
    
    @property
    def total_familias(self) -> int:
        return int(self.poblacion_base / self.personas_por_familia) if self.personas_por_familia > 0 else 0
    
    def generar_anios(self) -> List[int]:
        return list(range(self.anio_inicio, self.anio_inicio + self.periodo_diseno))
    
    def to_dict(self) -> dict:
        return asdict(self)
    
    @classmethod
    def from_dict(cls, d: dict) -> "ConfiguracionProyecto":
        campos_validos = {k: v for k, v in d.items() if k in cls.__dataclass_fields__}
        return cls(**campos_validos)
    
    def guardar(self, ruta: str = "config_proyecto.json"):
        with open(ruta, 'w', encoding='utf-8') as f:
            json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)
    
    @classmethod
    def cargar(cls, ruta: str = "config_proyecto.json") -> "ConfiguracionProyecto":
        if not os.path.exists(ruta):
            return cls()
        with open(ruta, 'r', encoding='utf-8') as f:
            return cls.from_dict(json.load(f))