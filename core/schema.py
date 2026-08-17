# core/schema.py
from typing import Dict, List, Tuple

# ============================================================
# SCHEMA DEL PUENTE: proyecto_activo.xlsx
# ============================================================

HOJAS_PUENTE = {
    "Cultivos": {
        "descripcion": "Datos de cultivos SP/CP con desglose RPC",
        "obligatoria": True,
        "columnas": [
            "Nombre", "Codigo", "Sup_SP_Ha", "Sup_CP_Ha",
            "Rend_SP", "Rend_CP", "Perd_SP_%", "Perd_CP_%",
            "Precio_Bs_Ton", "CostoTotal_SP", "CostoTotal_CP",
            "BT_SP", "BNT_SP", "MONR_SP", "MONU_SP", "MOS_SP", "MOC_SP",
            "BT_CP", "BNT_CP", "MONR_CP", "MONU_CP", "MOS_CP", "MOC_CP",
            "Ingreso_SP", "Ingreso_CP", "Fecha_Agregado"
        ],
        "tipos": {
            "Nombre": str, "Codigo": str,
            "Sup_SP_Ha": float, "Sup_CP_Ha": float,
            "Rend_SP": float, "Rend_CP": float,
            "Perd_SP_%": float, "Perd_CP_%": float,
            "Precio_Bs_Ton": float,
            "CostoTotal_SP": float, "CostoTotal_CP": float,
            "BT_SP": float, "BNT_SP": float, "MONR_SP": float,
            "MONU_SP": float, "MOS_SP": float, "MOC_SP": float,
            "BT_CP": float, "BNT_CP": float, "MONR_CP": float,
            "MONU_CP": float, "MOS_CP": float, "MOC_CP": float,
            "Ingreso_SP": float, "Ingreso_CP": float,
            "Fecha_Agregado": str
        }
    },
    "Inversion_Resumen": {
        "descripcion": "Resumen de inversión por categoría RPC",
        "obligatoria": True,
        "columnas": ["Categoría", "BT", "BNT", "MOC", "MOS", "MONU", "MONR", "TOTAL"],
        "tipos": {
            "Categoría": str,
            "BT": float, "BNT": float, "MOC": float,
            "MOS": float, "MONU": float, "MONR": float, "TOTAL": float
        },
        "filas_esperadas": [
            "OBRAS CIVILES",
            "Asistencia Técnica Integral",
            "Supervisión de Obras",
            "Operación y Mantenimiento",
            "Mitigación Ambiental",
            "TOTAL INVERSIÓN"
        ]
    },
    "Obras_Detalle": {
        "descripcion": "Ítems detallados de obras civiles clasificados",
        "obligatoria": False,
        "columnas": [
            "codigo_item", "descripcion", "unidad", "cantidad",
            "precio_unitario", "parcial", "tipo_hoja", "subcategoria",
            "tipo_rpc", "fecha_clasificacion"
        ]
    }
}

# Alias normalizados para lectura robusta
ALIAS_COLUMNAS = {
    'categoria': ['Categoría', 'CATEGORIA', 'categoria', 'Categoría'],
    'bt': ['BT', 'bt', 'Bienes Transables'],
    'bnt': ['BNT', 'bnt', 'Bienes No Transables', 'Materiales Locales', 'ML'],
    'moc': ['MOC', 'moc', 'Mano de Obra Calificada'],
    'mos': ['MOS', 'mos', 'Mano de Obra Semicalificada'],
    'monu': ['MONU', 'monu', 'M.O. No Calificada Urbana'],
    'monr': ['MONR', 'monr', 'M.O. No Calificada Rural', 'MOL', 'mol'],
}


def normalizar_columna(nombre_col: str, alias_dict: Dict[str, List[str]] = ALIAS_COLUMNAS) -> str:
    """Normaliza un nombre de columna al estándar interno (lowercase)."""
    nombre_upper = str(nombre_col).strip().upper()
    for estandar, aliases in alias_dict.items():
        if nombre_upper in [a.upper() for a in aliases]:
            return estandar
    return str(nombre_col).strip().lower()


def validar_hoja(df, nombre_hoja: str) -> Tuple[bool, List[str]]:
    """Valida que un DataFrame cumpla el schema de una hoja."""
    if nombre_hoja not in HOJAS_PUENTE:
        return False, [f"Hoja '{nombre_hoja}' no está definida en el schema"]
    
    schema = HOJAS_PUENTE[nombre_hoja]
    errores = []
    
    # Verificar columnas obligatorias
    cols_requeridas = set(schema["columnas"])
    cols_presentes = set(df.columns)
    
    faltantes = cols_requeridas - cols_presentes
    if faltantes:
        errores.append(f"Columnas faltantes en '{nombre_hoja}': {faltantes}")
    
    return len(errores) == 0, errores

# Validación definido por columna, especificando el tipo esperado y restricciones

VALIDATION_SCHEMA = {
    "Cultivos": {
        "columnas": {
            "nombre": {"tipo": str, "requerido": True},
            "codigo": {"tipo": str, "requerido": False},
            "sup_sp_ha": {"tipo": float, "min": 0.0},
            "sup_cp_ha": {"tipo": float, "min": 0.0},
            "rend_sp": {"tipo": float, "min": 0.0},
            "rend_cp": {"tipo": float, "min": 0.0},
            "perd_sp_%": {"tipo": float, "min": 0.0, "max": 100.0},
            "perd_cp_%": {"tipo": float, "min": 0.0, "max": 100.0},
            "precio_bs_ton": {"tipo": float, "min": 0.0},
            "costototal_sp": {"tipo": float, "min": 0.0},
            "costototal_cp": {"tipo": float, "min": 0.0},
            "bt_sp": {"tipo": float, "min": 0.0},
            "bnt_sp": {"tipo": float, "min": 0.0},
            "monr_sp": {"tipo": float, "min": 0.0},
            "monu_sp": {"tipo": float, "min": 0.0},
            "mos_sp": {"tipo": float, "min": 0.0},
            "moc_sp": {"tipo": float, "min": 0.0},
            "bt_cp": {"tipo": float, "min": 0.0},
            "bnt_cp": {"tipo": float, "min": 0.0},
            "monr_cp": {"tipo": float, "min": 0.0},
            "monu_cp": {"tipo": float, "min": 0.0},
            "mos_cp": {"tipo": float, "min": 0.0},
            "moc_cp": {"tipo": float, "min": 0.0},
            "ingreso_sp": {"tipo": float, "min": 0.0},
            "ingreso_cp": {"tipo": float, "min": 0.0},
            "fecha_agregado": {"tipo": str, "requerido": False}
        }
    },
    "Inversion_Resumen": {
        "columnas": {
            "Categoría": {"tipo": str, "requerido": True},
            "BT": {"tipo": float, "min": 0.0},
            "BNT": {"tipo": float, "min": 0.0},
            "MOC": {"tipo": float, "min": 0.0},
            "MOS": {"tipo": float, "min": 0.0},
            "MONU": {"tipo": float, "min": 0.0},
            "MONR": {"tipo": float, "min": 0.0},
            "TOTAL": {"tipo": float, "min": 0.0}
        }
    },
    # Agregar otras hojas según sea necesario
}