# core/database.py
import json
import os
import sqlite3
import pandas as pd
from typing import List, Dict, Optional
from datetime import datetime

class ProyectoDB:

    def __init__(self, ruta_db: str):
        self.ruta_db = ruta_db
        os.makedirs(os.path.dirname(ruta_db), exist_ok=True)
        self._inicializar()

    def _get_conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.ruta_db)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = WAL")
        return conn

    def _inicializar(self):
        """Crea las tablas específicas del proyecto si no existen, y aplica migraciones."""
        with self._get_conn() as conn:
            # 1. Tablas existentes
            conn.execute("""
                CREATE TABLE IF NOT EXISTS proyecto_cultivos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL,
                    codigo TEXT DEFAULT '',
                    sup_sp_ha REAL DEFAULT 0,
                    sup_cp_ha REAL DEFAULT 0,
                    rend_sp REAL DEFAULT 0,
                    rend_cp REAL DEFAULT 0,
                    perd_sp_pct REAL DEFAULT 0,
                    perd_cp_pct REAL DEFAULT 0,
                    precio_bs_ton REAL DEFAULT 0,
                    costo_total_sp REAL DEFAULT 0,
                    costo_total_cp REAL DEFAULT 0,
                    bt_sp REAL DEFAULT 0, bnt_sp REAL DEFAULT 0,
                    monr_sp REAL DEFAULT 0, monu_sp REAL DEFAULT 0,
                    mos_sp REAL DEFAULT 0, moc_sp REAL DEFAULT 0,
                    bt_cp REAL DEFAULT 0, bnt_cp REAL DEFAULT 0,
                    monr_cp REAL DEFAULT 0, monu_cp REAL DEFAULT 0,
                    mos_cp REAL DEFAULT 0, moc_cp REAL DEFAULT 0,
                    ingreso_sp REAL DEFAULT 0,
                    ingreso_cp REAL DEFAULT 0,
                    fecha_agregado TEXT DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(nombre)
                )
            """)

            conn.execute("""
                CREATE TABLE IF NOT EXISTS obras_detalle (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo_item TEXT,
                    descripcion TEXT,
                    descripcion_norm TEXT,
                    unidad TEXT,
                    cantidad REAL DEFAULT 0,
                    precio_unitario REAL DEFAULT 0,
                    parcial_directo REAL DEFAULT 0,
                    tipo_hoja TEXT,
                    subcategoria TEXT,
                    tipo_rpc TEXT DEFAULT 'N',
                    fecha_clasificacion TEXT,
                    precio_unitario_final REAL DEFAULT 0,
                    parcial_real REAL DEFAULT 0,
                    factor_apu REAL DEFAULT 0,
                    indirectos_asignados REAL DEFAULT 0,
                    parcial REAL DEFAULT 0
                )
            """)

            conn.execute("""
                CREATE TABLE IF NOT EXISTS servicios (
                    categoria TEXT PRIMARY KEY,
                    descripcion TEXT,
                    bt REAL DEFAULT 0, bnt REAL DEFAULT 0,
                    moc REAL DEFAULT 0, mos REAL DEFAULT 0,
                    monu REAL DEFAULT 0, monr REAL DEFAULT 0,
                    detalle_adjunto TEXT DEFAULT ''
                )
            """)

            conn.execute("""
                CREATE TABLE IF NOT EXISTS apu_componentes (
                    apu TEXT PRIMARY KEY,
                    cantidad REAL DEFAULT 1.0,
                    cargas_sociales REAL DEFAULT 0,
                    herramientas REAL DEFAULT 0,
                    gastos_generales REAL DEFAULT 0,
                    utilidad REAL DEFAULT 0,
                    it REAL DEFAULT 0,
                    iva REAL DEFAULT 0,
                    precio_unitario_final REAL DEFAULT 0,
                    asignacion_json TEXT DEFAULT '{}'   -- ← NUEVO: guarda overrides de indirectos
                )
            """)
            
            #BLOQUE TEMPORAL
            cursor = conn.execute("PRAGMA table_info(apu_componentes)")
            cols_existentes = [row[1] for row in cursor.fetchall()]
            for col, tipo in [('cantidad', 'REAL DEFAULT 1.0'), ('asignacion_json', "TEXT DEFAULT '{}'")]:
                if col not in cols_existentes:
                    conn.execute(f"ALTER TABLE apu_componentes ADD COLUMN {col} {tipo}")
            #TEMPORAL HASTA AQUI

            # 2. Tabla de costos calculados (NUEVA ESTRUCTURA)
            # Se crea si no existe
            conn.execute("""
                CREATE TABLE IF NOT EXISTS proyecto_costos_calculados (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    plantilla_id INTEGER NOT NULL,
                    sup_sp_ha REAL DEFAULT 0,
                    sup_cp_ha REAL DEFAULT 0,
                    costo_conceptos_sp REAL DEFAULT 0,
                    costo_conceptos_cp REAL DEFAULT 0,
                    costo_adicionales_sp REAL DEFAULT 0,
                    costo_adicionales_cp REAL DEFAULT 0,
                    costo_total_sp REAL DEFAULT 0,
                    costo_total_cp REAL DEFAULT 0,
                    bt_sp REAL DEFAULT 0,
                    bnt_sp REAL DEFAULT 0,
                    monr_sp REAL DEFAULT 0,
                    monu_sp REAL DEFAULT 0,
                    mos_sp REAL DEFAULT 0,
                    moc_sp REAL DEFAULT 0,
                    bt_cp REAL DEFAULT 0,
                    bnt_cp REAL DEFAULT 0,
                    monr_cp REAL DEFAULT 0,
                    monu_cp REAL DEFAULT 0,
                    mos_cp REAL DEFAULT 0,
                    moc_cp REAL DEFAULT 0,
                    total_proy_sp REAL DEFAULT 0,
                    total_proy_cp REAL DEFAULT 0,
                    ingreso_sp REAL DEFAULT 0,
                    ingreso_cp REAL DEFAULT 0,
                    utilidad_sp REAL DEFAULT 0,
                    utilidad_cp REAL DEFAULT 0,
                    relacion_bc_sp REAL DEFAULT 0,
                    relacion_bc_cp REAL DEFAULT 0,
                    detalle_conceptos TEXT,
                    detalle_adicionales TEXT,
                    fecha_calculo TEXT DEFAULT CURRENT_TIMESTAMP                    
                )
            """)

            # 3. Migración: agregar columnas faltantes si la tabla ya existía
            cursor = conn.execute("PRAGMA table_info(proyecto_costos_calculados)")
            columnas_existentes = [row[1] for row in cursor.fetchall()]

            # TODAS las columnas que debe tener la tabla (base + nuevas)
            columnas_requeridas = {
                'sup_sp_ha': 'REAL DEFAULT 0',
                'sup_cp_ha': 'REAL DEFAULT 0',
                'costo_conceptos_sp': 'REAL DEFAULT 0',
                'costo_conceptos_cp': 'REAL DEFAULT 0',
                'costo_adicionales_sp': 'REAL DEFAULT 0',
                'costo_adicionales_cp': 'REAL DEFAULT 0',
                'costo_total_sp': 'REAL DEFAULT 0',
                'costo_total_cp': 'REAL DEFAULT 0',
                'bt_sp': 'REAL DEFAULT 0',
                'bnt_sp': 'REAL DEFAULT 0',
                'monr_sp': 'REAL DEFAULT 0',
                'monu_sp': 'REAL DEFAULT 0',
                'mos_sp': 'REAL DEFAULT 0',
                'moc_sp': 'REAL DEFAULT 0',
                'bt_cp': 'REAL DEFAULT 0',
                'bnt_cp': 'REAL DEFAULT 0',
                'monr_cp': 'REAL DEFAULT 0',
                'monu_cp': 'REAL DEFAULT 0',
                'mos_cp': 'REAL DEFAULT 0',
                'moc_cp': 'REAL DEFAULT 0',
                'total_proy_sp': 'REAL DEFAULT 0',
                'total_proy_cp': 'REAL DEFAULT 0',
                'ingreso_sp': 'REAL DEFAULT 0',
                'ingreso_cp': 'REAL DEFAULT 0',
                'utilidad_sp': 'REAL DEFAULT 0',
                'utilidad_cp': 'REAL DEFAULT 0',
                'relacion_bc_sp': 'REAL DEFAULT 0',
                'relacion_bc_cp': 'REAL DEFAULT 0',
                'detalle_conceptos': 'TEXT',
                'detalle_adicionales': 'TEXT',
            }

            for col, tipo in columnas_requeridas.items():
                if col not in columnas_existentes:
                    conn.execute(f"ALTER TABLE proyecto_costos_calculados ADD COLUMN {col} {tipo}")
                    print(f"Migración: columna '{col}' agregada a proyecto_costos_calculados")

            # Índices
            conn.execute("CREATE INDEX IF NOT EXISTS idx_obras_rpc ON obras_detalle(tipo_rpc)")
            conn.commit()
    # ============================================================
    # 1. GESTIÓN DE CULTIVOS EN PROYECTO (snapshots)
    # ============================================================   
    def agregar_cultivo_proyecto(self, datos: Dict) -> bool:
        """
        Agrega o actualiza un cultivo en el proyecto (snapshot).
        datos debe contener: nombre, sup_sp_ha, sup_cp_ha, rend_sp, rend_cp,
                             perd_sp_pct, perd_cp_pct, precio_bs_ton,
                             costo_total_sp, costo_total_cp,
                             bt_sp, bnt_sp, monr_sp, monu_sp, mos_sp, moc_sp,
                             bt_cp, bnt_cp, monr_cp, monu_cp, mos_sp, moc_cp,
                             ingreso_sp, ingreso_cp
        """
        try:
            with self._get_conn() as conn:
                conn.execute("""
                    INSERT OR REPLACE INTO proyecto_cultivos (
                        nombre, codigo, sup_sp_ha, sup_cp_ha,
                        rend_sp, rend_cp, perd_sp_pct, perd_cp_pct,
                        precio_bs_ton, costo_total_sp, costo_total_cp,
                        bt_sp, bnt_sp, monr_sp, monu_sp, mos_sp, moc_sp,
                        bt_cp, bnt_cp, monr_cp, monu_cp, mos_cp, moc_cp,
                        ingreso_sp, ingreso_cp, fecha_agregado
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    datos.get('nombre', ''),
                    datos.get('codigo', ''),
                    float(datos.get('sup_sp_ha', 0) or 0),
                    float(datos.get('sup_cp_ha', 0) or 0),
                    float(datos.get('rend_sp', 0) or 0),
                    float(datos.get('rend_cp', 0) or 0),
                    float(datos.get('perd_sp_pct', 0) or 0),
                    float(datos.get('perd_cp_pct', 0) or 0),
                    float(datos.get('precio_bs_ton', 0) or 0),
                    float(datos.get('costo_total_sp', 0) or 0),
                    float(datos.get('costo_total_cp', 0) or 0),
                    float(datos.get('bt_sp', 0) or 0),
                    float(datos.get('bnt_sp', 0) or 0),
                    float(datos.get('monr_sp', 0) or 0),
                    float(datos.get('monu_sp', 0) or 0),
                    float(datos.get('mos_sp', 0) or 0),
                    float(datos.get('moc_sp', 0) or 0),
                    float(datos.get('bt_cp', 0) or 0),
                    float(datos.get('bnt_cp', 0) or 0),
                    float(datos.get('monr_cp', 0) or 0),
                    float(datos.get('monu_cp', 0) or 0),
                    float(datos.get('mos_cp', 0) or 0),
                    float(datos.get('moc_cp', 0) or 0),
                    float(datos.get('ingreso_sp', 0) or 0),
                    float(datos.get('ingreso_cp', 0) or 0),
                    datetime.now().isoformat()
                ))
                conn.commit()
            return True
        except Exception as e:
            print(f"Error agregando cultivo al proyecto: {e}")
            return False

    def eliminar_cultivo_proyecto(self, nombre: str) -> bool:
        try:
            with self._get_conn() as conn:
                conn.execute("DELETE FROM proyecto_cultivos WHERE nombre = ?", (nombre,))
                conn.commit()
            return True
        except Exception as e:
            print(f"Error eliminando cultivo: {e}")
            return False

    def listar_cultivos_proyecto(self) -> pd.DataFrame:
        """
        Retorna un DataFrame con la estructura exacta de la hoja 'Cultivos' del Excel.
        Columnas: Nombre, Codigo, Sup_SP_Ha, Sup_CP_Ha, Rend_SP, Rend_CP,
                  Perd_SP_%, Perd_CP_%, Precio_Bs_Ton, CostoTotal_SP, CostoTotal_CP,
                  BT_SP, BNT_SP, MONR_SP, MONU_SP, MOS_SP, MOC_SP,
                  BT_CP, BNT_CP, MONR_CP, MONU_CP, MOS_CP, MOC_CP,
                  Ingreso_SP, Ingreso_CP, Fecha_Agregado
        """
        with self._get_conn() as conn:
            df = pd.read_sql("""
                SELECT
                    nombre as Nombre,
                    codigo as Codigo,
                    sup_sp_ha as `Sup_SP_Ha`,
                    sup_cp_ha as `Sup_CP_Ha`,
                    rend_sp as `Rend_SP`,
                    rend_cp as `Rend_CP`,
                    perd_sp_pct as `Perd_SP_%`,
                    perd_cp_pct as `Perd_CP_%`,
                    precio_bs_ton as `Precio_Bs_Ton`,
                    costo_total_sp as `CostoTotal_SP`,
                    costo_total_cp as `CostoTotal_CP`,
                    bt_sp as `BT_SP`,
                    bnt_sp as `BNT_SP`,
                    monr_sp as `MONR_SP`,
                    monu_sp as `MONU_SP`,
                    mos_sp as `MOS_SP`,
                    moc_sp as `MOC_SP`,
                    bt_cp as `BT_CP`,
                    bnt_cp as `BNT_CP`,
                    monr_cp as `MONR_CP`,
                    monu_cp as `MONU_CP`,
                    mos_cp as `MOS_CP`,
                    moc_cp as `MOC_CP`,
                    ingreso_sp as `Ingreso_SP`,
                    ingreso_cp as `Ingreso_CP`,
                    fecha_agregado as `Fecha_Agregado`
                FROM proyecto_cultivos
                ORDER BY nombre
            """, conn)
        # Rellenar nulos
        for col in df.columns:
            if col in ['Nombre', 'Codigo', 'Fecha_Agregado']:
                df[col] = df[col].fillna('').astype(str)
            else:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        return df

    # ============================================================
    # 2. GESTIÓN DE OBRAS DE INVERSIÓN
    # ============================================================
    def guardar_obras(self, df_obras: pd.DataFrame) -> bool:
        """Reemplaza todas las obras por el contenido del DataFrame."""
        try:
            with self._get_conn() as conn:
                conn.execute("DELETE FROM obras_detalle")
                if not df_obras.empty:
                    # Asegurar columnas
                    expected_cols = ['codigo_item', 'descripcion', 'descripcion_norm', 'unidad',
                                     'cantidad', 'precio_unitario', 'parcial_directo', 'tipo_hoja',
                                     'subcategoria', 'tipo_rpc', 'fecha_clasificacion',
                                     'precio_unitario_final', 'parcial_real', 'factor_apu',
                                     'indirectos_asignados', 'parcial']
                    for col in expected_cols:
                        if col not in df_obras.columns:
                            df_obras[col] = 0 if col in ['cantidad', 'precio_unitario', 'parcial_directo',
                                                         'precio_unitario_final', 'parcial_real',
                                                         'factor_apu', 'indirectos_asignados', 'parcial'] else ''
                    df_obras.to_sql('obras_detalle', conn, if_exists='append', index=False)
                conn.commit()
            return True
        except Exception as e:
            print(f"Error guardando obras: {e}")
            return False

    def listar_obras(self) -> pd.DataFrame:
        with self._get_conn() as conn:
            return pd.read_sql("SELECT * FROM obras_detalle ORDER BY subcategoria, codigo_item", conn)

    def obtener_resumen_obras(self) -> Dict[str, float]:
        """Retorna sumatorios por RPC para todas las obras."""
        df = self.listar_obras()
        if df.empty:
            return {'BT': 0, 'BNT': 0, 'MOC': 0, 'MOS': 0, 'MONU': 0, 'MONR': 0}
        # Usamos la columna 'parcial' que ya incluye factores APU
        col = 'parcial' if 'parcial' in df.columns and df['parcial'].sum() > 0 else 'parcial_directo'
        rpc_map = {
            'bt': 'BT', 'bnt': 'BNT', 'moc': 'MOC',
            'mos': 'MOS', 'monu': 'MONU', 'monr': 'MONR'
        }
        totales = {v: 0 for v in rpc_map.values()}
        for _, row in df.iterrows():
            tipo = str(row.get('tipo_rpc', '')).lower()
            if tipo in rpc_map:
                totales[rpc_map[tipo]] += float(row.get(col, 0) or 0)
        return totales

    # ============================================================
    # 3. GESTIÓN DE SERVICIOS
    # ============================================================
    def guardar_servicios(self, df_servicios: pd.DataFrame) -> bool:
        try:
            with self._get_conn() as conn:
                conn.execute("DELETE FROM servicios")
                if not df_servicios.empty:
                    df_servicios.to_sql('servicios', conn, if_exists='append', index=False)
                conn.commit()
            return True
        except Exception as e:
            print(f"Error guardando servicios: {e}")
            return False

    def listar_servicios(self) -> pd.DataFrame:
        with self._get_conn() as conn:
            return pd.read_sql("SELECT * FROM servicios", conn)

    def obtener_resumen_servicios(self) -> Dict[str, float]:
        df = self.listar_servicios()
        if df.empty:
            return {'BT': 0, 'BNT': 0, 'MOC': 0, 'MOS': 0, 'MONU': 0, 'MONR': 0}
        totales = {'BT': 0, 'BNT': 0, 'MOC': 0, 'MOS': 0, 'MONU': 0, 'MONR': 0}
        for _, row in df.iterrows():
            totales['BT'] += float(row.get('bt', 0) or 0)
            totales['BNT'] += float(row.get('bnt', 0) or 0)
            totales['MOC'] += float(row.get('moc', 0) or 0)
            totales['MOS'] += float(row.get('mos', 0) or 0)
            totales['MONU'] += float(row.get('monu', 0) or 0)
            totales['MONR'] += float(row.get('monr', 0) or 0)
        return totales

    # ============================================================
    # 4. GESTIÓN DE APU COMPONENTES
    # ============================================================
    def guardar_apu_componentes(self, df_apu: pd.DataFrame) -> bool:
        try:
            with self._get_conn() as conn:
                conn.execute("DELETE FROM apu_componentes")
                if not df_apu.empty:
                    df_apu.to_sql('apu_componentes', conn, if_exists='append', index=False)
                conn.commit()
            return True
        except Exception as e:
            print(f"Error guardando APU: {e}")
            return False

    def listar_apu_componentes(self) -> pd.DataFrame:
        with self._get_conn() as conn:
            return pd.read_sql("SELECT * FROM apu_componentes", conn)

    # ============================================================
    # 5. MÉTODOS DE AYUDA (resumen para el Excel)
    # ============================================================
    def obtener_resumen_inversion(self) -> pd.DataFrame:
        """
        Retorna un DataFrame con el resumen de inversión por categoría.
        Estructura para la hoja 'Inversion_Resumen' del Excel.
        Columnas: Categoría, BT, BNT, MOC, MOS, MONU, MONR, TOTAL
        """
        obras = self.obtener_resumen_obras()
        
        cat_nombres = {
            'ati': 'Asistencia Técnica Integral',
            'supervision': 'Supervisión de Obras',
            'om': 'Operación y Mantenimiento',
            'ambiental': 'Mitigación Ambiental'
        }
        
        rows = []
        # 1. Obras civiles
        rows.append({
            'Categoría': 'OBRAS CIVILES',
            'BT': obras['BT'],
            'BNT': obras['BNT'],
            'MOC': obras['MOC'],
            'MOS': obras['MOS'],
            'MONU': obras['MONU'],
            'MONR': obras['MONR'],
            'TOTAL': sum(obras.values())
        })
        
        # 2. Servicios individuales
        df_serv = self.listar_servicios()
        for _, row in df_serv.iterrows():
            cat_key = row['categoria']
            nombre = cat_nombres.get(cat_key, cat_key)
            rows.append({
                'Categoría': nombre,
                'BT': float(row.get('bt', 0)),
                'BNT': float(row.get('bnt', 0)),
                'MOC': float(row.get('moc', 0)),
                'MOS': float(row.get('mos', 0)),
                'MONU': float(row.get('monu', 0)),
                'MONR': float(row.get('monr', 0)),
                'TOTAL': sum(float(row.get(c, 0)) for c in ['bt', 'bnt', 'moc', 'mos', 'monu', 'monr'])
            })
        
        # 3. Total
        totals = {
            'BT': obras['BT'] + sum(float(row.get('bt', 0)) for _, row in df_serv.iterrows()),
            'BNT': obras['BNT'] + sum(float(row.get('bnt', 0)) for _, row in df_serv.iterrows()),
            'MOC': obras['MOC'] + sum(float(row.get('moc', 0)) for _, row in df_serv.iterrows()),
            'MOS': obras['MOS'] + sum(float(row.get('mos', 0)) for _, row in df_serv.iterrows()),
            'MONU': obras['MONU'] + sum(float(row.get('monu', 0)) for _, row in df_serv.iterrows()),
            'MONR': obras['MONR'] + sum(float(row.get('monr', 0)) for _, row in df_serv.iterrows()),
        }
        rows.append({
            'Categoría': 'TOTAL INVERSIÓN',
            'BT': totals['BT'],
            'BNT': totals['BNT'],
            'MOC': totals['MOC'],
            'MOS': totals['MOS'],
            'MONU': totals['MONU'],
            'MONR': totals['MONR'],
            'TOTAL': sum(totals.values())
        })
        return pd.DataFrame(rows)

    # ============================================================
    # 6. SECCION DE CALCULOS PERSONALIZADOS
    # ============================================================
    def guardar_costos_calculados_v2(self, datos: Dict) -> bool:
        """
        Guarda resultados completos del calculo incluyendo desglose RPC,
        totales por superficie, ingresos y utilidades.
        """
        try:
            with self._get_conn() as conn:
                conn.execute("""
                    INSERT OR REPLACE INTO proyecto_costos_calculados (
                        plantilla_id, sup_sp_ha, sup_cp_ha,
                        costo_conceptos_sp, costo_conceptos_cp,
                        costo_adicionales_sp, costo_adicionales_cp,
                        costo_total_sp, costo_total_cp,
                        bt_sp, bnt_sp, monr_sp, monu_sp, mos_sp, moc_sp,
                        bt_cp, bnt_cp, monr_cp, monu_cp, mos_cp, moc_cp,
                        total_proy_sp, total_proy_cp,
                        ingreso_sp, ingreso_cp, utilidad_sp, utilidad_cp,
                        relacion_bc_sp, relacion_bc_cp,
                        detalle_conceptos, detalle_adicionales,
                        fecha_calculo
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    datos.get('plantilla_id'),
                    float(datos.get('sup_sp_ha', 0) or 0),
                    float(datos.get('sup_cp_ha', 0) or 0),
                    float(datos.get('costo_conceptos_sp', 0) or 0),
                    float(datos.get('costo_conceptos_cp', 0) or 0),
                    float(datos.get('costo_adicionales_sp', 0) or 0),
                    float(datos.get('costo_adicionales_cp', 0) or 0),
                    float(datos.get('costo_total_sp', 0) or 0),
                    float(datos.get('costo_total_cp', 0) or 0),
                    float(datos.get('bt_sp', 0) or 0),
                    float(datos.get('bnt_sp', 0) or 0),
                    float(datos.get('monr_sp', 0) or 0),
                    float(datos.get('monu_sp', 0) or 0),
                    float(datos.get('mos_sp', 0) or 0),
                    float(datos.get('moc_sp', 0) or 0),
                    float(datos.get('bt_cp', 0) or 0),
                    float(datos.get('bnt_cp', 0) or 0),
                    float(datos.get('monr_cp', 0) or 0),
                    float(datos.get('monu_cp', 0) or 0),
                    float(datos.get('mos_cp', 0) or 0),
                    float(datos.get('moc_cp', 0) or 0),
                    float(datos.get('total_proy_sp', 0) or 0),
                    float(datos.get('total_proy_cp', 0) or 0),
                    float(datos.get('ingreso_sp', 0) or 0),
                    float(datos.get('ingreso_cp', 0) or 0),
                    float(datos.get('utilidad_sp', 0) or 0),
                    float(datos.get('utilidad_cp', 0) or 0),
                    float(datos.get('relacion_bc_sp', 0) or 0),
                    float(datos.get('relacion_bc_cp', 0) or 0),
                    json.dumps(datos.get('detalle_conceptos', []), ensure_ascii=False),
                    json.dumps(datos.get('detalle_adicionales', []), ensure_ascii=False),
                    datetime.now().isoformat()
                ))
                conn.commit()
            return True
        except Exception as e:
            print(f"Error guardando costos calculados v2: {e}")
            return False

    def obtener_costos_calculados(self, plantilla_id: int) -> Optional[Dict]:
        """
        Obtiene los costos calculados guardados para una plantilla.
        Retorna un diccionario con los mismos campos que guardar_costos_calculados.
        """
        with self._get_conn() as conn:
            row = conn.execute(
                "SELECT * FROM proyecto_costos_calculados WHERE plantilla_id = ? ORDER BY fecha_calculo DESC LIMIT 1",
                (plantilla_id,)
            ).fetchone()
            if not row:
                return None
            data = dict(row)
            # Parsear JSON
            if 'detalle_conceptos' in data and data['detalle_conceptos']:
                data['detalle_conceptos'] = json.loads(data['detalle_conceptos'])
            else:
                data['detalle_conceptos'] = []
            if 'detalle_adicionales' in data and data['detalle_adicionales']:
                data['detalle_adicionales'] = json.loads(data['detalle_adicionales'])
            else:
                data['detalle_adicionales'] = []
            return data
    # ============================================================
    # 7. LIMPIEZA
    # ============================================================
    def limpiar_datos_proyecto(self):
        """Vacía todas las tablas de datos del proyecto (no afecta catálogos globales)."""
        with self._get_conn() as conn:
            for tabla in ['proyecto_cultivos', 'obras_detalle', 'servicios', 'apu_componentes']:
                conn.execute(f"DELETE FROM {tabla}")
            conn.commit()