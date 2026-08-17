import os
import sqlite3
import pandas as pd
from typing import List, Dict, Optional, Any
from datetime import datetime

from dotenv import load_dotenv
load_dotenv()
# ============================================================
# DETECCIÓN AUTOMÁTICA DE BASE DE DATOS
# ============================================================
_env_global_db = os.getenv("GLOBAL_DATABASE_URL", "").strip()
USE_POSTGRES = _env_global_db.startswith("postgresql")

if USE_POSTGRES:
    import psycopg2
    from psycopg2.extras import RealDictCursor
    print(f"[GlobalDB] Modo POSTGRESQL activo: {_env_global_db.split('@')[1].split('/')[0]}")
else:
    print("[GlobalDB] Modo SQLITE activo (local)")

# ============================================================
# HELPERS DE SQL (placeholders y tipos)
# ============================================================
PH = "%s" if USE_POSTGRES else "?"

# PostgreSQL: BOOLEAN nativo. SQLite: INTEGER 0/1
BOOL_TRUE = "TRUE" if USE_POSTGRES else "1"
BOOL_FALSE = "FALSE" if USE_POSTGRES else "0"

def to_bool(val) -> bool:
    """Normaliza booleanos entre SQLite (0/1) y PostgreSQL (True/False)."""
    if val is None:
        return False
    if isinstance(val, bool):
        return val
    return bool(int(val))

class GlobalDB:
    """
    Capa de acceso a la base de datos global.
    - Si GLOBAL_DATABASE_URL empieza con postgresql → usa Neon/Supabase (PostgreSQL)
    - Si no → usa SQLite local (data/global.db)
    """

    def __init__(self, db_path: str):
        self.db_path = db_path
        self.pg_url = _env_global_db if USE_POSTGRES else None
        
        if not USE_POSTGRES:
            os.makedirs(os.path.dirname(db_path), exist_ok=True)
        self._init_tables_if_not_exists()

    def _get_conn(self):
        """Retorna una conexión: psycopg2 (PostgreSQL) o sqlite3 (SQLite)."""
        if USE_POSTGRES:
            conn = psycopg2.connect(self.pg_url)
            return conn
        else:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA foreign_keys = ON")
            return conn

    def _cursor(self, conn):
        """Retorna un cursor adecuado: RealDictCursor para PostgreSQL, normal para SQLite."""
        if USE_POSTGRES:
            return conn.cursor(cursor_factory=RealDictCursor)
        return conn.cursor()

    def _row_to_dict(self, row) -> Optional[Dict]:
        """Normaliza una fila a diccionario, independiente del motor."""
        if row is None:
            return None
        if USE_POSTGRES:
            return dict(row)
        return dict(row)

    def _execute(self, conn, sql: str, params: tuple = ()):
        """Ejecuta SQL adaptando placeholders automáticamente."""
        # En SQLite usamos ?, en PostgreSQL %s
        # Pero como el código fuente usa ?, hacemos replace si es PostgreSQL
        if USE_POSTGRES:
            sql_pg = sql.replace("?", "%s")
            cur = conn.cursor()
            cur.execute(sql_pg, params)
            return cur
        else:
            cur = conn.cursor()
            cur.execute(sql, params)
            return cur

    def _executemany(self, conn, sql: str, params_list: List[tuple]):
        if USE_POSTGRES:
            sql_pg = sql.replace("?", "%s")
            cur = conn.cursor()
            cur.executemany(sql_pg, params_list)
            return cur
        else:
            cur = conn.cursor()
            cur.executemany(sql, params_list)
            return cur

    def _commit_and_close(self, conn):
        if USE_POSTGRES:
            conn.commit()
            conn.close()
        else:
            conn.commit()
            conn.close()

    def _init_tables_if_not_exists(self):
        """Crea las tablas si no existen (DDL adaptado a PostgreSQL o SQLite)."""
        conn = self._get_conn()
        cur = self._cursor(conn)
        
        if USE_POSTGRES:
            # PostgreSQL DDL
            cur.execute("""
                CREATE TABLE IF NOT EXISTS unidad_medida (
                    id SERIAL PRIMARY KEY,
                    nombre TEXT NOT NULL UNIQUE,
                    abreviatura TEXT
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS departamento (
                    id SERIAL PRIMARY KEY,
                    nombre TEXT NOT NULL UNIQUE
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS municipio (
                    id SERIAL PRIMARY KEY,
                    nombre TEXT NOT NULL,
                    departamento_id INTEGER NOT NULL REFERENCES departamento(id),
                    UNIQUE(nombre, departamento_id)
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS cultivo (
                    id SERIAL PRIMARY KEY,
                    nombre TEXT NOT NULL UNIQUE,
                    codigo TEXT DEFAULT '',
                    familia TEXT DEFAULT ''
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS variedad (
                    id SERIAL PRIMARY KEY,
                    nombre TEXT NOT NULL,
                    cultivo_id INTEGER NOT NULL REFERENCES cultivo(id),
                    UNIQUE(nombre, cultivo_id)
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS nivel_tecnologico (
                    id SERIAL PRIMARY KEY,
                    nombre TEXT NOT NULL UNIQUE
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS concepto (
                    id SERIAL PRIMARY KEY,
                    nombre TEXT NOT NULL UNIQUE,
                    unidad_id INTEGER NOT NULL REFERENCES unidad_medida(id),
                    categoria TEXT DEFAULT '',
                    clase_rpc TEXT DEFAULT 'bnt',
                    precio_referencial REAL DEFAULT 0,
                    es_transable BOOLEAN DEFAULT FALSE,
                    observaciones TEXT DEFAULT '',
                    lugar TEXT DEFAULT '',
                    cultivo_asociado TEXT DEFAULT '',
                    fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS plantilla_costos (
                    id SERIAL PRIMARY KEY,
                    cultivo_id INTEGER NOT NULL REFERENCES cultivo(id),
                    variedad_id INTEGER REFERENCES variedad(id),
                    nivel_tecnologico_id INTEGER NOT NULL REFERENCES nivel_tecnologico(id),
                    departamento_id INTEGER REFERENCES departamento(id),
                    municipio_id INTEGER REFERENCES municipio(id),
                    campania TEXT DEFAULT '',
                    activo BOOLEAN DEFAULT TRUE,
                    rendimiento_sp REAL DEFAULT 0,
                    rendimiento_cp REAL DEFAULT 0,
                    rendimiento_kg_ha REAL DEFAULT 0,
                    precio_ref_bs_ton REAL DEFAULT 0,
                    perdidas_sp_pct REAL DEFAULT 0,
                    perdidas_cp_pct REAL DEFAULT 0,
                    fuente TEXT DEFAULT '',
                    fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS plantilla_concepto (
                    id SERIAL PRIMARY KEY,
                    plantilla_id INTEGER NOT NULL REFERENCES plantilla_costos(id) ON DELETE CASCADE,
                    concepto_id INTEGER NOT NULL REFERENCES concepto(id),
                    cantidad REAL DEFAULT 0,
                    cantidad_cp REAL,
                    precio_override REAL DEFAULT 0,
                    observaciones TEXT DEFAULT '',
                    orden INTEGER DEFAULT 0,
                    factor_incremento_sp REAL DEFAULT 1.0,
                    factor_incremento_cp REAL DEFAULT 1.0,
                    UNIQUE(plantilla_id, concepto_id)
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS plantilla_gasto_general (
                    id SERIAL PRIMARY KEY,
                    plantilla_id INTEGER NOT NULL REFERENCES plantilla_costos(id) ON DELETE CASCADE,
                    tipo TEXT NOT NULL,
                    porcentaje REAL DEFAULT 0,
                    monto_fijo REAL DEFAULT 0,
                    base_calculo TEXT DEFAULT 'directo',
                    descripcion TEXT DEFAULT ''
                )
            """)
            # Índices
            cur.execute("CREATE INDEX IF NOT EXISTS idx_concepto_nombre ON concepto(nombre)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_plantilla_cultivo ON plantilla_costos(cultivo_id)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_plantilla_concepto_plantilla ON plantilla_concepto(plantilla_id)")
        else:
            # SQLite DDL (código original)
            conn.executescript("""
                PRAGMA foreign_keys = ON;

                CREATE TABLE IF NOT EXISTS unidad_medida (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL UNIQUE,
                    abreviatura TEXT
                );

                CREATE TABLE IF NOT EXISTS departamento (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL UNIQUE
                );

                CREATE TABLE IF NOT EXISTS municipio (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL,
                    departamento_id INTEGER NOT NULL,
                    UNIQUE(nombre, departamento_id),
                    FOREIGN KEY (departamento_id) REFERENCES departamento(id)
                );

                CREATE TABLE IF NOT EXISTS cultivo (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL UNIQUE,
                    codigo TEXT DEFAULT '',
                    familia TEXT DEFAULT ''
                );

                CREATE TABLE IF NOT EXISTS variedad (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL,
                    cultivo_id INTEGER NOT NULL,
                    UNIQUE(nombre, cultivo_id),
                    FOREIGN KEY (cultivo_id) REFERENCES cultivo(id)
                );

                CREATE TABLE IF NOT EXISTS nivel_tecnologico (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL UNIQUE
                );

                CREATE TABLE IF NOT EXISTS concepto (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL UNIQUE,
                    unidad_id INTEGER NOT NULL,
                    categoria TEXT DEFAULT '',
                    clase_rpc TEXT DEFAULT 'bnt',
                    precio_referencial REAL DEFAULT 0,
                    es_transable BOOLEAN DEFAULT 0,
                    observaciones TEXT DEFAULT '',
                    lugar TEXT DEFAULT '',
                    cultivo_asociado TEXT DEFAULT '',
                    fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (unidad_id) REFERENCES unidad_medida(id)
                );

                CREATE TABLE IF NOT EXISTS plantilla_costos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    cultivo_id INTEGER NOT NULL,
                    variedad_id INTEGER,
                    nivel_tecnologico_id INTEGER NOT NULL,
                    departamento_id INTEGER,
                    municipio_id INTEGER,
                    campania TEXT DEFAULT '',
                    activo BOOLEAN DEFAULT 1,
                    rendimiento_kg_ha REAL DEFAULT 0,
                    precio_ref_bs_ton REAL DEFAULT 0,
                    perdidas_sp_pct REAL DEFAULT 0,
                    perdidas_cp_pct REAL DEFAULT 0,
                    fuente TEXT DEFAULT '',
                    fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (cultivo_id) REFERENCES cultivo(id),
                    FOREIGN KEY (variedad_id) REFERENCES variedad(id),
                    FOREIGN KEY (nivel_tecnologico_id) REFERENCES nivel_tecnologico(id),
                    FOREIGN KEY (departamento_id) REFERENCES departamento(id),
                    FOREIGN KEY (municipio_id) REFERENCES municipio(id)
                );

                CREATE TABLE IF NOT EXISTS plantilla_concepto (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    plantilla_id INTEGER NOT NULL,
                    concepto_id INTEGER NOT NULL,
                    cantidad REAL DEFAULT 0,
                    precio_override REAL DEFAULT 0,
                    observaciones TEXT DEFAULT '',
                    orden INTEGER DEFAULT 0,
                    factor_incremento_sp REAL DEFAULT 1.0,
                    factor_incremento_cp REAL DEFAULT 1.0,
                    UNIQUE(plantilla_id, concepto_id),
                    FOREIGN KEY (plantilla_id) REFERENCES plantilla_costos(id) ON DELETE CASCADE,
                    FOREIGN KEY (concepto_id) REFERENCES concepto(id)
                );

                CREATE TABLE IF NOT EXISTS plantilla_gasto_general (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    plantilla_id INTEGER NOT NULL,
                    tipo TEXT NOT NULL,
                    porcentaje REAL DEFAULT 0,
                    monto_fijo REAL DEFAULT 0,
                    base_calculo TEXT DEFAULT 'directo',
                    descripcion TEXT DEFAULT '',
                    FOREIGN KEY (plantilla_id) REFERENCES plantilla_costos(id) ON DELETE CASCADE
                );

                CREATE INDEX IF NOT EXISTS idx_concepto_nombre ON concepto(nombre);
                CREATE INDEX IF NOT EXISTS idx_plantilla_cultivo ON plantilla_costos(cultivo_id);
                CREATE INDEX IF NOT EXISTS idx_plantilla_concepto_plantilla ON plantilla_concepto(plantilla_id);
            """)
            # Verificar si la tabla plantilla_costos ya existe y tiene las columnas nuevas
            cursor = conn.execute("PRAGMA table_info(plantilla_costos)")
            columnas = [row[1] for row in cursor.fetchall()]
            if 'rendimiento_sp' not in columnas:
                conn.execute("ALTER TABLE plantilla_costos ADD COLUMN rendimiento_sp REAL DEFAULT 0")
            if 'rendimiento_cp' not in columnas:
                conn.execute("ALTER TABLE plantilla_costos ADD COLUMN rendimiento_cp REAL DEFAULT 0")

            # Verificar si plantilla_concepto ya tiene cantidad_cp
            cursor = conn.execute("PRAGMA table_info(plantilla_concepto)")
            columnas_pc = [row[1] for row in cursor.fetchall()]
            if 'cantidad_cp' not in columnas_pc:
                conn.execute("ALTER TABLE plantilla_concepto ADD COLUMN cantidad_cp REAL")
        
        self._commit_and_close(conn)

    # ============================================================
    # MÉTODOS AUXILIARES Y DE CATÁLOGO
    # ============================================================

    def _get_id(self, tabla: str, campo: str, valor: str) -> Optional[int]:
        if not valor:
            return None
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute(f"SELECT id FROM {tabla} WHERE {campo} = ?", (valor.strip(),))
        row = cur.fetchone()
        conn.close()
        return row['id'] if row else None

    def obtener_cultivos(self) -> List[Dict]:
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT id, nombre FROM cultivo ORDER BY nombre")
        rows = cur.fetchall()
        conn.close()
        return [self._row_to_dict(row) for row in rows]

    def obtener_variedades(self, cultivo_id: int = None) -> List[Dict]:
        conn = self._get_conn()
        cur = self._cursor(conn)
        if cultivo_id:
            cur.execute("SELECT id, nombre, cultivo_id FROM variedad WHERE cultivo_id = ? ORDER BY nombre", (cultivo_id,))
        else:
            cur.execute("SELECT id, nombre, cultivo_id FROM variedad ORDER BY nombre")
        rows = cur.fetchall()
        conn.close()
        return [self._row_to_dict(row) for row in rows]

    def obtener_niveles_tecnologicos(self) -> List[Dict]:
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT id, nombre FROM nivel_tecnologico ORDER BY nombre")
        rows = cur.fetchall()
        conn.close()
        return [self._row_to_dict(row) for row in rows]

    def obtener_departamentos(self) -> List[Dict]:
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT id, nombre FROM departamento ORDER BY nombre")
        rows = cur.fetchall()
        conn.close()
        return [self._row_to_dict(row) for row in rows]

    def obtener_municipios(self, depto_id: int = None) -> List[Dict]:
        conn = self._get_conn()
        cur = self._cursor(conn)
        if depto_id:
            cur.execute("SELECT id, nombre, departamento_id FROM municipio WHERE departamento_id = ? ORDER BY nombre", (depto_id,))
        else:
            cur.execute("SELECT id, nombre, departamento_id FROM municipio ORDER BY nombre")
        rows = cur.fetchall()
        conn.close()
        return [self._row_to_dict(row) for row in rows]

    def obtener_unidades(self) -> List[Dict]:
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT id, nombre, abreviatura FROM unidad_medida ORDER BY nombre")
        rows = cur.fetchall()
        conn.close()
        return [self._row_to_dict(row) for row in rows]

    # ============================================================
    # GET-OR-CREATE PARA CATÁLOGOS MAESTROS
    # ============================================================

    def obtener_o_crear_cultivo(self, nombre: str) -> int:
        nombre = nombre.strip()
        if not nombre:
            raise ValueError("El nombre del cultivo es obligatorio")
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT id FROM cultivo WHERE nombre = ?", (nombre,))
        row = cur.fetchone()
        if row:
            conn.close()
            return row['id']
        cur.execute("INSERT INTO cultivo (nombre) VALUES (?) RETURNING id" if USE_POSTGRES else "INSERT INTO cultivo (nombre) VALUES (?)", (nombre,))
        if USE_POSTGRES:
            cid = cur.fetchone()['id']
        else:
            cid = cur.lastrowid
        conn.commit()
        conn.close()
        return cid

    def obtener_o_crear_variedad(self, nombre: str, cultivo_id: int) -> Optional[int]:
        nombre = nombre.strip()
        if not nombre or nombre.upper() in ("(NINGUNA)", "NINGUNA", ""):
            return None
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT id FROM variedad WHERE nombre = ? AND cultivo_id = ?", (nombre, cultivo_id))
        row = cur.fetchone()
        if row:
            conn.close()
            return row['id']
        cur.execute("INSERT INTO variedad (nombre, cultivo_id) VALUES (?, ?) RETURNING id" if USE_POSTGRES else "INSERT INTO variedad (nombre, cultivo_id) VALUES (?, ?)", (nombre, cultivo_id))
        if USE_POSTGRES:
            vid = cur.fetchone()['id']
        else:
            vid = cur.lastrowid
        conn.commit()
        conn.close()
        return vid

    def obtener_o_crear_departamento(self, nombre: str) -> int:
        nombre = nombre.strip()
        if not nombre:
            raise ValueError("El nombre del departamento es obligatorio")
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT id FROM departamento WHERE nombre = ?", (nombre,))
        row = cur.fetchone()
        if row:
            conn.close()
            return row['id']
        cur.execute("INSERT INTO departamento (nombre) VALUES (?) RETURNING id" if USE_POSTGRES else "INSERT INTO departamento (nombre) VALUES (?)", (nombre,))
        if USE_POSTGRES:
            did = cur.fetchone()['id']
        else:
            did = cur.lastrowid
        conn.commit()
        conn.close()
        return did

    def obtener_o_crear_municipio(self, nombre: str, departamento_id: int) -> Optional[int]:
        nombre = nombre.strip()
        if not nombre or nombre.upper() in ("(NINGUNO)", "NINGUNO", ""):
            return None
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT id FROM municipio WHERE nombre = ? AND departamento_id = ?", (nombre, departamento_id))
        row = cur.fetchone()
        if row:
            conn.close()
            return row['id']
        cur.execute("INSERT INTO municipio (nombre, departamento_id) VALUES (?, ?) RETURNING id" if USE_POSTGRES else "INSERT INTO municipio (nombre, departamento_id) VALUES (?, ?)", (nombre, departamento_id))
        if USE_POSTGRES:
            mid = cur.fetchone()['id']
        else:
            mid = cur.lastrowid
        conn.commit()
        conn.close()
        return mid

    def obtener_o_crear_nivel_tecnologico(self, nombre: str) -> int:
        nombre = nombre.strip()
        if not nombre:
            raise ValueError("El nivel tecnológico es obligatorio")
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT id FROM nivel_tecnologico WHERE nombre = ?", (nombre,))
        row = cur.fetchone()
        if row:
            conn.close()
            return row['id']
        cur.execute("INSERT INTO nivel_tecnologico (nombre) VALUES (?) RETURNING id" if USE_POSTGRES else "INSERT INTO nivel_tecnologico (nombre) VALUES (?)", (nombre,))
        if USE_POSTGRES:
            nid = cur.fetchone()['id']
        else:
            nid = cur.lastrowid
        conn.commit()
        conn.close()
        return nid

    # --- getters por ID ---
    def obtener_nombre_cultivo(self, cultivo_id: int) -> str:
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT nombre FROM cultivo WHERE id = ?", (cultivo_id,))
        row = cur.fetchone()
        conn.close()
        return row['nombre'] if row else ""

    def obtener_nombre_variedad(self, variedad_id: Optional[int]) -> str:
        if not variedad_id:
            return ""
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT nombre FROM variedad WHERE id = ?", (variedad_id,))
        row = cur.fetchone()
        conn.close()
        return row['nombre'] if row else ""

    def obtener_nombre_departamento(self, depto_id: int) -> str:
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT nombre FROM departamento WHERE id = ?", (depto_id,))
        row = cur.fetchone()
        conn.close()
        return row['nombre'] if row else ""

    def obtener_nombre_municipio(self, municipio_id: Optional[int]) -> str:
        if not municipio_id:
            return ""
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT nombre FROM municipio WHERE id = ?", (municipio_id,))
        row = cur.fetchone()
        conn.close()
        return row['nombre'] if row else ""

    def obtener_nombre_nivel_tecnologico(self, nivel_id: int) -> str:
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT nombre FROM nivel_tecnologico WHERE id = ?", (nivel_id,))
        row = cur.fetchone()
        conn.close()
        return row['nombre'] if row else ""

    # ============================================================
    # CRUD CONCEPTOS
    # ============================================================

    def listar_conceptos(self) -> pd.DataFrame:
        conn = self._get_conn()
        query = """
            SELECT 
                c.id,
                c.nombre as CONCEPTO,
                u.nombre as UNIDAD,
                c.precio_referencial as PRECIO_UNITARIO,
                c.categoria as CATEGORIA,
                c.clase_rpc as CLASE_RPC,
                c.lugar as LUGAR,
                c.cultivo_asociado as CULTIVO,
                c.observaciones as OBSERVACIONES
            FROM concepto c
            LEFT JOIN unidad_medida u ON c.unidad_id = u.id
            ORDER BY c.categoria, c.nombre
        """
        df = pd.read_sql(query, conn)
        conn.close()
        if not df.empty:
            df['PRECIO_UNITARIO'] = pd.to_numeric(df['PRECIO_UNITARIO'], errors='coerce').fillna(0)
            string_cols = ['CONCEPTO', 'UNIDAD', 'CATEGORIA', 'CLASE_RPC', 'LUGAR', 'CULTIVO', 'OBSERVACIONES']
            for col in string_cols:
                if col in df.columns:
                    df[col] = df[col].fillna('').astype(str)
        return df

    def obtener_concepto_por_nombre(self, nombre: str) -> Optional[Dict]:
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT * FROM concepto WHERE nombre = ?", (nombre,))
        row = cur.fetchone()
        conn.close()
        return self._row_to_dict(row)

    def obtener_conceptos_por_nombres(self, nombres: List[str]) -> pd.DataFrame:
        if not nombres:
            return pd.DataFrame()
        placeholders = ','.join(['?'] * len(nombres))
        conn = self._get_conn()
        query = f"""
            SELECT 
                c.id,
                c.nombre as CONCEPTO,
                c.categoria as CATEGORIA,
                c.clase_rpc as CLASE_RPC,
                c.precio_referencial as PRECIO_UNITARIO,
                c.observaciones as OBSERVACIONES,
                u.nombre as UNIDAD
            FROM concepto c
            LEFT JOIN unidad_medida u ON c.unidad_id = u.id
            WHERE c.nombre IN ({placeholders})
        """
        df = pd.read_sql(query, conn, params=nombres)
        conn.close()
        if not df.empty:
            df['PRECIO_UNITARIO'] = pd.to_numeric(df['PRECIO_UNITARIO'], errors='coerce').fillna(0)
            for col in ['CONCEPTO', 'UNIDAD', 'CATEGORIA', 'CLASE_RPC', 'OBSERVACIONES']:
                if col in df.columns:
                    df[col] = df[col].fillna('').astype(str)
        return df

    def agregar_concepto(self, datos: Dict) -> bool:
        try:
            conn = self._get_conn()
            cur = self._cursor(conn)
            unidad_nombre = datos.get('UNIDAD', '').strip()
            unidad_id = self._get_id('unidad_medida', 'nombre', unidad_nombre) if unidad_nombre else None
            if not unidad_id:
                cur.execute("INSERT INTO unidad_medida (nombre) VALUES (?) ON CONFLICT (nombre) DO NOTHING" if USE_POSTGRES else "INSERT OR IGNORE INTO unidad_medida (nombre) VALUES (?)", (unidad_nombre or 'Unidad',))
                if USE_POSTGRES:
                    conn.commit()
                unidad_id = self._get_id('unidad_medida', 'nombre', unidad_nombre or 'Unidad')
            if not unidad_id:
                unidad_id = 1
            
            cur.execute("""
                INSERT INTO concepto 
                (nombre, unidad_id, categoria, clase_rpc, precio_referencial, 
                 lugar, cultivo_asociado, observaciones)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                datos.get('CONCEPTO', '').strip(),
                unidad_id,
                datos.get('CATEGORIA', '').upper().strip(),
                datos.get('CLASE_RPC', 'bnt').lower().strip(),
                float(datos.get('PRECIO_UNITARIO', 0) or 0),
                datos.get('LUGAR', '').strip(),
                datos.get('CULTIVO', '').strip(),
                datos.get('OBSERVACIONES', '').strip()
            ))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error agregando concepto: {e}")
            return False

    def actualizar_concepto(self, id: int, datos: Dict) -> bool:
        try:
            conn = self._get_conn()
            cur = self._cursor(conn)
            campos = []
            valores = []
            mapping = {
                'CONCEPTO': 'nombre',
                'UNIDAD': 'unidad_id',
                'PRECIO_UNITARIO': 'precio_referencial',
                'CATEGORIA': 'categoria',
                'CLASE_RPC': 'clase_rpc',
                'LUGAR': 'lugar',
                'CULTIVO': 'cultivo_asociado',
                'OBSERVACIONES': 'observaciones'
            }
            if 'UNIDAD' in datos:
                unidad_nombre = datos['UNIDAD'].strip()
                unidad_id = self._get_id('unidad_medida', 'nombre', unidad_nombre) if unidad_nombre else None
                if not unidad_id and unidad_nombre:
                    cur.execute("INSERT INTO unidad_medida (nombre) VALUES (?) ON CONFLICT (nombre) DO NOTHING" if USE_POSTGRES else "INSERT OR IGNORE INTO unidad_medida (nombre) VALUES (?)", (unidad_nombre,))
                    if USE_POSTGRES:
                        conn.commit()
                    unidad_id = self._get_id('unidad_medida', 'nombre', unidad_nombre)
                if unidad_id:
                    campos.append('unidad_id = ?')
                    valores.append(unidad_id)
            for col_pandas, col_db in mapping.items():
                if col_pandas in datos and col_pandas != 'UNIDAD':
                    val = datos[col_pandas]
                    if col_pandas in ['PRECIO_UNITARIO']:
                        val = float(val or 0)
                    else:
                        val = str(val).strip()
                    campos.append(f'{col_db} = ?')
                    valores.append(val)
            if not campos:
                return False
            valores.append(id)
            cur.execute(f"UPDATE concepto SET {', '.join(campos)} WHERE id = ?", tuple(valores))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error actualizando concepto: {e}")
            return False

    def eliminar_concepto(self, id: int) -> bool:
        try:
            conn = self._get_conn()
            cur = self._cursor(conn)
            cur.execute("DELETE FROM concepto WHERE id = ?", (id,))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error eliminando concepto: {e}")
            return False

    def guardar_conceptos_bulk(self, df: pd.DataFrame) -> bool:
        try:
            conn = self._get_conn()
            cur = self._cursor(conn)
            for _, row in df.iterrows():
                id_val = row.get('id')
                concepto = str(row.get('CONCEPTO', '')).strip()
                if not concepto:
                    continue
                unidad_nombre = str(row.get('UNIDAD', '')).strip()
                unidad_id = self._get_id('unidad_medida', 'nombre', unidad_nombre) if unidad_nombre else None
                if not unidad_id and unidad_nombre:
                    cur.execute("INSERT INTO unidad_medida (nombre) VALUES (?) ON CONFLICT (nombre) DO NOTHING" if USE_POSTGRES else "INSERT OR IGNORE INTO unidad_medida (nombre) VALUES (?)", (unidad_nombre,))
                    if USE_POSTGRES:
                        conn.commit()
                    unidad_id = self._get_id('unidad_medida', 'nombre', unidad_nombre)
                if not unidad_id:
                    unidad_id = 1
                precio = float(row.get('PRECIO_UNITARIO', 0) or 0)
                categoria = str(row.get('CATEGORIA', '')).upper().strip()
                clase_rpc = str(row.get('CLASE_RPC', 'bnt')).lower().strip()
                lugar = str(row.get('LUGAR', '')).strip()
                cultivo_asoc = str(row.get('CULTIVO', '')).strip()
                observaciones = str(row.get('OBSERVACIONES', '')).strip()

                if id_val and isinstance(id_val, (int, float)):
                    cur.execute("""
                        UPDATE concepto SET
                            nombre = ?, unidad_id = ?, precio_referencial = ?,
                            categoria = ?, clase_rpc = ?, lugar = ?,
                            cultivo_asociado = ?, observaciones = ?
                        WHERE id = ?
                    """, (concepto, unidad_id, precio, categoria, clase_rpc, lugar,
                          cultivo_asoc, observaciones, id_val))
                else:
                    cur.execute("""
                        INSERT INTO concepto
                        (nombre, unidad_id, precio_referencial, categoria, clase_rpc,
                         lugar, cultivo_asociado, observaciones)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT (nombre) DO NOTHING
                    """ if USE_POSTGRES else """
                        INSERT OR IGNORE INTO concepto
                        (nombre, unidad_id, precio_referencial, categoria, clase_rpc,
                         lugar, cultivo_asociado, observaciones)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (concepto, unidad_id, precio, categoria, clase_rpc,
                          lugar, cultivo_asoc, observaciones))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error guardando conceptos en masa: {e}")
            return False

    # ============================================================
    # CRUD PLANTILLAS
    # ============================================================
    def listar_plantillas(self, solo_activas: bool = True) -> pd.DataFrame:
        conn = self._get_conn()
        query = """
            SELECT 
                p.id,
                c.nombre as CULTIVO,
                v.nombre as VARIEDAD,
                nt.nombre as NIVEL_TECNOLOGICO,
                d.nombre as DEPARTAMENTO,
                m.nombre as MUNICIPIO,
                p.campania as CAMPANIA,
                p.rendimiento_sp as RENDIMIENTO_SP,
                p.rendimiento_cp as RENDIMIENTO_CP,
                p.precio_ref_bs_ton as PRECIO_REF_BS_TON,
                p.perdidas_sp_pct as PERDIDAS_SP_PCT,
                p.perdidas_cp_pct as PERDIDAS_CP_PCT,
                p.activo as ACTIVO,
                p.fuente as FUENTE
            FROM plantilla_costos p
            LEFT JOIN cultivo c ON p.cultivo_id = c.id
            LEFT JOIN variedad v ON p.variedad_id = v.id
            LEFT JOIN nivel_tecnologico nt ON p.nivel_tecnologico_id = nt.id
            LEFT JOIN departamento d ON p.departamento_id = d.id
            LEFT JOIN municipio m ON p.municipio_id = m.id
        """
        if solo_activas:
            query += " WHERE p.activo = " + ("TRUE" if USE_POSTGRES else "1")
        query += " ORDER BY c.nombre, nt.nombre, d.nombre, m.nombre"
        df = pd.read_sql(query, conn)
        conn.close()
        for col in ['CULTIVO', 'VARIEDAD', 'NIVEL_TECNOLOGICO', 'DEPARTAMENTO', 'MUNICIPIO', 'CAMPANIA', 'FUENTE']:
            if col in df.columns:
                df[col] = df[col].fillna('').astype(str)
        numeric_cols = ['RENDIMIENTO_SP', 'RENDIMIENTO_CP', 'PRECIO_REF_BS_TON', 'PERDIDAS_SP_PCT', 'PERDIDAS_CP_PCT']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        return df

    def obtener_plantilla_por_id(self, plantilla_id: int) -> Optional[Dict]:
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT * FROM plantilla_costos WHERE id = ?", (plantilla_id,))
        row = cur.fetchone()
        conn.close()
        return self._row_to_dict(row)

    def agregar_plantilla(self, datos: Dict) -> Optional[int]:
        try:
            conn = self._get_conn()
            cur = self._cursor(conn)
            cur.execute("""
                INSERT INTO plantilla_costos
                (cultivo_id, variedad_id, nivel_tecnologico_id, departamento_id, municipio_id,
                 campania, rendimiento_sp, rendimiento_cp, precio_ref_bs_ton, perdidas_sp_pct, perdidas_cp_pct, fuente)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                RETURNING id
            """ if USE_POSTGRES else """
                INSERT INTO plantilla_costos
                (cultivo_id, variedad_id, nivel_tecnologico_id, departamento_id, municipio_id,
                 campania, rendimiento_sp, rendimiento_cp, precio_ref_bs_ton, perdidas_sp_pct, perdidas_cp_pct, fuente)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                datos.get('cultivo_id'),
                datos.get('variedad_id'),
                datos.get('nivel_tecnologico_id'),
                datos.get('departamento_id'),
                datos.get('municipio_id'),
                datos.get('campania', ''),
                float(datos.get('rendimiento_sp', 0) or 0),
                float(datos.get('rendimiento_cp', 0) or 0),
                float(datos.get('precio_ref_bs_ton', 0) or 0),
                float(datos.get('perdidas_sp_pct', 0) or 0),
                float(datos.get('perdidas_cp_pct', 0) or 0),
                datos.get('fuente', '')
            ))
            if USE_POSTGRES:
                plantilla_id = cur.fetchone()['id']
            else:
                plantilla_id = cur.lastrowid
            conn.commit()
            conn.close()
            return plantilla_id
        except Exception as e:
            print(f"Error agregando plantilla: {e}")
            return None

    def actualizar_plantilla(self, plantilla_id: int, datos: Dict) -> bool:
        try:
            conn = self._get_conn()
            cur = self._cursor(conn)
            campos = []
            valores = []
            for key in ['cultivo_id', 'variedad_id', 'nivel_tecnologico_id', 'departamento_id', 'municipio_id',
                        'campania', 'rendimiento_sp', 'rendimiento_cp', 'precio_ref_bs_ton', 'perdidas_sp_pct', 'perdidas_cp_pct',
                        'activo', 'fuente']:
                if key in datos:
                    campos.append(f"{key} = ?")
                    valores.append(datos[key])
            if not campos:
                return False
            valores.append(plantilla_id)
            cur.execute(f"UPDATE plantilla_costos SET {', '.join(campos)} WHERE id = ?", tuple(valores))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error actualizando plantilla: {e}")
            return False

    def eliminar_plantilla(self, plantilla_id: int) -> bool:
        try:
            conn = self._get_conn()
            cur = self._cursor(conn)
            cur.execute("DELETE FROM plantilla_costos WHERE id = ?", (plantilla_id,))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error eliminando plantilla: {e}")
            return False

    # ============================================================
    # RELACIONES PLANTILLA ↔ CONCEPTOS
    # ============================================================
    def obtener_conceptos_de_plantilla(self, plantilla_id: int) -> pd.DataFrame:
        conn = self._get_conn()
        query = """
            SELECT 
                pc.id as rel_id,
                c.id as concepto_id,
                c.nombre as CONCEPTO,
                u.nombre as UNIDAD,
                pc.cantidad as CANTIDAD,
                pc.cantidad_cp as CANTIDAD_CP,
                c.precio_referencial as PRECIO_UNITARIO,
                c.categoria as CATEGORIA,
                c.clase_rpc as CLASE_RPC,
                pc.precio_override as PRECIO_OVERRIDE,
                pc.observaciones as OBSERVACIONES,
                pc.orden as ORDEN
            FROM plantilla_concepto pc
            JOIN concepto c ON pc.concepto_id = c.id
            LEFT JOIN unidad_medida u ON c.unidad_id = u.id
            WHERE pc.plantilla_id = ?
            ORDER BY pc.orden, c.nombre
        """
        df = pd.read_sql(query, conn, params=(plantilla_id,))
        conn.close()
        if not df.empty:
            for col in ['CANTIDAD', 'CANTIDAD_CP', 'PRECIO_UNITARIO', 'PRECIO_OVERRIDE']:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
            for col in ['CONCEPTO', 'UNIDAD', 'CATEGORIA', 'CLASE_RPC', 'OBSERVACIONES']:
                if col in df.columns:
                    df[col] = df[col].fillna('').astype(str)
        return df

    def asignar_conceptos_a_plantilla(self, plantilla_id: int, conceptos_ids: List[int]) -> bool:
        try:
            conn = self._get_conn()
            cur = self._cursor(conn)
            cur.execute("DELETE FROM plantilla_concepto WHERE plantilla_id = ?", (plantilla_id,))
            for idx, c_id in enumerate(conceptos_ids):
                cur.execute("""
                    INSERT INTO plantilla_concepto
                    (plantilla_id, concepto_id, cantidad, cantidad_cp, orden)
                    VALUES (?, ?, ?, ?, ?)
                """, (plantilla_id, c_id, 1.0, None, idx))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error asignando conceptos: {e}")
            return False

    def actualizar_relacion_concepto_plantilla(self, rel_id: int, cantidad: float, precio_override: float, observaciones: str) -> bool:
        try:
            conn = self._get_conn()
            cur = self._cursor(conn)
            cur.execute("""
                UPDATE plantilla_concepto
                SET cantidad = ?, precio_override = ?, observaciones = ?
                WHERE id = ?
            """, (cantidad, precio_override, observaciones, rel_id))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error actualizando relación: {e}")
            return False

    # ============================================================
    # GASTOS GENERALES DE PLANTILLA
    # ============================================================

    def obtener_gastos_generales(self, plantilla_id: int) -> pd.DataFrame:
        conn = self._get_conn()
        df = pd.read_sql("SELECT * FROM plantilla_gasto_general WHERE plantilla_id = ?", conn, params=(plantilla_id,))
        conn.close()
        return df

    def guardar_gasto_general(self, plantilla_id: int, tipo: str, porcentaje: float = 0, monto_fijo: float = 0, base_calculo: str = 'directo', descripcion: str = '') -> bool:
        try:
            conn = self._get_conn()
            cur = self._cursor(conn)
            if USE_POSTGRES:
                cur.execute("""
                    INSERT INTO plantilla_gasto_general (plantilla_id, tipo, porcentaje, monto_fijo, base_calculo, descripcion)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ON CONFLICT (id) DO UPDATE SET
                        tipo = EXCLUDED.tipo,
                        porcentaje = EXCLUDED.porcentaje,
                        monto_fijo = EXCLUDED.monto_fijo,
                        base_calculo = EXCLUDED.base_calculo,
                        descripcion = EXCLUDED.descripcion
                """, (plantilla_id, tipo, porcentaje, monto_fijo, base_calculo, descripcion))
            else:
                cur.execute("""
                    INSERT OR REPLACE INTO plantilla_gasto_general
                    (plantilla_id, tipo, porcentaje, monto_fijo, base_calculo, descripcion)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (plantilla_id, tipo, porcentaje, monto_fijo, base_calculo, descripcion))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error guardando gasto general: {e}")
            return False

    def eliminar_gasto_general(self, id: int) -> bool:
        try:
            conn = self._get_conn()
            cur = self._cursor(conn)
            cur.execute("DELETE FROM plantilla_gasto_general WHERE id = ?", (id,))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error eliminando gasto general: {e}")
            return False

    # ============================================================
    # MÉTODOS PARA COMPATIBILIDAD CON LA ANTIGUA UI
    # ============================================================

    def obtener_cultivos_disponibles(self) -> List[str]:
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT nombre FROM cultivo ORDER BY nombre")
        rows = cur.fetchall()
        conn.close()
        return [row['nombre'] for row in rows]

    def obtener_nombres_conceptos(self) -> List[str]:
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT nombre FROM concepto ORDER BY nombre")
        rows = cur.fetchall()
        conn.close()
        return [row['nombre'] for row in rows]

    def obtener_conceptos_por_categoria(self, categoria: str) -> List[str]:
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT nombre FROM concepto WHERE categoria = ? ORDER BY nombre", (categoria.upper(),))
        rows = cur.fetchall()
        conn.close()
        return [row['nombre'] for row in rows]

    def get_categoria_conceptos_dict(self) -> Dict[str, List[str]]:
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("SELECT categoria, nombre FROM concepto ORDER BY categoria, nombre")
        rows = cur.fetchall()
        conn.close()
        result = {}
        for row in rows:
            cat = row['categoria'] or 'OTROS'
            if cat not in result:
                result[cat] = []
            result[cat].append(row['nombre'])
        return result

    def get_cultivos_con_variedades(self) -> Dict[str, List[str]]:
        conn = self._get_conn()
        cur = self._cursor(conn)
        cur.execute("""
            SELECT c.nombre as cultivo, v.nombre as variedad
            FROM cultivo c
            LEFT JOIN variedad v ON c.id = v.cultivo_id
            ORDER BY c.nombre, v.nombre
        """)
        rows = cur.fetchall()
        conn.close()
        result = {}
        for row in rows:
            cultivo = row['cultivo']
            variedad = row['variedad'] or ''
            if cultivo not in result:
                result[cultivo] = []
            if variedad:
                result[cultivo].append(variedad)
        return result

    # ============================================================
    # BÚSQUEDA INTELIGENTE DE CONCEPTOS POR CULTIVO
    # ============================================================

    def obtener_conceptos_por_cultivo(self, cultivo_nombre: str) -> pd.DataFrame:
        if not cultivo_nombre:
            return self.listar_conceptos()
        primera = cultivo_nombre.split()[0].strip().lower()
        if not primera:
            return self.listar_conceptos()
        conn = self._get_conn()
        query = """
            SELECT 
                c.id,
                c.nombre as CONCEPTO,
                u.nombre as UNIDAD,
                c.precio_referencial as PRECIO_UNITARIO,
                c.categoria as CATEGORIA,
                c.clase_rpc as CLASE_RPC,
                c.lugar as LUGAR,
                c.cultivo_asociado as CULTIVO,
                c.observaciones as OBSERVACIONES
            FROM concepto c
            LEFT JOIN unidad_medida u ON c.unidad_id = u.id
            WHERE 
                LOWER(c.cultivo_asociado) LIKE (? || ' %')
                OR LOWER(c.cultivo_asociado) LIKE (? || ',%')
                OR LOWER(c.cultivo_asociado) = ?
                OR c.cultivo_asociado = ''
                OR c.cultivo_asociado IS NULL
            ORDER BY c.categoria, c.nombre
        """
        df = pd.read_sql(query, conn, params=(primera, primera, primera))
        conn.close()
        if not df.empty:
            df['PRECIO_UNITARIO'] = pd.to_numeric(df['PRECIO_UNITARIO'], errors='coerce').fillna(0)
            for col in ['CONCEPTO', 'UNIDAD', 'CATEGORIA', 'CLASE_RPC', 'LUGAR', 'CULTIVO', 'OBSERVACIONES']:
                if col in df.columns:
                    df[col] = df[col].fillna('').astype(str)
        return df

    def obtener_conceptos_por_cultivo_y_lugar(self, cultivo_nombre: str, lugar: str) -> pd.DataFrame:
        if not cultivo_nombre:
            return self.listar_conceptos()
        primera = cultivo_nombre.split()[0].strip().lower()
        lugar = lugar.strip().lower()
        conn = self._get_conn()
        if lugar:
            query = """
                SELECT 
                    c.id,
                    c.nombre as CONCEPTO,
                    u.nombre as UNIDAD,
                    c.precio_referencial as PRECIO_UNITARIO,
                    c.categoria as CATEGORIA,
                    c.clase_rpc as CLASE_RPC,
                    c.lugar as LUGAR,
                    c.cultivo_asociado as CULTIVO,
                    c.observaciones as OBSERVACIONES
                FROM concepto c
                LEFT JOIN unidad_medida u ON c.unidad_id = u.id
                WHERE 
                    (LOWER(c.cultivo_asociado) LIKE (? || ' %')
                     OR LOWER(c.cultivo_asociado) LIKE (? || ',%')
                     OR LOWER(c.cultivo_asociado) = ?
                     OR c.cultivo_asociado = ''
                     OR c.cultivo_asociado IS NULL)
                    AND (LOWER(c.lugar) LIKE ('%' || ? || '%')
                         OR c.lugar = ''
                         OR c.lugar IS NULL)
                ORDER BY c.categoria, c.nombre
            """
            params = (primera, primera, primera, lugar)
        else:
            return self.obtener_conceptos_por_cultivo(cultivo_nombre)
        df = pd.read_sql(query, conn, params=params)
        conn.close()
        if not df.empty:
            df['PRECIO_UNITARIO'] = pd.to_numeric(df['PRECIO_UNITARIO'], errors='coerce').fillna(0)
            for col in ['CONCEPTO', 'UNIDAD', 'CATEGORIA', 'CLASE_RPC', 'LUGAR', 'CULTIVO', 'OBSERVACIONES']:
                if col in df.columns:
                    df[col] = df[col].fillna('').astype(str)
        return df

    def obtener_conceptos_de_plantilla_para_calculo(self, plantilla_id: int) -> pd.DataFrame:
        conn = self._get_conn()
        query = """
            SELECT 
                c.nombre as CONCEPTO,
                u.nombre as UNIDAD,
                pc.cantidad as CANTIDAD,
                pc.cantidad_cp as CANTIDAD_CP,
                c.precio_referencial as PRECIO_UNITARIO,
                pc.precio_override as PRECIO_OVERRIDE,
                c.categoria as CATEGORIA,
                c.clase_rpc as CLASE_RPC,
                COALESCE(NULLIF(pc.observaciones, ''), c.observaciones) as OBSERVACIONES
            FROM plantilla_concepto pc
            JOIN concepto c ON pc.concepto_id = c.id
            LEFT JOIN unidad_medida u ON c.unidad_id = u.id
            WHERE pc.plantilla_id = ?
            ORDER BY pc.orden, c.nombre
        """
        df = pd.read_sql(query, conn, params=(plantilla_id,))
        conn.close()
        if not df.empty:
            df['CANTIDAD'] = pd.to_numeric(df['CANTIDAD'], errors='coerce').fillna(1.0)
            df['CANTIDAD_CP'] = pd.to_numeric(df['CANTIDAD_CP'], errors='coerce')
            df['PRECIO_UNITARIO'] = pd.to_numeric(df['PRECIO_UNITARIO'], errors='coerce').fillna(0)
            df['PRECIO_OVERRIDE'] = pd.to_numeric(df['PRECIO_OVERRIDE'], errors='coerce').fillna(0)
            for col in ['CONCEPTO', 'UNIDAD', 'CATEGORIA', 'CLASE_RPC', 'OBSERVACIONES']:
                if col in df.columns:
                    df[col] = df[col].fillna('').astype(str)
        return df

    def guardar_conceptos_plantilla_bulk(self, plantilla_id: int, df: pd.DataFrame) -> bool:
        try:
            conn = self._get_conn()
            cur = self._cursor(conn)
            for _, row in df.iterrows():
                rel_id = row.get('rel_id')
                if pd.isna(rel_id):
                    continue
                cantidad = float(row.get('CANTIDAD', 1.0) or 1.0)
                cantidad_cp_raw = row.get('CANTIDAD_CP')
                if pd.isna(cantidad_cp_raw):
                    cantidad_cp = None
                else:
                    cantidad_cp = float(cantidad_cp_raw)
                precio_ov = float(row.get('PRECIO_OVERRIDE', 0) or 0)
                obs = str(row.get('OBSERVACIONES', ''))
                cur.execute("""
                    UPDATE plantilla_concepto
                    SET cantidad = ?, cantidad_cp = ?, precio_override = ?, observaciones = ?
                    WHERE id = ? AND plantilla_id = ?
                """, (cantidad, cantidad_cp, precio_ov, obs, int(rel_id), plantilla_id))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error guardando ajustes de plantilla_concepto: {e}")
            return False
