# core/project_manager.py
import os
import json
import shutil
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from .data_manager import DataManager
from .config import ConfiguracionProyecto

from saas.auth import obtener_usuario_actual, incrementar_proyectos_usados
from saas.limites import verificar_puede_crear_proyecto, SuscripcionVencidaError, LimiteExcedidoError

class ProjectManager:
    
    BASE_DIR = "data/proyectos"
    INDEX_FILE = "data/proyectos/index.json"

    def get_ruta_global_db(self) -> str:
        return os.path.join("data", "global.db")

    def __init__(self):
        os.makedirs(self.BASE_DIR, exist_ok=True)
        self.index = self._cargar_index()
    
    def _cargar_index(self) -> Dict:
        """Carga o crea el índice de proyectos."""
        if os.path.exists(self.INDEX_FILE):
            try:
                with open(self.INDEX_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return {"proyectos": {}, "activo": None}
    
    def _guardar_index(self):
        """Persiste el índice."""
        os.makedirs(os.path.dirname(self.INDEX_FILE), exist_ok=True)
        with open(self.INDEX_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.index, f, indent=2, ensure_ascii=False)
    
    def _ruta_proyecto(self, slug: str) -> str:
        """Retorna la ruta base de un proyecto."""
        return os.path.join(self.BASE_DIR, slug)
    
    def _ruta_excel(self, slug: str) -> str:
        return os.path.join(self._ruta_proyecto(slug), "proyecto_activo.xlsx")
    
    def _ruta_config(self, slug: str) -> str:
        return os.path.join(self._ruta_proyecto(slug), "config_proyecto.json")
    
    def listar_proyectos(self) -> List[Dict]:
        """Lista todos los proyectos registrados."""
        proyectos = []
        for slug, meta in self.index.get("proyectos", {}).items():
            proyectos.append({
                "slug": slug,
                "nombre": meta.get("nombre", slug),
                "creado": meta.get("creado", ""),
                "modificado": meta.get("modificado", ""),
                "activo": (slug == self.index.get("activo"))
            })
        return sorted(proyectos, key=lambda x: x["modificado"], reverse=True)
    
    #def crear_proyecto(self, nombre: str, config: ConfiguracionProyecto = None) -> Tuple[bool, str]:
    def crear_proyecto(self, nombre: str, config=None):
        # --- VALIDACIÓN SaaS ---
        usuario = obtener_usuario_actual()
        if usuario and usuario.rol.value != "super_admin":
            try:
                verificar_puede_crear_proyecto(usuario.id)
            except (SuscripcionVencidaError, LimiteExcedidoError) as e:
                return False, str(e)
        # Hasta aqui, solo para DESPLIEGUE

        # Generar slug único
        base_slug = "".join(c if c.isalnum() else "_" for c in nombre).lower()
        slug = base_slug
        contador = 1
        while slug in self.index["proyectos"] or os.path.exists(self._ruta_proyecto(slug)):
            slug = f"{base_slug}_{contador}"
            contador += 1
        
        # Crear carpeta
        ruta = self._ruta_proyecto(slug)
        os.makedirs(ruta, exist_ok=True)
        
        # Crear archivos base
        dm = DataManager(self._ruta_excel(slug))
        dm.crear_proyecto_nuevo()
        
        if config is None:
            config = ConfiguracionProyecto(nombre=nombre)
        config.guardar(self._ruta_config(slug))
        
        # Registrar en índice
        self.index["proyectos"][slug] = {
            "nombre": nombre,
            "creado": datetime.now().isoformat(),
            "modificado": datetime.now().isoformat()
        }
        self.index["activo"] = slug
        self._guardar_index()

        # Al final, solo para DESPLIEGUE
        if usuario and usuario.rol.value != "super_admin":
            incrementar_proyectos_usados(usuario.id)
        return True, slug

    def cargar_proyecto(self, slug: str) -> Tuple[bool, str]:
        """Activa un proyecto existente."""
        if slug not in self.index["proyectos"]:
            return False, f"Proyecto '{slug}' no existe"
        
        self.index["activo"] = slug
        self.index["proyectos"][slug]["modificado"] = datetime.now().isoformat()
        self._guardar_index()
        return True, slug
    
    def eliminar_proyecto(self, slug: str) -> bool:
        """Elimina un proyecto y su carpeta."""
        if slug not in self.index["proyectos"]:
            return False
        
        ruta = self._ruta_proyecto(slug)
        if os.path.exists(ruta):
            shutil.rmtree(ruta)
        
        del self.index["proyectos"][slug]
        if self.index.get("activo") == slug:
            self.index["activo"] = None
        self._guardar_index()
        
        # Al final, solo para DESPLIEGUE
        usuario = obtener_usuario_actual()
        if usuario and usuario.rol.value != "super_admin":
            from saas.auth import decrementar_proyectos_usados
            decrementar_proyectos_usados(usuario.id)

        return True

    def duplicar_proyecto(self, slug_origen: str, nombre_nuevo: str) -> Tuple[bool, str]:
        """Duplica un proyecto existente."""
        if slug_origen not in self.index["proyectos"]:
            return False, "Proyecto origen no existe"
        
        ok, slug_nuevo = self.crear_proyecto(nombre_nuevo)
        if not ok:
            return False, slug_nuevo
        
        # Copiar archivos
        ruta_origen = self._ruta_proyecto(slug_origen)
        ruta_nuevo = self._ruta_proyecto(slug_nuevo)
        
        for archivo in ["proyecto_activo.xlsx", "config_proyecto.json"]:
            origen = os.path.join(ruta_origen, archivo)
            destino = os.path.join(ruta_nuevo, archivo)
            if os.path.exists(origen):
                shutil.copy2(origen, destino)
        
        return True, slug_nuevo
    
    def get_rutas_activas(self) -> Optional[Dict[str, str]]:
        """Retorna las rutas del proyecto activo."""
        slug = self.index.get("activo")
        if slug is None:
            return None
        return {
            "slug": slug,
            "nombre": self.index["proyectos"][slug]["nombre"],
            "excel": self._ruta_excel(slug),
            "config": self._ruta_config(slug),
            "carpeta": self._ruta_proyecto(slug),
            "db": os.path.join(self._ruta_proyecto(slug), "proyecto.db") 
        }
    
    def existe_activo(self) -> bool:
        """Verifica si hay un proyecto activo válido."""
        rutas = self.get_rutas_activas()
        if rutas is None:
            return False
        return os.path.exists(rutas["excel"]) and os.path.exists(rutas["config"])

    def get_ruta_db(self) -> Optional[str]:
        """Retorna la ruta de la BD SQLite del proyecto activo."""
        rutas = self.get_rutas_activas()
        if not rutas:
            return None
        return os.path.join(rutas["carpeta"], "proyecto.db")

    def obtener_db(self) -> Optional['ProyectoDB']:
        """Instancia la BD del proyecto activo."""
        ruta = self.get_ruta_db()
        if ruta:
            from .database import ProyectoDB
            return ProyectoDB(ruta)
        return None