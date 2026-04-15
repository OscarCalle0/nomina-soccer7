"""
datos.py — Modelos de datos y persistencia para Sistema de Nómina Soccer 7
Usa JSON en Google Drive (o local) como base de datos simple y portátil.
"""

import json
import os
from datetime import datetime, date
from dataclasses import dataclass, field, asdict
from typing import Optional, List, Dict
import pandas as pd

# ── Ruta de datos ──────────────────────────────────────────────────────────────
# En Streamlit Cloud, guardar en el directorio de la app
DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
os.makedirs(DATA_DIR, exist_ok=True)

COLABORADORES_FILE  = os.path.join(DATA_DIR, "colaboradores.json")
QUINCENAS_FILE      = os.path.join(DATA_DIR, "quincenas.json")
HISTORICO_FILE      = os.path.join(DATA_DIR, "historico_nomina.json")


# ── Modelos ────────────────────────────────────────────────────────────────────

@dataclass
class Colaborador:
    id: str                         # cédula
    nombre_completo: str            # BERTHA LIBIA RESTREPO LEDEZMA
    nombre_reloj: str               # BERTHA RESTREPO (como aparece en el reloj)
    cargo: str
    salario_mensual: float
    fecha_ingreso: str              # YYYY-MM-DD
    fecha_retiro: Optional[str]     # YYYY-MM-DD o None si activo
    banco: str
    cuenta: str
    eps: str
    tipo: str                       # "empleado" | "prestador"
    valor_hora_prestador: float     # Solo para prestadores
    activo: bool
    notas: str = ""

    def salario_en(self, fecha_str: str) -> float:
        """Salario vigente en una fecha (para histórico)"""
        return self.salario_mensual

    def esta_activo_en(self, p_ini: date, p_fin: date) -> bool:
        """¿Trabajó en este período?"""
        if not self.activo and self.fecha_retiro:
            retiro = date.fromisoformat(self.fecha_retiro)
            ingreso = date.fromisoformat(self.fecha_ingreso)
            return ingreso <= p_fin and retiro >= p_ini
        ingreso = date.fromisoformat(self.fecha_ingreso)
        return ingreso <= p_fin

    def dias_en_periodo(self, p_ini: date, p_fin: date) -> int:
        """Días que corresponden en el período (para proporcional)"""
        ingreso = date.fromisoformat(self.fecha_ingreso)
        inicio_real = max(p_ini, ingreso)
        if self.fecha_retiro:
            retiro = date.fromisoformat(self.fecha_retiro)
            fin_real = min(p_fin, retiro)
        else:
            fin_real = p_fin
        if fin_real < inicio_real:
            return 0
        return (fin_real - inicio_real).days + 1


@dataclass
class RegistroQuincena:
    """Una quincena procesada — guardada en el histórico"""
    id: str                      # "2026-04-01_2026-04-15"
    periodo_ini: str
    periodo_fin: str
    fecha_procesado: str
    colaboradores: List[Dict]    # lista de resultados por colaborador
    total_nomina: float
    total_extras_efectivo: float
    total_recargo_nocturno: float
    notas: str = ""


# ── Persistencia ───────────────────────────────────────────────────────────────

def _load_json(path: str, default):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return default
    return default

def _save_json(path: str, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)


# ── COLABORADORES ──────────────────────────────────────────────────────────────

def cargar_colaboradores() -> List[Colaborador]:
    data = _load_json(COLABORADORES_FILE, [])
    result = []
    for d in data:
        try:
            result.append(Colaborador(**d))
        except:
            pass
    return result

def guardar_colaboradores(lista: List[Colaborador]):
    _save_json(COLABORADORES_FILE, [asdict(c) for c in lista])

def get_colaborador_por_reloj(nombre_reloj: str, lista: List[Colaborador]) -> Optional[Colaborador]:
    """Busca colaborador por nombre_reloj (insensible a espacios dobles)"""
    nombre_norm = " ".join(nombre_reloj.strip().split())
    for c in lista:
        if " ".join(c.nombre_reloj.strip().split()) == nombre_norm:
            return c
    return None

def get_colaborador_por_id(cedula: str, lista: List[Colaborador]) -> Optional[Colaborador]:
    for c in lista:
        if c.id == cedula:
            return c
    return None

def colaboradores_activos_en(lista: List[Colaborador], p_ini: date, p_fin: date) -> List[Colaborador]:
    return [c for c in lista if c.esta_activo_en(p_ini, p_fin)]

def agregar_colaborador(col: Colaborador) -> bool:
    lista = cargar_colaboradores()
    if any(c.id == col.id for c in lista):
        return False  # ya existe
    lista.append(col)
    guardar_colaboradores(lista)
    return True

def actualizar_colaborador(col: Colaborador):
    lista = cargar_colaboradores()
    for i, c in enumerate(lista):
        if c.id == col.id:
            lista[i] = col
            break
    guardar_colaboradores(lista)

def retirar_colaborador(cedula: str, fecha_retiro: str):
    lista = cargar_colaboradores()
    for c in lista:
        if c.id == cedula:
            c.fecha_retiro = fecha_retiro
            c.activo = False
            break
    guardar_colaboradores(lista)


# ── HISTÓRICO QUINCENAS ────────────────────────────────────────────────────────

def cargar_historico() -> List[Dict]:
    return _load_json(HISTORICO_FILE, [])

def guardar_quincena_historico(registro: Dict):
    historico = cargar_historico()
    # Reemplazar si ya existe el período
    historico = [h for h in historico if h.get("id") != registro.get("id")]
    historico.append(registro)
    historico.sort(key=lambda x: x.get("periodo_ini", ""), reverse=True)
    _save_json(HISTORICO_FILE, historico)

def get_quincena_historico(periodo_id: str) -> Optional[Dict]:
    for h in cargar_historico():
        if h.get("id") == periodo_id:
            return h
    return None

def historico_a_dataframe() -> pd.DataFrame:
    hist = cargar_historico()
    if not hist:
        return pd.DataFrame()
    rows = []
    for q in hist:
        for col in q.get("colaboradores", []):
            rows.append({
                "periodo": q["id"],
                "periodo_ini": q["periodo_ini"],
                "periodo_fin": q["periodo_fin"],
                "nombre": col.get("nombre", ""),
                "horas_trabajadas": col.get("tot_h", 0),
                "horas_extras": col.get("ext_h", 0),
                "recargo_nocturno_h": col.get("noct_h", 0),
                "val_recargo": col.get("val_noct", 0),
                "val_extras_nom": col.get("val_en", 0),
                "val_extras_ef": col.get("val_ee", 0),
                "neto_pagado": col.get("neto", 0),
                "novedades": col.get("novedades_desc", ""),
            })
    return pd.DataFrame(rows)


# ── DATOS INICIALES (colaboradores actuales) ──────────────────────────────────

COLABORADORES_INICIALES = [
    Colaborador(
        id="43413529", nombre_completo="BERTHA LIBIA RESTREPO LEDEZMA",
        nombre_reloj="BERTHA RESTREPO", cargo="JEFE DE COCINA",
        salario_mensual=1806565, fecha_ingreso="2024-01-01", fecha_retiro=None,
        banco="AHORROS DAVIVIENDA", cuenta="3974-0007-7277", eps="NUEVA EPS",
        tipo="empleado", valor_hora_prestador=0, activo=True
    ),
    Colaborador(
        id="1047994257", nombre_completo="ERICA YORLADIS BERMUDEZ RAMIREZ",
        nombre_reloj="ERIKA  BERMUDEZ", cargo="AUXILIAR DE COCINA",
        salario_mensual=1750905, fecha_ingreso="2024-01-01", fecha_retiro=None,
        banco="AHORROS DAVIVIENDA", cuenta="4884-4268-0267", eps="SAVIA SALUD",
        tipo="empleado", valor_hora_prestador=0, activo=True
    ),
    Colaborador(
        id="1033340824", nombre_completo="DANIELA SANCHEZ TORRES",
        nombre_reloj="DANIELA SANCHEZ", cargo="AUXILIAR DE COCINA",
        salario_mensual=1750905, fecha_ingreso="2024-01-01", fecha_retiro=None,
        banco="AHORROS DAVIVIENDA", cuenta="4884-3737-2466", eps="SAVIA SALUD",
        tipo="empleado", valor_hora_prestador=0, activo=True
    ),
    Colaborador(
        id="1033336422", nombre_completo="CAROLINA ARANGO GOMEZ",
        nombre_reloj="CAROLINA ARANGO", cargo="MESERA",
        salario_mensual=1750905, fecha_ingreso="2024-01-01", fecha_retiro=None,
        banco="AHORROS DAVIVIENDA", cuenta="4884-4926-8538", eps="SAVIA SALUD",
        tipo="empleado", valor_hora_prestador=0, activo=True
    ),
    Colaborador(
        id="1045018453", nombre_completo="KAROL DAYANA QUINTERO AGUDELO",
        nombre_reloj="KAROL QUINTERO", cargo="MESERA",
        salario_mensual=1750905, fecha_ingreso="2024-01-01", fecha_retiro=None,
        banco="AHORROS DAVIVIENDA", cuenta="3974-0008-5007", eps="SAVIA SALUD",
        tipo="empleado", valor_hora_prestador=0, activo=True
    ),
    Colaborador(
        id="1000211127", nombre_completo="KATERIN MARYORY SEPULVEDA CRUZ",
        nombre_reloj="KATERINE SEPULVEDA", cargo="MESERA",
        salario_mensual=1750905, fecha_ingreso="2024-01-01", fecha_retiro=None,
        banco="AHORROS DAVIVIENDA", cuenta="4884-4926-8785", eps="SAVIA SALUD",
        tipo="empleado", valor_hora_prestador=0, activo=True
    ),
    Colaborador(
        id="43844703", nombre_completo="MARIA LIGELLA LOTERO",
        nombre_reloj="MARIA LIGELLA LOTERO", cargo="AUXILIAR ADMINISTRATIVA",
        salario_mensual=1750905, fecha_ingreso="2024-01-01", fecha_retiro=None,
        banco="AHORROS DAVIVIENDA", cuenta="4884-5662-0761", eps="SAVIA SALUD",
        tipo="empleado", valor_hora_prestador=0, activo=True
    ),
    Colaborador(
        id="1000397698", nombre_completo="VALENTINA GRANDA AGUDELO",
        nombre_reloj="VALENTINA GRANDA", cargo="AUXILIAR ADMINISTRATIVA",
        salario_mensual=1750905, fecha_ingreso="2024-01-01", fecha_retiro=None,
        banco="AHORROS DAVIVIENDA", cuenta="4884-5408-1842", eps="SAVIA SALUD",
        tipo="empleado", valor_hora_prestador=0, activo=True
    ),
    Colaborador(
        id="0", nombre_completo="JULIANA GOMEZ",
        nombre_reloj="JULIANA GOMEZ", cargo="MESERA",
        salario_mensual=0, fecha_ingreso="2024-01-01", fecha_retiro=None,
        banco="AHORROS DAVIVIENDA", cuenta="", eps="",
        tipo="prestador", valor_hora_prestador=10000, activo=True
    ),
]

def inicializar_datos():
    """Crea los datos iniciales si no existen"""
    if not os.path.exists(COLABORADORES_FILE):
        guardar_colaboradores(COLABORADORES_INICIALES)
        print("✅ Datos iniciales creados")
    else:
        print("✅ Datos existentes cargados")
