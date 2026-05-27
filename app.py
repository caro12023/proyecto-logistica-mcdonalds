import io
import math
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple, Optional

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

STATIONS = [
    "Parrilla",
    "Freidora",
    "Bebidas/Postres",
    "Ensamble",
    "Staging/Bolseo",
]

RULES = [
    "FIFO",
    "SPT",
    "EDD",
    "WSPT",
    "CR",
    "ATC",
    "Holgura crítica + SPT",
]

RULE_ALIASES = {
    "FIFO": "FIFO",
    "SPT": "SPT",
    "EDD": "EDD",
    "WSPT": "WSPT",
    "CR": "CR",
    "ATC": "ATC",
    "Holgura crítica + SPT": "HCritSPT",
}

# Horizonte amplio usado para que el calendario no cierre artificialmente
# antes de que terminen todas las operaciones. El valor final se calcula
# por instancia con calculate_horizon_seg().
DEFAULT_HORIZON_SEG = 20000
HORIZON_BUFFER_SEG = 2000

# Archivo Excel esperado. Déjalo en la misma carpeta de este .py
# para que la app lo cargue automáticamente sin botón de carga.
EXCEL_FILE_NAME = "instancias_mcdonalds(7).xlsx"
EXCEL_CANDIDATES = [
    EXCEL_FILE_NAME,
    "instancias_mcdonalds.xlsx",
    "data/instancias_mcdonalds.xlsx",
    "data/instancias_mcdonalds(7).xlsx",
]


# ============================================================
# FUNCIONES DE LIMPIEZA
# ============================================================

def normalize_text(x) -> str:
    if pd.isna(x):
        return ""
    x = str(x)
    x = unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode("ascii")
    return x.lower().strip()


def normalize_yes(x) -> bool:
    y = normalize_text(x)
    return y in {"si", "s", "yes", "true", "1", "x"}


def normalize_station(x) -> str:
    y = normalize_text(x)

    if "parrilla" in y:
        return "Parrilla"
    if "freidor" in y:
        return "Freidora"
    if "bebida" in y or "postre" in y:
        return "Bebidas/Postres"
    if "ensamble" in y:
        return "Ensamble"
    if "staging" in y or "bolseo" in y:
        return "Staging/Bolseo"
    if "todas" in y or "todos" in y or "general" in y:
        return "Todas"
    if "ninguna" in y or "sin evento" in y or "no aplica" in y:
        return "Ninguna"

    return str(x)


def normalize_channel(x) -> str:
    y = normalize_text(x)

    if "automac" in y or "auto mac" in y or y == "auto":
        return "AutoMac"
    if "mostrador" in y:
        return "Mostrador"
    if "pickup" in y or "pick up" in y:
        return "Pickup"
    if "delivery" in y or "domicilio" in y:
        return "Delivery"

    return str(x)


def to_num(x, default=0.0) -> float:
    try:
        if pd.isna(x):
            return default
        if isinstance(x, str):
            x = x.strip().replace(",", ".")
        return float(x)
    except Exception:
        return default


def safe_col(df: pd.DataFrame, col: str, default=None):
    if col in df.columns:
        return df[col]
    return pd.Series([default] * len(df))


# ============================================================
# LECTURA DE EXCEL
# ============================================================

def read_instance_excel(uploaded_file, instance_number: int):
    xls = pd.ExcelFile(excel_readable(uploaded_file))

    possible_instance_sheets = [
        f"Instancia {instance_number}",
        f"{instance_number} Instancia",
        f"Instancia_{instance_number}",
        str(instance_number),
    ]

    instance_sheet = None
    for sheet in possible_instance_sheets:
        if sheet in xls.sheet_names:
            instance_sheet = sheet
            break

    if instance_sheet is None:
        raise ValueError(
            f"No encontré una hoja para la Instancia {instance_number}. "
            f"Hojas disponibles: {xls.sheet_names}"
        )

    jobs = pd.read_excel(excel_readable(uploaded_file), sheet_name=instance_sheet)

    events = pd.DataFrame()
    capacities = pd.DataFrame()

    if "Eventos" in xls.sheet_names:
        events = pd.read_excel(excel_readable(uploaded_file), sheet_name="Eventos")
        if "Instancia" in events.columns:
            events = events[events["Instancia"] == instance_number].copy()

    if "Capacidades" in xls.sheet_names:
        capacities = pd.read_excel(excel_readable(uploaded_file), sheet_name="Capacidades")
        if "Instancia" in capacities.columns:
            capacities = capacities[capacities["Instancia"] == instance_number].copy()

    return jobs, events, capacities, instance_sheet



def detect_instance_numbers(uploaded_file) -> List[int]:
    """
    Detecta automáticamente las hojas de instancias del archivo.
    Acepta nombres como 'Instancia 1', '1 Instancia', 'Instancia_1' o '1'.
    """
    xls = pd.ExcelFile(excel_readable(uploaded_file))
    instances = []

    for sheet in xls.sheet_names:
        y = normalize_text(sheet)
        match = None

        if "instancia" in y:
            nums = re.findall(r"\d+", y)
            if nums:
                match = int(nums[0])
        elif y.isdigit():
            match = int(y)

        if match is not None:
            instances.append(match)

    return sorted(set(instances))


def calculate_horizon_seg(jobs: pd.DataFrame, capacities: pd.DataFrame) -> float:
    """
    Calcula un horizonte suficientemente amplio para que la programación
    pueda terminar aunque haya tardanzas. Esto evita que el calendario se
    cierre artificialmente en 3600, 4800, 5400, etc.
    """
    max_cap_until = 0.0

    if capacities is not None and not capacities.empty and "Hasta seg" in capacities.columns:
        max_cap_until = capacities["Hasta seg"].apply(to_num).max()

    max_release = jobs["r j seg"].apply(to_num).max() if "r j seg" in jobs.columns else 0.0

    processing_cols = [
        "p Parrilla seg",
        "p Freidora seg",
        "p BebidaPostre seg",
        "p Ensamble seg",
        "p Staging seg",
    ]

    total_processing = 0.0
    for col in processing_cols:
        if col in jobs.columns:
            total_processing += jobs[col].apply(to_num).sum()

    return max(
        DEFAULT_HORIZON_SEG,
        max_cap_until + HORIZON_BUFFER_SEG,
        max_release + total_processing + HORIZON_BUFFER_SEG,
    )


def make_calendar_table(calendars: Dict[str, List[Tuple[float, float]]]) -> pd.DataFrame:
    rows = []

    for resource, intervals in calendars.items():
        rows.append({
            "Recurso": resource,
            "Estacion": station_from_resource(resource),
            "Intervalos activos": " | ".join([f"{a:.0f}-{b:.0f}" for a, b in intervals]),
        })

    return pd.DataFrame(rows)


def clean_sheet_name(name: str) -> str:
    """
    Limpia nombres de hojas para Excel.
    """
    name = str(name)
    for bad in ["[", "]", "*", "?", "/", "\\", ":"]:
        name = name.replace(bad, "_")
    return name[:31]


def excel_readable(excel_source):
    """
    Devuelve una fuente legible por pandas.
    Si el Excel viene de st.file_uploader, se trabaja con bytes para poder
    leerlo varias veces sin depender de la posición interna del archivo.
    Si viene como Path, pandas lo lee directamente.
    """
    if isinstance(excel_source, (bytes, bytearray)):
        return io.BytesIO(excel_source)
    return excel_source


def find_project_excel() -> Path:
    """
    Busca el Excel del proyecto sin pedir carga manual.
    Primero revisa la carpeta donde está este archivo .py y luego
    la carpeta actual desde donde se ejecuta Streamlit.
    """
    script_dir = Path(__file__).resolve().parent
    current_dir = Path.cwd().resolve()

    search_dirs = []
    for directory in [script_dir, current_dir]:
        if directory not in search_dirs:
            search_dirs.append(directory)

    checked_paths = []

    for directory in search_dirs:
        for candidate in EXCEL_CANDIDATES:
            path = directory / candidate
            checked_paths.append(path)
            if path.exists() and path.is_file():
                return path

    checked_text = "\n".join([f"- {p}" for p in checked_paths])
    raise FileNotFoundError(
        "No encontré el archivo Excel del proyecto. "
        "Deja el Excel en la misma carpeta de este archivo Python con alguno de estos nombres: "
        f"{', '.join(EXCEL_CANDIDATES)}.\n\nRutas revisadas:\n{checked_text}"
    )

# ============================================================
# PREPARACIÓN DE PEDIDOS
# ============================================================

def prepare_jobs(jobs: pd.DataFrame) -> pd.DataFrame:
    jobs = jobs.copy()

    if "Canal" in jobs.columns:
        jobs["Canal"] = jobs["Canal"].apply(normalize_channel)
    else:
        jobs["Canal"] = "No definido"

    numeric_cols = [
        "r j seg",
        "SLA seg",
        "d j seg",
        "w j",
        "p Parrilla seg",
        "p Freidora seg",
        "p BebidaPostre seg",
        "p Ensamble seg",
        "p Staging seg",
    ]

    for col in numeric_cols:
        if col in jobs.columns:
            jobs[col] = jobs[col].apply(to_num)
        else:
            jobs[col] = 0.0

    if "d j seg" not in jobs.columns or jobs["d j seg"].sum() == 0:
        jobs["d j seg"] = jobs["r j seg"] + jobs["SLA seg"]

    if "w j" not in jobs.columns or jobs["w j"].sum() == 0:
        jobs["w j"] = 1.0

    if "ID Pedido" not in jobs.columns:
        jobs["ID Pedido"] = [f"P{i+1:03d}" for i in range(len(jobs))]

    if "Tipo Pedido" not in jobs.columns:
        jobs["Tipo Pedido"] = "No definido"

    requirement_cols = [
        "Req Parrilla",
        "Req Freidora",
        "Req BebidaPostre",
        "Req Ensamble",
        "Req Staging",
    ]

    for col in requirement_cols:
        if col not in jobs.columns:
            jobs[col] = "No"

    jobs["Tiempo total procesamiento"] = (
        jobs["p Parrilla seg"]
        + jobs["p Freidora seg"]
        + jobs["p BebidaPostre seg"]
        + jobs["p Ensamble seg"]
        + jobs["p Staging seg"]
    )

    return jobs


# ============================================================
# CREACIÓN DE OPERACIONES
# ============================================================

def build_operations(jobs: pd.DataFrame) -> pd.DataFrame:
    rows = []

    for _, job in jobs.iterrows():
        job_id = str(job["ID Pedido"])
        rj = to_num(job["r j seg"])
        dj = to_num(job["d j seg"])
        wj = max(to_num(job["w j"], 1.0), 1.0)
        tipo = str(job.get("Tipo Pedido", "No definido"))
        canal = str(job.get("Canal", "No definido"))
        pj_total = to_num(job["Tiempo total procesamiento"])

        prep_ops = []

        if normalize_yes(job.get("Req Parrilla", "No")) and to_num(job["p Parrilla seg"]) > 0:
            op_id = f"{job_id}_Parrilla"
            prep_ops.append(op_id)
            rows.append({
                "OpID": op_id,
                "ID Pedido": job_id,
                "Canal": canal,
                "TipoPedido": tipo,
                "Etapa": "Preparación",
                "Estacion": "Parrilla",
                "DuracionProceso": to_num(job["p Parrilla seg"]),
                "rj": rj,
                "dj": dj,
                "wj": wj,
                "pj_total": pj_total,
                "Predecesores": [],
            })

        if normalize_yes(job.get("Req Freidora", "No")) and to_num(job["p Freidora seg"]) > 0:
            op_id = f"{job_id}_Freidora"
            prep_ops.append(op_id)
            rows.append({
                "OpID": op_id,
                "ID Pedido": job_id,
                "Canal": canal,
                "TipoPedido": tipo,
                "Etapa": "Preparación",
                "Estacion": "Freidora",
                "DuracionProceso": to_num(job["p Freidora seg"]),
                "rj": rj,
                "dj": dj,
                "wj": wj,
                "pj_total": pj_total,
                "Predecesores": [],
            })

        if normalize_yes(job.get("Req BebidaPostre", "No")) and to_num(job["p BebidaPostre seg"]) > 0:
            op_id = f"{job_id}_Bebidas_Postres"
            prep_ops.append(op_id)
            rows.append({
                "OpID": op_id,
                "ID Pedido": job_id,
                "Canal": canal,
                "TipoPedido": tipo,
                "Etapa": "Preparación",
                "Estacion": "Bebidas/Postres",
                "DuracionProceso": to_num(job["p BebidaPostre seg"]),
                "rj": rj,
                "dj": dj,
                "wj": wj,
                "pj_total": pj_total,
                "Predecesores": [],
            })

        ensamble_op = None

        if normalize_yes(job.get("Req Ensamble", "No")) and to_num(job["p Ensamble seg"]) > 0:
            ensamble_op = f"{job_id}_Ensamble"
            rows.append({
                "OpID": ensamble_op,
                "ID Pedido": job_id,
                "Canal": canal,
                "TipoPedido": tipo,
                "Etapa": "Ensamble",
                "Estacion": "Ensamble",
                "DuracionProceso": to_num(job["p Ensamble seg"]),
                "rj": rj,
                "dj": dj,
                "wj": wj,
                "pj_total": pj_total,
                "Predecesores": prep_ops.copy(),
            })

        if normalize_yes(job.get("Req Staging", "No")) and to_num(job["p Staging seg"]) > 0:
            staging_pred = [ensamble_op] if ensamble_op is not None else prep_ops.copy()

            rows.append({
                "OpID": f"{job_id}_Staging_Bolseo",
                "ID Pedido": job_id,
                "Canal": canal,
                "TipoPedido": tipo,
                "Etapa": "Staging/Bolseo",
                "Estacion": "Staging/Bolseo",
                "DuracionProceso": to_num(job["p Staging seg"]),
                "rj": rj,
                "dj": dj,
                "wj": wj,
                "pj_total": pj_total,
                "Predecesores": staging_pred,
            })

    operations = pd.DataFrame(rows)

    if operations.empty:
        raise ValueError("No se generaron operaciones. Revisa columnas Req y tiempos p.")

    return operations


# ============================================================
# CAPACIDADES Y CALENDARIOS DE RECURSOS
# ============================================================

def prepare_capacities(capacities: pd.DataFrame, horizon_seg: float = DEFAULT_HORIZON_SEG) -> pd.DataFrame:
    capacities = capacities.copy()

    if capacities.empty:
        capacities = pd.DataFrame({
            "Estacion": STATIONS,
            "Capacidad Inicial": [1, 1, 1, 1, 1],
            "Capacidad Durante Pico": [1, 1, 1, 1, 1],
            "Desde seg": [0, 0, 0, 0, 0],
            "Hasta seg": [horizon_seg] * len(STATIONS),
        })

    capacities["Estacion"] = capacities["Estacion"].apply(normalize_station)

    if "Capacidad Durante Pico" not in capacities.columns:
        capacities["Capacidad Durante Pico"] = capacities["Capacidad Inicial"]

    for col in ["Capacidad Inicial", "Capacidad Durante Pico", "Desde seg", "Hasta seg"]:
        if col not in capacities.columns:
            if col == "Desde seg":
                capacities[col] = 0
            elif col == "Hasta seg":
                capacities[col] = horizon_seg
            else:
                capacities[col] = 1
        capacities[col] = capacities[col].apply(to_num)

    capacities["Capacidad aplicada"] = np.where(
        capacities["Capacidad Durante Pico"].isna() | (capacities["Capacidad Durante Pico"] <= 0),
        capacities["Capacidad Inicial"],
        capacities["Capacidad Durante Pico"],
    )

    capacities["Capacidad aplicada"] = capacities["Capacidad aplicada"].clip(lower=1).round().astype(int)

    return capacities


def build_resource_calendar(capacities: pd.DataFrame, horizon_seg: float = DEFAULT_HORIZON_SEG) -> Dict[str, List[Tuple[float, float]]]:
    calendars = {}

    for station in STATIONS:
        cap_station = capacities[capacities["Estacion"] == station].copy()

        if cap_station.empty:
            cap_station = pd.DataFrame({
                "Estacion": [station],
                "Capacidad aplicada": [1],
                "Desde seg": [0],
                "Hasta seg": [horizon_seg],
            })

        max_cap = int(cap_station["Capacidad aplicada"].max())

        for k in range(1, max_cap + 1):
            resource = f"{station}_{k}"
            intervals = []

            cap_station = cap_station.sort_values("Desde seg").reset_index(drop=True)

            for _, row in cap_station.iterrows():
                cap = int(row["Capacidad aplicada"])
                start = to_num(row["Desde seg"])
                end = to_num(row["Hasta seg"])

                if k <= cap and end > start:
                    intervals.append((start, end))

            # Extensión del último estado de capacidad:
            # si el escenario termina en 3600/4800/5400 segundos pero aún hay cola,
            # el recurso no debe desaparecer. Se prolonga la última capacidad observada.
            if not cap_station.empty:
                last_row = cap_station.sort_values("Hasta seg").iloc[-1]
                last_end = to_num(last_row["Hasta seg"])
                last_cap = int(last_row["Capacidad aplicada"])

                if horizon_seg > last_end and k <= last_cap:
                    intervals.append((last_end, horizon_seg))

            calendars[resource] = merge_intervals(intervals)

    return calendars


def merge_intervals(intervals: List[Tuple[float, float]]) -> List[Tuple[float, float]]:
    if not intervals:
        return []

    intervals = sorted(intervals, key=lambda x: x[0])
    merged = [intervals[0]]

    for start, end in intervals[1:]:
        last_start, last_end = merged[-1]

        if start <= last_end:
            merged[-1] = (last_start, max(last_end, end))
        else:
            merged.append((start, end))

    return merged


def station_from_resource(resource: str) -> str:
    if resource.startswith("Bebidas/Postres"):
        return "Bebidas/Postres"
    if resource.startswith("Staging/Bolseo"):
        return "Staging/Bolseo"

    return resource.rsplit("_", 1)[0]


def next_feasible_start(
    resource: str,
    earliest_start: float,
    duration: float,
    calendars: Dict[str, List[Tuple[float, float]]],
) -> Optional[float]:
    intervals = calendars.get(resource, [])

    for start, end in intervals:
        candidate_start = max(earliest_start, start)

        if candidate_start + duration <= end:
            return candidate_start

    return None


# ============================================================
# REGLAS DE PRIORIDAD
# ============================================================

# Nota: la versión anterior incluía una función apply_rule() basada en DataFrames
# y una compute_remaining_processing() que sumaba todas las operaciones pendientes
# del pedido (incluyendo hermanas paralelas). Ambas se eliminaron porque el scheduler
# real opera con select_candidate() (que sigue la misma lógica que en R: SPT/WSPT por
# operación, CR con remanente, ATC con pbar fijo por estación y holgura sin truncar)
# y con downstream_processing() (que solo baja por sucesores reales para no contaminar
# el remanente con hermanas paralelas en preparación).


# ============================================================
# SCHEDULER DINÁMICO POR EVENTOS
# ============================================================


def schedule_operations(
    operations: pd.DataFrame,
    calendars: Dict[str, List[Tuple[float, float]]],
    rule: str,
    store_decisions: bool = True,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Scheduler dinámico por eventos, optimizado para poder calcular todas
    las instancias y reglas sin depender de filtros de pandas en cada iteración.
    """
    pending_ops = operations.to_dict("records")
    scheduled_rows = []
    decision_rows = []

    finish_times: Dict[str, float] = {}
    resource_available: Dict[str, float] = {}

    for resource, intervals in calendars.items():
        if intervals:
            resource_available[resource] = intervals[0][0]

    if not resource_available:
        raise ValueError("No hay recursos disponibles en los calendarios.")

    # pbar fijo por estación (no se recalcula en cada decisión).
    # Es el promedio de las duraciones de procesamiento de TODAS las
    # operaciones de la estación, calculado una sola vez. Coherente con
    # la versión final en R: estabiliza el ATC y lo hace reproducible.
    pbar_by_station: Dict[str, float] = {}
    for st_name in operations["Estacion"].unique():
        durations = operations.loc[operations["Estacion"] == st_name, "DuracionProceso"]
        mean_dur = float(durations.mean()) if len(durations) > 0 else 1.0
        pbar_by_station[str(st_name)] = max(mean_dur, 1.0)

    # Umbral de holgura crítica DATA-DRIVEN: mediana de las holguras iniciales
    # positivas (dj - rj - pj_total) de los pedidos de la instancia. Esto evita
    # un umbral arbitrario: un pedido se considera 'crítico recuperable' si tiene
    # menos margen que la mitad de los pedidos del turno. Se adapta a cada
    # instancia y es robusto frente a valores extremos.
    holguras_iniciales = []
    seen_jobs = set()
    for op in operations.to_dict("records"):
        job_key = str(op.get("ID Pedido"))
        if job_key in seen_jobs:
            continue
        seen_jobs.add(job_key)
        h_inicial = float(op["dj"]) - float(op["rj"]) - float(op["pj_total"])
        if h_inicial > 0:
            holguras_iniciales.append(h_inicial)
    umbral_holgura_critica = float(np.median(holguras_iniciales)) if holguras_iniciales else 60.0

    order_counter = 1

    # Mapa de sucesores para calcular el procesamiento remanente de forma
    # más fiel al flexible job shop: para una operación de preparación no se
    # suman sus operaciones paralelas hermanas, solo la operación actual y las
    # etapas posteriores que dependen de ella, como Ensamble y Staging/Bolseo.
    operations_by_id = {str(op["OpID"]): op for op in operations.to_dict("records")}
    successors_by_id: Dict[str, List[str]] = {op_id: [] for op_id in operations_by_id}

    for op_id, op in operations_by_id.items():
        for pred in op.get("Predecesores", []):
            pred = str(pred)
            if pred in successors_by_id:
                successors_by_id[pred].append(op_id)

    def downstream_processing(op_id: str, pending_ids: set) -> float:
        visited = set()
        stack = [str(op_id)]
        total = 0.0

        while stack:
            current = stack.pop()
            if current in visited:
                continue
            visited.add(current)

            if current in pending_ids and current in operations_by_id:
                total += float(operations_by_id[current].get("DuracionProceso", 0.0))

            for nxt in successors_by_id.get(current, []):
                if nxt not in visited:
                    stack.append(nxt)

        return max(total, float(operations_by_id[str(op_id)].get("DuracionProceso", 0.0)))

    def pred_status(op: Dict) -> Tuple[bool, float]:
        predecessors = op.get("Predecesores", [])
        pred_done = all(pred in finish_times for pred in predecessors)
        pred_finish = max(
            [finish_times[pred] for pred in predecessors if pred in finish_times],
            default=op["rj"],
        )
        return pred_done, pred_finish

    def add_priority_metrics(op: Dict, current_time: float, p_rem: float) -> Dict:
        row = dict(op)

        pij = max(float(row["DuracionProceso"]), 1.0)
        wj_safe = max(float(row["wj"]), 1.0)
        p_rem_safe = max(float(p_rem), 1.0)

        row["pij"] = pij
        row["wj_safe"] = wj_safe
        row["Pij_wj"] = pij / wj_safe
        row["CR"] = (float(row["dj"]) - current_time) / p_rem_safe
        row["Holgura"] = float(row["dj"]) - current_time - float(p_rem)

        return row

    def select_candidate(candidate_rows: List[Dict], current_time: float) -> List[Dict]:
        if not candidate_rows:
            return []

        # pbar fijo por estación (no se recalcula sobre los candidatos del momento).
        station_name = str(candidate_rows[0].get("Estacion", ""))
        p_bar = pbar_by_station.get(station_name, 1.0)
        k = 2.0

        for row in candidate_rows:
            # ATC ajustado a operación: dj de la operación = dj - trabajo posterior.
            row["TrabajoPosterior"] = max(float(row["P_remanente"]) - float(row["pij"]), 0.0)
            row["d_op"] = float(row["dj"]) - row["TrabajoPosterior"]
            # Holgura SIN truncar en 0: discrimina entre pedidos muy atrasados.
            row["Holgura_ATC"] = row["d_op"] - float(row["pij"]) - current_time
            # Tope numérico para que exp() no devuelva Inf con holguras muy negativas.
            exponente = -row["Holgura_ATC"] / (k * p_bar)
            exponente = min(700.0, exponente)
            row["ATC"] = (row["wj_safe"] / row["pij"]) * math.exp(exponente)

            if rule == "Holgura crítica + SPT":
                # Umbral data-driven (mediana de holguras iniciales positivas).
                if 0.0 <= row["Holgura"] <= umbral_holgura_critica:
                    row["GrupoPrioridad"] = 1
                    row["ClavePrioridad"] = row["Holgura"]
                elif row["Holgura"] > umbral_holgura_critica:
                    row["GrupoPrioridad"] = 2
                    row["ClavePrioridad"] = row["pij"]
                else:
                    row["GrupoPrioridad"] = 3
                    row["ClavePrioridad"] = row["pij"]

        if rule == "FIFO":
            return sorted(candidate_rows, key=lambda x: (x["rj"], x["ID Pedido"], x["OpID"]))

        if rule == "SPT":
            return sorted(candidate_rows, key=lambda x: (x["pij"], x["rj"], x["dj"], x["ID Pedido"]))

        if rule == "EDD":
            return sorted(candidate_rows, key=lambda x: (x["dj"], x["rj"], x["ID Pedido"]))

        if rule == "WSPT":
            return sorted(candidate_rows, key=lambda x: (x["Pij_wj"], x["rj"], x["dj"], x["ID Pedido"]))

        if rule == "CR":
            return sorted(candidate_rows, key=lambda x: (x["CR"], x["rj"], x["dj"], x["ID Pedido"]))

        if rule == "ATC":
            return sorted(candidate_rows, key=lambda x: (-x["ATC"], x["d_op"], x["rj"], x["ID Pedido"]))

        if rule == "Holgura crítica + SPT":
            return sorted(
                candidate_rows,
                key=lambda x: (
                    x["GrupoPrioridad"],
                    x["ClavePrioridad"],
                    x.get("Disponible_desde", 0.0),
                    x["dj"],
                    -x.get("wj_safe", 1.0),
                    x["ID Pedido"],
                ),
            )

        raise ValueError(f"Regla no reconocida: {rule}")

    while pending_ops:
        possible_decisions = []

        for resource, resource_time in resource_available.items():
            station = station_from_resource(resource)

            for op in pending_ops:
                if op["Estacion"] != station:
                    continue

                pred_done, pred_finish = pred_status(op)

                if not pred_done:
                    continue

                op_available_from = max(float(op["rj"]), float(pred_finish))
                earliest = max(float(resource_time), float(op_available_from))

                feasible_start = next_feasible_start(
                    resource=resource,
                    earliest_start=earliest,
                    duration=float(op["DuracionProceso"]),
                    calendars=calendars,
                )

                if feasible_start is not None:
                    possible_decisions.append({
                        "Resource": resource,
                        "Station": station,
                        "DecisionTime": float(feasible_start),
                    })

        if not possible_decisions:
            pending_debug = pd.DataFrame(pending_ops)
            raise ValueError(
                "No se pudo programar alguna operación. "
                "Revisa precedencias, capacidades o calendarios. "
                f"Operaciones pendientes: {len(pending_debug)}."
            )

        selected_decision = sorted(
            possible_decisions,
            key=lambda x: (x["DecisionTime"], x["Resource"]),
        )[0]

        decision_time = float(selected_decision["DecisionTime"])
        resource = str(selected_decision["Resource"])
        station = str(selected_decision["Station"])

        pending_ids = {str(op["OpID"]) for op in pending_ops}

        candidate_rows = []

        for op in pending_ops:
            if op["Estacion"] != station:
                continue

            pred_done, pred_finish = pred_status(op)

            if not pred_done:
                continue

            op_available_from = max(float(op["rj"]), float(pred_finish))

            if op_available_from <= decision_time + 1e-9:
                feasible_start = next_feasible_start(
                    resource=resource,
                    earliest_start=decision_time,
                    duration=float(op["DuracionProceso"]),
                    calendars=calendars,
                )

                if feasible_start is not None and abs(feasible_start - decision_time) <= 1e-9:
                    p_rem = downstream_processing(str(op["OpID"]), pending_ids)

                    op_dict = add_priority_metrics(op, decision_time, p_rem)
                    op_dict["Disponible_desde"] = op_available_from
                    op_dict["P_remanente"] = p_rem
                    candidate_rows.append(op_dict)

        if not candidate_rows:
            resource_available[resource] = decision_time + 1e-6
            continue

        ordered_candidates = select_candidate(candidate_rows, decision_time)
        selected = ordered_candidates[0]

        start = decision_time
        end = start + float(selected["DuracionProceso"])

        scheduled_row = {
            "Regla": rule,
            "Orden": order_counter,
            "OpID": selected["OpID"],
            "ID Pedido": selected["ID Pedido"],
            "Canal": selected["Canal"],
            "TipoPedido": selected["TipoPedido"],
            "Etapa": selected["Etapa"],
            "Estacion": station,
            "Recurso": resource,
            "Inicio": start,
            "InicioProceso": start,
            "Fin": end,
            "Pij": float(selected["DuracionProceso"]),
            "Duracion": float(selected["DuracionProceso"]),
            "rj": float(selected["rj"]),
            "dj": float(selected["dj"]),
            "wj": float(selected["wj"]),
            "pj": float(selected["pj_total"]),
        }

        scheduled_rows.append(scheduled_row)

        if store_decisions:
            for row in ordered_candidates:
                decision_row = dict(row)
                decision_row["Regla"] = rule
                decision_row["OrdenDecision"] = order_counter
                decision_row["TiempoDecision"] = decision_time
                decision_row["Recurso"] = resource
                decision_row["Seleccionado"] = "Sí" if row["OpID"] == selected["OpID"] else "No"
                decision_rows.append(decision_row)

        finish_times[selected["OpID"]] = end
        resource_available[resource] = end

        selected_opid = selected["OpID"]
        pending_ops = [op for op in pending_ops if op["OpID"] != selected_opid]

        order_counter += 1

    schedule = pd.DataFrame(scheduled_rows)
    decisions = pd.DataFrame(decision_rows) if decision_rows else pd.DataFrame()

    return schedule, decisions



# ============================================================
# MÉTRICAS
# ============================================================

def calculate_results(schedule: pd.DataFrame, jobs: pd.DataFrame) -> pd.DataFrame:
    completions = (
        schedule
        .groupby("ID Pedido", as_index=False)
        .agg(
            Cj=("Fin", "max"),
            InicioPrimeraOperacion=("Inicio", "min"),
        )
    )

    base = jobs[[
        "ID Pedido",
        "Canal",
        "Tipo Pedido",
        "r j seg",
        "d j seg",
        "w j",
        "Tiempo total procesamiento",
    ]].copy()

    base = base.rename(columns={
        "Tipo Pedido": "TipoPedido",
        "r j seg": "rj",
        "d j seg": "dj",
        "w j": "wj",
        "Tiempo total procesamiento": "pj",
    })

    results = base.merge(completions, on="ID Pedido", how="left")

    results["Cj"] = results["Cj"].fillna(0)
    results["Fj"] = results["Cj"] - results["rj"]
    results["Lj"] = results["Cj"] - results["dj"]
    results["Tj"] = results["Lj"].clip(lower=0)
    results["Uj"] = (results["Tj"] > 0).astype(int)
    results["w_j C_j"] = results["wj"] * results["Cj"]
    results["w_j T_j"] = results["wj"] * results["Tj"]
    results["Margen"] = results["dj"] - results["Cj"]

    return results.sort_values("ID Pedido").reset_index(drop=True)


def summarize_rule(rule: str, results: pd.DataFrame) -> Dict:
    return {
        "Regla": rule,
        "Cmax": results["Cj"].max(),
        "∑Cj": results["Cj"].sum(),
        "∑Fj": results["Fj"].sum(),
        "∑Uj": results["Uj"].sum(),
        "Tmax": results["Tj"].max(),
        "Lmax": results["Lj"].max(),
        "∑Tj": results["Tj"].sum(),
        "∑Lj": results["Lj"].sum(),
        "∑wjCj": results["w_j C_j"].sum(),
        "∑wjTj": results["w_j T_j"].sum(),
        "Cumplimiento SLA (%)": (results["Uj"].eq(0).mean() * 100),
    }


def build_comparison(all_results: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    rows = []

    for rule, results in all_results.items():
        rows.append(summarize_rule(rule, results))

    comparison = pd.DataFrame(rows)

    numeric_cols = comparison.select_dtypes(include=[np.number]).columns
    comparison[numeric_cols] = comparison[numeric_cols].round(2)

    return comparison


def build_ranking(comparison: pd.DataFrame) -> pd.DataFrame:
    metrics = [
        ("Menor Cmax", "Cmax", False),
        ("Menor ∑Cj", "∑Cj", False),
        ("Menor ∑Fj", "∑Fj", False),
        ("Menor ∑Uj", "∑Uj", False),
        ("Menor Tmax", "Tmax", False),
        ("Menor Lmax", "Lmax", False),
        ("Menor ∑Tj", "∑Tj", False),
        ("Menor ∑Lj", "∑Lj", False),
        ("Menor ∑wjCj", "∑wjCj", False),
        ("Menor ∑wjTj", "∑wjTj", False),
        ("Mayor cumplimiento SLA", "Cumplimiento SLA (%)", True),
    ]

    rows = []

    for label, col, higher_is_better in metrics:
        if higher_is_better:
            best_value = comparison[col].max()
        else:
            best_value = comparison[col].min()

        winners = comparison[np.isclose(comparison[col], best_value)]["Regla"].tolist()

        rows.append({
            "Indicador destacado": label,
            "Regla destacada": " | ".join(winners),
            "Valor obtenido": round(best_value, 2),
            "Lectura": "Empate" if len(winners) > 1 else "Mejor valor único",
        })

    return pd.DataFrame(rows)


def build_index(comparison: pd.DataFrame) -> pd.DataFrame:
    """
    Índice comparativo normalizado (min-max) con orientación 'menor es mejor'.

    Cada indicador se normaliza entre 0 y 1: el mejor desempeño observado queda
    en 0 y el peor en 1. El cumplimiento de SLA se invierte (100 - SLA%) para que
    también apunte en la dirección 'menor es mejor'. El índice final promedia los
    cinco indicadores normalizados y representa la distancia relativa de cada
    regla frente al mejor desempeño observado.
    """
    def norm_menor_mejor(x):
        x = np.asarray(x, dtype=float)
        rango = x.max() - x.min()
        if rango == 0:
            return np.zeros_like(x)
        return (x - x.min()) / rango

    df = comparison.copy()
    df["n_Cmax"] = norm_menor_mejor(df["Cmax"])
    df["n_∑Tj"] = norm_menor_mejor(df["∑Tj"])
    df["n_∑Uj"] = norm_menor_mejor(df["∑Uj"])
    df["n_∑wjTj"] = norm_menor_mejor(df["∑wjTj"])
    df["n_SLA"] = norm_menor_mejor(100 - df["Cumplimiento SLA (%)"])  # invertido

    df["Índice comparativo"] = (
        (df["n_Cmax"] + df["n_∑Tj"] + df["n_∑Uj"] + df["n_∑wjTj"] + df["n_SLA"]) / 5
    ).round(4)

    out = df[[
        "Regla", "Cmax", "∑Tj", "∑Uj", "∑wjTj", "Cumplimiento SLA (%)",
        "Índice comparativo",
    ]].sort_values("Índice comparativo").reset_index(drop=True)

    return out


# ============================================================
# CUELLO DE BOTELLA
# ============================================================

def preliminary_bottleneck(jobs: pd.DataFrame, capacities: pd.DataFrame) -> pd.DataFrame:
    load = pd.DataFrame({
        "Estacion": STATIONS,
        "Carga total seg": [
            jobs["p Parrilla seg"].sum(),
            jobs["p Freidora seg"].sum(),
            jobs["p BebidaPostre seg"].sum(),
            jobs["p Ensamble seg"].sum(),
            jobs["p Staging seg"].sum(),
        ],
    })

    cap = capacities.copy()
    cap["Duracion intervalo"] = cap["Hasta seg"] - cap["Desde seg"]
    cap["Capacidad acumulada"] = cap["Capacidad aplicada"] * cap["Duracion intervalo"]

    cap_summary = (
        cap.groupby("Estacion", as_index=False)
        .agg(
            Capacidad_inicial=("Capacidad Inicial", "first"),
            Capacidad_minima=("Capacidad aplicada", "min"),
            Capacidad_maxima=("Capacidad aplicada", "max"),
            Tiempo_total=("Duracion intervalo", "sum"),
            Capacidad_acumulada=("Capacidad acumulada", "sum"),
        )
    )

    cap_summary["Capacidad promedio disponible"] = (
        cap_summary["Capacidad_acumulada"] / cap_summary["Tiempo_total"]
    )

    table = load.merge(cap_summary, on="Estacion", how="left")
    table["Capacidad promedio disponible"] = table["Capacidad promedio disponible"].fillna(1).clip(lower=1)
    table["Carga por capacidad"] = table["Carga total seg"] / table["Capacidad promedio disponible"]
    table["Presión relativa"] = table["Carga por capacidad"] / table["Carga por capacidad"].max()

    table["Lectura operativa"] = np.where(
        np.isclose(table["Carga por capacidad"], table["Carga por capacidad"].max()),
        "Cuello de botella preliminar",
        np.where(table["Presión relativa"] >= 0.85, "Alta presión relativa", "Presión controlada"),
    )

    numeric_cols = table.select_dtypes(include=[np.number]).columns
    table[numeric_cols] = table[numeric_cols].round(2)

    return table.sort_values("Carga por capacidad", ascending=False).reset_index(drop=True)


def real_bottleneck(schedule: pd.DataFrame, capacities: pd.DataFrame, rule: str) -> pd.DataFrame:
    """
    Cuello de botella REAL post-simulación, basado en saturación:

        Saturación = Carga real ejecutada / Segundos-máquina realmente disponibles

    A diferencia del cuello preliminar (estimación teórica con capacidad promedio),
    este cálculo usa los tiempos REALES de las operaciones programadas y los
    segundos-máquina exactos que aporta el calendario por tramos de la estación.
    Coherente con el scheduler v3 (calendario por recurso + no-span).
    """
    if schedule is None or schedule.empty:
        return pd.DataFrame()

    # Segundos-máquina realmente disponibles por estación (sumando todos los tramos).
    cap = capacities.copy()
    cap["DuracionIntervalo"] = cap["Hasta seg"] - cap["Desde seg"]
    cap["SegundosMaquina"] = cap["Capacidad aplicada"] * cap["DuracionIntervalo"]
    disponibilidad = (
        cap.groupby("Estacion", as_index=False)["SegundosMaquina"]
        .sum()
        .rename(columns={"SegundosMaquina": "Segundos maquina disponibles"})
    )

    # Carga real ejecutada por estación (suma de duraciones reales de operaciones).
    sched = schedule.copy()
    sched["CargaProceso"] = sched["Fin"] - sched["Inicio"]
    carga_real = (
        sched.groupby("Estacion", as_index=False)
        .agg(**{
            "Carga total": ("CargaProceso", "sum"),
            "Recursos usados": ("Recurso", "nunique"),
        })
    )

    table = (
        pd.DataFrame({"Estacion": STATIONS})
        .merge(carga_real, on="Estacion", how="left")
        .merge(disponibilidad, on="Estacion", how="left")
        .fillna({"Carga total": 0, "Recursos usados": 0, "Segundos maquina disponibles": 0})
    )

    # Saturación real y carga por recurso.
    table["Carga por recurso"] = np.where(
        table["Recursos usados"] > 0,
        table["Carga total"] / table["Recursos usados"],
        0,
    )
    table["Saturación"] = np.where(
        table["Segundos maquina disponibles"] > 0,
        table["Carga total"] / table["Segundos maquina disponibles"],
        0,
    )

    sat_max = table["Saturación"].max()
    table["Presión relativa"] = table["Saturación"] / sat_max if sat_max > 0 else 0

    table["Lectura operativa"] = np.where(
        np.isclose(table["Saturación"], sat_max) & (sat_max > 0),
        f"Cuello de botella {rule}",
        np.where(table["Presión relativa"] >= 0.85, "Alta presión relativa", "Presión controlada"),
    )

    numeric_cols = table.select_dtypes(include=[np.number]).columns
    table[numeric_cols] = table[numeric_cols].round(4)

    return table.sort_values(
        ["Saturación", "Carga por recurso"], ascending=[False, False]
    ).reset_index(drop=True)


# ============================================================
# VALIDACIONES
# ============================================================

def validate_no_overlap(schedule: pd.DataFrame) -> pd.DataFrame:
    violations = []

    for resource, group in schedule.groupby("Recurso"):
        group = group.sort_values("Inicio").reset_index(drop=True)

        for i in range(len(group) - 1):
            current_end = group.loc[i, "Fin"]
            next_start = group.loc[i + 1, "Inicio"]

            if next_start < current_end - 1e-9:
                violations.append({
                    "Tipo": "Solapamiento",
                    "Recurso": resource,
                    "OpID 1": group.loc[i, "OpID"],
                    "OpID 2": group.loc[i + 1, "OpID"],
                    "Fin 1": current_end,
                    "Inicio 2": next_start,
                })

    return pd.DataFrame(violations)


def validate_release_times(schedule: pd.DataFrame) -> pd.DataFrame:
    invalid = schedule[schedule["Inicio"] < schedule["rj"] - 1e-9].copy()

    if invalid.empty:
        return pd.DataFrame()

    invalid["Tipo"] = "Inicio antes de rj"
    return invalid[["Tipo", "OpID", "ID Pedido", "Inicio", "rj"]]


def validate_resource_calendar(
    schedule: pd.DataFrame,
    calendars: Dict[str, List[Tuple[float, float]]],
) -> pd.DataFrame:
    violations = []

    for _, row in schedule.iterrows():
        resource = row["Recurso"]
        start = row["Inicio"]
        end = row["Fin"]

        valid = False

        for a, b in calendars.get(resource, []):
            if start >= a - 1e-9 and end <= b + 1e-9:
                valid = True
                break

        if not valid:
            violations.append({
                "Tipo": "Violación calendario recurso",
                "OpID": row["OpID"],
                "ID Pedido": row["ID Pedido"],
                "Recurso": resource,
                "Inicio": start,
                "Fin": end,
                "Calendario válido": calendars.get(resource, []),
            })

    return pd.DataFrame(violations)


def validate_precedence(schedule: pd.DataFrame, operations: pd.DataFrame) -> pd.DataFrame:
    finish = dict(zip(schedule["OpID"], schedule["Fin"]))
    start = dict(zip(schedule["OpID"], schedule["Inicio"]))

    violations = []

    for _, op in operations.iterrows():
        op_id = op["OpID"]

        for pred in op["Predecesores"]:
            if pred not in finish or op_id not in start:
                continue

            if start[op_id] < finish[pred] - 1e-9:
                violations.append({
                    "Tipo": "Violación precedencia",
                    "OpID": op_id,
                    "Predecesor": pred,
                    "Inicio operación": start[op_id],
                    "Fin predecesor": finish[pred],
                })

    return pd.DataFrame(violations)


def validate_schedule(
    schedule: pd.DataFrame,
    operations: pd.DataFrame,
    calendars: Dict[str, List[Tuple[float, float]]],
) -> pd.DataFrame:
    validations = [
        validate_no_overlap(schedule),
        validate_release_times(schedule),
        validate_resource_calendar(schedule, calendars),
        validate_precedence(schedule, operations),
    ]

    validations = [v for v in validations if not v.empty]

    if not validations:
        return pd.DataFrame({
            "Resultado": ["OK"],
            "Detalle": ["No se detectaron violaciones de solapamiento, rj, precedencia o calendario."],
        })

    return pd.concat(validations, ignore_index=True)


# ============================================================
# GRÁFICAS
# ============================================================

def resource_sort_key(resource: str) -> Tuple[int, int, str]:
    """
    Ordena los recursos por la secuencia operativa de estaciones y luego
    por el número del recurso dentro de cada estación.
    """
    station = station_from_resource(resource)
    station_idx = STATIONS.index(station) if station in STATIONS else len(STATIONS)

    try:
        resource_num = int(str(resource).rsplit("_", 1)[1])
    except Exception:
        resource_num = 1

    return station_idx, resource_num, str(resource)


def plot_gantt(schedule: pd.DataFrame, rule: str):
    """
    Gantt 100% numérico en segundos.
    No se convierte a datetime y no se usa px.timeline, por eso el eje X
    no muestra meses, años ni horas calendario.
    """
    data = schedule.copy()

    data["Inicio"] = pd.to_numeric(data["Inicio"], errors="coerce").fillna(0.0)
    data["Fin"] = pd.to_numeric(data["Fin"], errors="coerce").fillna(0.0)
    data["Duracion"] = pd.to_numeric(data["Duracion"], errors="coerce").fillna(
        data["Fin"] - data["Inicio"]
    )

    data["Inicio_txt"] = data["Inicio"].round(0).astype(int).astype(str) + " s"
    data["Fin_txt"] = data["Fin"].round(0).astype(int).astype(str) + " s"
    data["Duracion_txt"] = data["Duracion"].round(0).astype(int).astype(str) + " s"
    # Texto visible dentro de cada barra: nombre/ID del trabajo.
    # Los tiempos quedan en las etiquetas emergentes del hover.
    data["EtiquetaTrabajo"] = data["ID Pedido"].astype(str)

    resource_order = sorted(data["Recurso"].unique().tolist(), key=resource_sort_key)

    max_finish = float(data["Fin"].max()) if not data.empty else 0.0
    tick_step = 300 if max_finish <= 4200 else 600
    tick_end = int(np.ceil(max_finish / tick_step) * tick_step) if max_finish > 0 else tick_step
    tick_vals = list(range(0, tick_end + tick_step, tick_step))
    tick_text = [f"{v} s" for v in tick_vals]

    fig = px.bar(
        data,
        x="Duracion",
        y="Recurso",
        base="Inicio",
        orientation="h",
        color="ID Pedido",
        text="EtiquetaTrabajo",
        category_orders={"Recurso": resource_order},
        custom_data=[
            "OpID",
            "ID Pedido",
            "Estacion",
            "Inicio_txt",
            "Fin_txt",
            "Duracion_txt",
            "rj",
            "dj",
            "wj",
        ],
        title=f"Diagrama de Gantt en segundos - {rule}",
    )

    fig.update_traces(
        textposition="inside",
        insidetextanchor="middle",
        textfont_size=10,
        marker_line_width=0.5,
        cliponaxis=False,
        hovertemplate=(
            "<b>%{customdata[0]}</b><br>"
            "Pedido: %{customdata[1]}<br>"
            "Estación: %{customdata[2]}<br>"
            "Inicio: %{customdata[3]}<br>"
            "Fin: %{customdata[4]}<br>"
            "Duración: %{customdata[5]}<br>"
            "rj: %{customdata[6]:.0f} s<br>"
            "dj: %{customdata[7]:.0f} s<br>"
            "wj: %{customdata[8]:.2f}"
            "<extra></extra>"
        ),
    )

    fig.update_yaxes(
        categoryorder="array",
        categoryarray=resource_order,
        autorange="reversed",
        title="Recursos ordenados por estación",
    )

    fig.update_xaxes(
        title="Tiempo operativo transcurrido (segundos)",
        type="linear",
        tickmode="array",
        tickvals=tick_vals,
        ticktext=tick_text,
        range=[0, max(tick_end, 1)],
        showgrid=True,
        zeroline=True,
    )

    fig.update_layout(
        height=max(600, 85 + 62 * len(resource_order)),
        showlegend=False,
        bargap=0.20,
        margin=dict(l=165, r=45, t=80, b=80),
        xaxis_tickangle=0,
    )

    return fig


def plot_tardiness(results: pd.DataFrame, rule: str):
    tardy = results[results["Tj"] > 0].copy()

    if tardy.empty:
        return None

    tardy = tardy.sort_values(["Tj", "w_j T_j"], ascending=[False, False])

    fig = px.bar(
        tardy,
        x="Tj",
        y="ID Pedido",
        orientation="h",
        color="w_j T_j",
        hover_data=["TipoPedido", "Cj", "dj", "Lj", "Tj", "wj", "w_j T_j"],
        title=f"Pedidos tardíos y tardanza - {rule}",
    )

    fig.update_layout(
        height=520,
        xaxis_title="Tardanza Tj (segundos)",
        yaxis_title="Pedido",
        margin=dict(l=110, r=30, t=70, b=50),
    )

    return fig


def plot_comparison(comparison: pd.DataFrame):
    data = comparison[["Regla", "∑Uj", "∑Tj"]].melt(
        id_vars="Regla",
        var_name="Indicador",
        value_name="Valor",
    )

    fig = px.bar(
        data,
        x="Regla",
        y="Valor",
        color="Indicador",
        barmode="group",
        text="Valor",
        title="Comparación de pedidos tardíos y tardanza total",
    )

    fig.update_layout(
        height=520,
        xaxis_title="Regla",
        yaxis_title="Valor",
        margin=dict(l=70, r=30, t=70, b=90),
    )

    return fig


def plot_cmax(comparison: pd.DataFrame):
    data = comparison.sort_values("Cmax").copy()

    fig = px.bar(
        data,
        x="Cmax",
        y="Regla",
        orientation="h",
        text="Cmax",
        title="Comparación de Cmax por regla",
    )

    fig.update_layout(
        height=480,
        xaxis_title="Cmax",
        yaxis_title="Regla",
        margin=dict(l=160, r=30, t=70, b=50),
    )

    return fig


# ============================================================
# CÓDIGO R EXPORTABLE
# ============================================================

def dataframe_to_r_tribble(df: pd.DataFrame, object_name: str = "comparacion") -> str:
    cols = list(df.columns)

    r = []
    r.append("library(tibble)")
    r.append("")
    r.append(f"{object_name} <- tribble(")
    r.append("  " + ", ".join([f"~`{c}`" for c in cols]) + ",")

    for i, row in df.iterrows():
        values = []

        for c in cols:
            value = row[c]

            if isinstance(value, str):
                value = value.replace("\\", "\\\\").replace('"', '\\"')
                values.append(f'"{value}"')
            elif pd.isna(value):
                values.append("NA")
            else:
                values.append(str(value))

        line = "  " + ", ".join(values)

        if i < len(df) - 1:
            line += ","

        r.append(line)

    r.append(")")
    return "\n".join(r)


def make_excel_download(dfs: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in dfs.items():
            clean_sheet = sheet_name[:31]
            df.to_excel(writer, sheet_name=clean_sheet, index=False)

    return output.getvalue()


# ============================================================
# HELPERS DE EJECUCIÓN POR INSTANCIA
# ============================================================

def compute_instance(
    excel_source,
    instance_number: int,
    selected_rules: List[str],
    store_decisions: bool,
) -> Dict:
    """
    Ejecuta una instancia completa (carga, prepara, programa todas las reglas,
    calcula métricas y cuellos). Devuelve un diccionario con todos los artefactos
    para que la UI o el agregador maestro los consuman.
    """
    jobs_raw, events, capacities_raw, sheet_name = read_instance_excel(
        excel_source, int(instance_number)
    )
    jobs = prepare_jobs(jobs_raw)
    horizon_seg = calculate_horizon_seg(jobs, capacities_raw)
    capacities = prepare_capacities(capacities_raw, horizon_seg=horizon_seg)
    operations = build_operations(jobs)
    calendars = build_resource_calendar(capacities, horizon_seg=horizon_seg)
    bottleneck_table = preliminary_bottleneck(jobs, capacities)
    calendar_df = make_calendar_table(calendars)

    # Umbral data-driven para mostrarlo en UI (mismo que usa el scheduler).
    holguras_init = jobs["d j seg"] - jobs["r j seg"] - jobs["Tiempo total procesamiento"]
    holguras_pos = holguras_init[holguras_init > 0]
    umbral_holgura = float(holguras_pos.median()) if len(holguras_pos) > 0 else 60.0

    all_schedules: Dict[str, pd.DataFrame] = {}
    all_results: Dict[str, pd.DataFrame] = {}
    all_decisions: Dict[str, pd.DataFrame] = {}
    real_bottlenecks: Dict[str, pd.DataFrame] = {}
    errors: List[Dict] = []

    for rule in selected_rules:
        try:
            schedule, decisions = schedule_operations(
                operations=operations, calendars=calendars,
                rule=rule, store_decisions=store_decisions,
            )
            results = calculate_results(schedule, jobs)
            all_schedules[rule] = schedule
            all_results[rule] = results
            all_decisions[rule] = decisions
            real_bottlenecks[rule] = real_bottleneck(schedule, capacities, rule)
        except Exception as rule_error:
            errors.append({
                "Instancia": instance_number, "Regla": rule, "Error": str(rule_error),
            })

    if not all_results:
        raise ValueError(
            f"No se pudo calcular ninguna regla para la Instancia {instance_number}."
        )

    comparison = build_comparison(all_results)
    ranking = build_ranking(comparison)
    index_table = build_index(comparison)

    return {
        "instance_number": instance_number,
        "sheet_name": sheet_name,
        "jobs": jobs,
        "events": events,
        "capacities": capacities,
        "operations": operations,
        "calendars": calendars,
        "horizon_seg": horizon_seg,
        "bottleneck_table": bottleneck_table,
        "calendar_df": calendar_df,
        "umbral_holgura": umbral_holgura,
        "all_schedules": all_schedules,
        "all_results": all_results,
        "all_decisions": all_decisions,
        "real_bottlenecks": real_bottlenecks,
        "comparison": comparison,
        "ranking": ranking,
        "index_table": index_table,
        "errors": errors,
    }


def render_instance_block(data: Dict, show_decisions: bool, excel_label: str):
    """
    Renderiza el bloque completo de una instancia (datos, comparación,
    resultados por regla y descarga). Mismo layout que el modo de una sola.
    """
    instance_number = data["instance_number"]
    sheet_name = data["sheet_name"]
    jobs = data["jobs"]
    events = data["events"]
    capacities = data["capacities"]
    operations = data["operations"]
    horizon_seg = data["horizon_seg"]
    bottleneck_table = data["bottleneck_table"]
    calendar_df = data["calendar_df"]
    umbral_holgura = data["umbral_holgura"]
    all_schedules = data["all_schedules"]
    all_results = data["all_results"]
    all_decisions = data["all_decisions"]
    real_bottlenecks = data["real_bottlenecks"]
    comparison = data["comparison"]
    ranking = data["ranking"]
    index_table = data["index_table"]
    errors = data["errors"]

    st.success(
        f"Instancia {instance_number} calculada desde la hoja: {sheet_name} | Archivo: {excel_label}"
    )

    # 1. Datos
    st.header("1. Datos de la instancia seleccionada")
    st.caption(
        f"Horizonte de calendario usado: {horizon_seg:.0f} segundos. "
        f"Umbral de holgura crítica (data-driven, mediana de holguras iniciales positivas): "
        f"**{umbral_holgura:.0f} seg**."
    )

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.metric("Pedidos", len(jobs))
    with c2:
        st.metric("Operaciones", len(operations))
    with c3:
        st.metric("Carga total", f"{jobs['Tiempo total procesamiento'].sum():.0f} seg")
    with c4:
        ventana_llegada = jobs["r j seg"].max() - jobs["r j seg"].min()
        st.metric("Ventana llegada", f"{ventana_llegada:.0f} seg")
    with c5:
        dominant_channel = jobs["Canal"].value_counts().idxmax()
        st.metric("Canal dominante", dominant_channel)

    with st.expander("Ver pedidos originales preparados", expanded=False):
        st.dataframe(jobs, width="stretch")

    col_l, col_r = st.columns(2)
    with col_l:
        st.subheader("Capacidades")
        st.dataframe(capacities, width="stretch")
    with col_r:
        st.subheader("Eventos")
        if events.empty:
            st.info("No hay eventos registrados para esta instancia.")
        else:
            st.dataframe(events, width="stretch")

    with st.expander("Ver operaciones generadas", expanded=False):
        st.dataframe(operations, width="stretch")

    col_l, col_r = st.columns(2)
    with col_l:
        st.subheader("Cuello de botella preliminar")
        st.caption(
            "Diagnóstico **teórico** previo a programar. Usa capacidad promedio. "
            "La programación NO usa este promedio: opera con calendario por recurso y no-span. "
            "El cuello **real** por saturación se muestra dentro de cada regla."
        )
        st.dataframe(bottleneck_table, width="stretch")
    with col_r:
        st.subheader("Calendario de recursos")
        st.dataframe(calendar_df, width="stretch")

    # 2. Comparación
    st.header("2. Comparación de reglas")
    st.dataframe(comparison, width="stretch")

    col_l, col_r = st.columns(2)
    with col_l:
        st.plotly_chart(plot_comparison(comparison), width="stretch", key=f"cmp_{instance_number}")
    with col_r:
        st.plotly_chart(plot_cmax(comparison), width="stretch", key=f"cmax_{instance_number}")

    st.subheader("Ranking por indicador")
    st.dataframe(ranking, width="stretch")

    st.subheader("Índice comparativo normalizado")
    st.caption(
        "Min-max por indicador; SLA invertido. Menor es mejor: distancia frente al mejor desempeño observado."
    )
    st.dataframe(index_table, width="stretch")

    if errors:
        st.warning("Algunas reglas no pudieron calcularse.")
        st.dataframe(pd.DataFrame(errors), width="stretch")

    # 3. Resultados por regla
    st.header("3. Resultados por regla")
    rule_tabs = st.tabs(list(all_schedules.keys()))
    for rule_tab, rule in zip(rule_tabs, all_schedules.keys()):
        with rule_tab:
            st.subheader(rule)
            st.markdown("### Gantt ordenado por estación")
            st.plotly_chart(
                plot_gantt(all_schedules[rule], rule),
                width="stretch", key=f"gantt_{instance_number}_{rule}",
            )
            st.markdown("### Operaciones programadas")
            st.dataframe(all_schedules[rule], width="stretch")
            st.markdown("### Resultados por pedido")
            st.dataframe(all_results[rule], width="stretch")
            st.markdown("### Cuello de botella real (saturación post-simulación)")
            st.caption("Saturación = carga real / segundos-máquina realmente disponibles.")
            real_cb = real_bottlenecks.get(rule, pd.DataFrame())
            if real_cb.empty:
                st.info("No hay operaciones programadas para calcular saturación.")
            else:
                st.dataframe(real_cb, width="stretch")
            st.markdown("### Pedidos tardíos")
            tardy_fig = plot_tardiness(all_results[rule], rule)
            if tardy_fig is None:
                st.info(f"Bajo {rule} no se registran pedidos tardíos.")
            else:
                st.plotly_chart(
                    tardy_fig, width="stretch",
                    key=f"tardy_{instance_number}_{rule}",
                )
            with st.expander("Ver tabla de decisiones"):
                if all_decisions[rule].empty:
                    st.info("Tabla de decisiones desactivada o vacía.")
                else:
                    st.dataframe(all_decisions[rule], width="stretch")

    # 4. Descarga por instancia
    st.header("4. Descargar resultados de esta instancia")
    output_dfs = {
        "comparacion": comparison,
        "ranking": ranking,
        "indice_comparativo": index_table,
        "pedidos": jobs,
        "operaciones_generadas": operations,
        "capacidades": capacities,
        "eventos": events if not events.empty else pd.DataFrame({"Detalle": ["Sin eventos"]}),
        "cuello_preliminar": bottleneck_table,
        "calendario_recursos": calendar_df,
    }
    if errors:
        output_dfs["errores"] = pd.DataFrame(errors)
    for rule in all_schedules:
        alias = RULE_ALIASES.get(rule, rule[:8])
        output_dfs[clean_sheet_name(f"{alias}_ops")] = all_schedules[rule]
        output_dfs[clean_sheet_name(f"{alias}_resultados")] = all_results[rule]
        if rule in real_bottlenecks and not real_bottlenecks[rule].empty:
            output_dfs[clean_sheet_name(f"{alias}_cuello_real")] = real_bottlenecks[rule]
        if show_decisions and not all_decisions[rule].empty:
            output_dfs[clean_sheet_name(f"{alias}_decisiones")] = all_decisions[rule]

    excel_bytes = make_excel_download(output_dfs)
    st.download_button(
        label=f"Descargar resultados Instancia {instance_number} (Excel)",
        data=excel_bytes,
        file_name=f"resultados_instancia_{instance_number}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_{instance_number}",
    )


def build_master_summary(all_instances_data: Dict[int, Dict]) -> pd.DataFrame:
    """
    Construye la tabla maestra con una fila por (instancia, regla)
    para comparar el desempeño de todas las reglas en todas las instancias.
    """
    rows = []
    for inst_num, data in all_instances_data.items():
        comp = data["comparison"]
        idx_t = data["index_table"]
        idx_map = dict(zip(idx_t["Regla"], idx_t["Índice comparativo"]))
        for _, r in comp.iterrows():
            rows.append({
                "Instancia": inst_num,
                "Regla": r["Regla"],
                "Cmax": r["Cmax"],
                "∑Tj": r["∑Tj"],
                "∑Uj": r["∑Uj"],
                "∑wjTj": r["∑wjTj"],
                "Cumplimiento SLA (%)": r["Cumplimiento SLA (%)"],
                "Índice comparativo": idx_map.get(r["Regla"], np.nan),
            })
    return pd.DataFrame(rows).sort_values(
        ["Instancia", "Índice comparativo"]
    ).reset_index(drop=True)


def build_winners_table(master: pd.DataFrame) -> pd.DataFrame:
    """
    Para cada instancia, identifica la regla ganadora por cada indicador clave
    (la de menor Cmax, menor ∑Tj, etc.) y la mejor regla por índice global.
    """
    rows = []
    for inst_num, grp in master.groupby("Instancia"):
        def best(col, lower=True):
            ext = grp[col].min() if lower else grp[col].max()
            wins = grp[np.isclose(grp[col], ext)]["Regla"].tolist()
            return " | ".join(wins), round(float(ext), 2)

        cmax_r, cmax_v = best("Cmax")
        tj_r, tj_v = best("∑Tj")
        uj_r, uj_v = best("∑Uj")
        wjtj_r, wjtj_v = best("∑wjTj")
        sla_r, sla_v = best("Cumplimiento SLA (%)", lower=False)
        idx_r, idx_v = best("Índice comparativo")

        rows.append({
            "Instancia": inst_num,
            "Mejor Cmax": f"{cmax_r} ({cmax_v})",
            "Mejor ∑Tj": f"{tj_r} ({tj_v})",
            "Mejor ∑Uj": f"{uj_r} ({uj_v})",
            "Mejor ∑wjTj": f"{wjtj_r} ({wjtj_v})",
            "Mejor SLA %": f"{sla_r} ({sla_v})",
            "Mejor según índice": f"{idx_r} ({idx_v})",
        })
    return pd.DataFrame(rows)


# ============================================================
# STREAMLIT APP
# ============================================================

st.set_page_config(
    page_title="Proyecto Logística de Producción | Caso McDonald's",
    layout="wide",
)

st.title("Proyecto de Logística de Producción: Caso McDonald's")
st.caption(
    "Programación de pedidos en un flexible job shop, comparación de reglas de despacho "
    "y visualización del Gantt en segundos por estación."
)

st.markdown(
    """
    Puedes trabajar de dos formas: subir un Excel nuevo con la misma estructura del proyecto,
    o dejar vacío el cargador para usar automáticamente el archivo local del proyecto.
    La app detecta las hojas de instancia disponibles y permite seleccionar la que quieras resolver.
    """
)

uploaded_excel = st.file_uploader(
    "Sube un Excel de configuración compatible",
    type=["xlsx"],
    help=(
        "Debe tener hojas de instancia como 'Instancia 1', 'Instancia_1' o '1'. "
        "Opcionalmente puede incluir hojas 'Capacidades' y 'Eventos' con columna Instancia."
    ),
)

if uploaded_excel is not None:
    excel_source = uploaded_excel.getvalue()
    excel_label = uploaded_excel.name
    st.success(f"Excel cargado desde archivo subido: {excel_label}")
else:
    try:
        excel_path = find_project_excel()
        excel_source = excel_path
        excel_label = excel_path.name
        st.info(f"No subiste un Excel; se usó automáticamente el archivo local: {excel_label}")
        st.caption(f"Ruta detectada: `{excel_path}`")
    except Exception as e:
        st.error("No encontré un Excel local. Sube un archivo compatible o deja el Excel en la misma carpeta del .py.")
        st.exception(e)
        st.stop()

try:
    detected_instances = detect_instance_numbers(excel_source)
except Exception as e:
    st.error("No pude leer las hojas del archivo Excel seleccionado.")
    st.exception(e)
    st.stop()

if not detected_instances:
    st.error("No se detectaron hojas de instancia en el archivo. Revisa nombres como 'Instancia 1', 'Instancia_1' o '1'.")
    st.stop()

# Selector de modo
modo = st.radio(
    "Modo de cálculo",
    options=["Una instancia", "Todas las instancias"],
    horizontal=True,
    help=(
        "**Una instancia**: elige cuál calcular (más rápido, más detalle). "
        "**Todas las instancias**: calcula todas las detectadas y muestra una tabla "
        "maestra comparativa más una pestaña por instancia."
    ),
)

if modo == "Una instancia":
    col_a, col_b, col_c = st.columns([1, 2, 1])
    with col_a:
        instance_number = st.selectbox(
            "Instancia a resolver",
            options=detected_instances,
            index=0,
            format_func=lambda x: f"Instancia {x}",
        )
    with col_b:
        selected_rules = st.multiselect(
            "Reglas a calcular", options=RULES, default=RULES,
        )
    with col_c:
        show_decisions = st.checkbox(
            "Mostrar decisiones", value=True,
            help="Tabla de candidatos por decisión.",
        )
    instances_to_run = [int(instance_number)]
else:
    col_a, col_b = st.columns([2, 1])
    with col_a:
        selected_rules = st.multiselect(
            "Reglas a calcular (aplica a todas las instancias)",
            options=RULES, default=RULES,
        )
    with col_b:
        show_decisions = st.checkbox(
            "Mostrar decisiones", value=False,
            help="Desactivado por defecto en modo múltiple para velocidad.",
        )
    st.caption(
        f"Se calcularán **{len(detected_instances)} instancias**: "
        f"{', '.join(str(x) for x in detected_instances)}."
    )
    instances_to_run = list(detected_instances)

run_button = st.button("Calcular programación", type="primary")

if not run_button:
    st.info("Selecciona la configuración y presiona 'Calcular programación'.")
    st.stop()

try:
    if not selected_rules:
        st.warning("Selecciona al menos una regla para calcular.")
        st.stop()

    # Ejecutar todas las instancias solicitadas
    all_instances_data: Dict[int, Dict] = {}
    instance_errors: List[Dict] = []
    progress = st.progress(0)
    status = st.empty()

    for idx, inst_num in enumerate(instances_to_run):
        status.text(f"Calculando Instancia {inst_num} ({idx + 1}/{len(instances_to_run)})...")
        try:
            data = compute_instance(
                excel_source=excel_source,
                instance_number=int(inst_num),
                selected_rules=selected_rules,
                store_decisions=show_decisions,
            )
            all_instances_data[int(inst_num)] = data
        except Exception as inst_err:
            instance_errors.append({"Instancia": inst_num, "Error": str(inst_err)})
        progress.progress((idx + 1) / len(instances_to_run))

    status.empty()
    progress.empty()

    if not all_instances_data:
        raise ValueError("No se pudo calcular ninguna instancia.")

    # --------------------------------------------------------
    # MODO: UNA INSTANCIA (render directo)
    # --------------------------------------------------------
    if modo == "Una instancia":
        data = list(all_instances_data.values())[0]
        render_instance_block(data, show_decisions, excel_label)

    # --------------------------------------------------------
    # MODO: TODAS LAS INSTANCIAS (tabla maestra + pestañas)
    # --------------------------------------------------------
    else:
        st.success(
            f"Se calcularon {len(all_instances_data)} instancia(s) | Archivo: {excel_label}"
        )

        if instance_errors:
            st.warning("Algunas instancias no pudieron calcularse:")
            st.dataframe(pd.DataFrame(instance_errors), width="stretch")

        # === TABLA MAESTRA ===
        st.header("Tabla maestra: todas las instancias × todas las reglas")
        master = build_master_summary(all_instances_data)
        st.caption(
            "Una fila por combinación instancia × regla. La columna **Índice comparativo** "
            "se normaliza dentro de cada instancia (min-max, SLA invertido); menor es mejor."
        )
        st.dataframe(master, width="stretch")

        st.subheader("Reglas ganadoras por instancia")
        st.caption(
            "Para cada instancia, muestra qué regla ganó en cada indicador. "
            "La última columna es la regla con mejor índice comparativo global."
        )
        winners = build_winners_table(master)
        st.dataframe(winners, width="stretch")

        # Gráfico comparativo del índice por instancia
        try:
            fig_master = px.bar(
                master, x="Regla", y="Índice comparativo",
                color="Regla", facet_col="Instancia", facet_col_wrap=3,
                title="Índice comparativo por regla en cada instancia (menor es mejor)",
            )
            fig_master.update_layout(showlegend=False, height=500)
            st.plotly_chart(fig_master, width="stretch", key="master_index")
        except Exception:
            pass

        # === DESCARGA MAESTRA ===
        st.subheader("Descarga consolidada")
        master_output = {
            "maestra_instancia_regla": master,
            "reglas_ganadoras": winners,
        }
        if instance_errors:
            master_output["errores_instancia"] = pd.DataFrame(instance_errors)
        for inst_num, data in all_instances_data.items():
            master_output[clean_sheet_name(f"i{inst_num}_comparacion")] = data["comparison"]
            master_output[clean_sheet_name(f"i{inst_num}_indice")] = data["index_table"]
            for rule, sched in data["all_schedules"].items():
                alias = RULE_ALIASES.get(rule, rule[:6])
                master_output[clean_sheet_name(f"i{inst_num}_{alias}_ops")] = sched
                master_output[clean_sheet_name(f"i{inst_num}_{alias}_res")] = data["all_results"][rule]

        master_bytes = make_excel_download(master_output)
        st.download_button(
            label="Descargar resultados consolidados (Excel)",
            data=master_bytes,
            file_name="resultados_todas_instancias.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_master",
        )

        # === PESTAÑAS POR INSTANCIA (detalle completo) ===
        st.header("Detalle por instancia")
        inst_tabs = st.tabs([f"Instancia {n}" for n in all_instances_data.keys()])
        for tab, (inst_num, data) in zip(inst_tabs, all_instances_data.items()):
            with tab:
                render_instance_block(data, show_decisions, excel_label)

except Exception as e:
    st.error("Ocurrió un error al calcular la programación.")
    st.exception(e)