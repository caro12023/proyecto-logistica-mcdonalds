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
import plotly.graph_objects as go
from plotly.subplots import make_subplots
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
    "Holgura crítica por margen disponible + SPT",
]

RULE_ALIASES = {
    "FIFO": "FIFO",
    "SPT": "SPT",
    "EDD": "EDD",
    "WSPT": "WSPT",
    "CR": "CR",
    "ATC": "ATC",
    "Holgura crítica por margen disponible + SPT": "HCritSPT",
}

# Paleta de colores institucional para reglas
RULE_COLORS = {
    "FIFO":    "#1f77b4",
    "SPT":     "#ff7f0e",
    "EDD":     "#2ca02c",
    "WSPT":    "#d62728",
    "CR":      "#9467bd",
    "ATC":     "#8c564b",
    "Holgura crítica por margen disponible + SPT": "#e377c2",
}

STATION_COLORS = {
    "Parrilla":        "#e63946",
    "Freidora":        "#f4a261",
    "Bebidas/Postres": "#2a9d8f",
    "Ensamble":        "#457b9d",
    "Staging/Bolseo":  "#6d6875",
}

DEFAULT_HORIZON_SEG = 20000
HORIZON_BUFFER_SEG = 2000

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
    max_cap_until = 0.0
    if capacities is not None and not capacities.empty and "Hasta seg" in capacities.columns:
        max_cap_until = capacities["Hasta seg"].apply(to_num).max()
    max_release = jobs["r j seg"].apply(to_num).max() if "r j seg" in jobs.columns else 0.0
    processing_cols = [
        "p Parrilla seg", "p Freidora seg", "p BebidaPostre seg",
        "p Ensamble seg", "p Staging seg",
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
    name = str(name)
    for bad in ["[", "]", "*", "?", "/", "\\", ":"]:
        name = name.replace(bad, "_")
    return name[:31]


def excel_readable(excel_source):
    if isinstance(excel_source, (bytes, bytearray)):
        return io.BytesIO(excel_source)
    return excel_source


def find_project_excel() -> Path:
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
        "r j seg", "SLA seg", "d j seg", "w j",
        "p Parrilla seg", "p Freidora seg", "p BebidaPostre seg",
        "p Ensamble seg", "p Staging seg",
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
        "Req Parrilla", "Req Freidora", "Req BebidaPostre",
        "Req Ensamble", "Req Staging",
    ]
    for col in requirement_cols:
        if col not in jobs.columns:
            jobs[col] = "No"
    jobs["Tiempo total procesamiento"] = (
        jobs["p Parrilla seg"] + jobs["p Freidora seg"] + jobs["p BebidaPostre seg"]
        + jobs["p Ensamble seg"] + jobs["p Staging seg"]
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
                "OpID": op_id, "ID Pedido": job_id, "Canal": canal, "TipoPedido": tipo,
                "Etapa": "Preparación", "Estacion": "Parrilla",
                "DuracionProceso": to_num(job["p Parrilla seg"]),
                "rj": rj, "dj": dj, "wj": wj, "pj_total": pj_total, "Predecesores": [],
            })

        if normalize_yes(job.get("Req Freidora", "No")) and to_num(job["p Freidora seg"]) > 0:
            op_id = f"{job_id}_Freidora"
            prep_ops.append(op_id)
            rows.append({
                "OpID": op_id, "ID Pedido": job_id, "Canal": canal, "TipoPedido": tipo,
                "Etapa": "Preparación", "Estacion": "Freidora",
                "DuracionProceso": to_num(job["p Freidora seg"]),
                "rj": rj, "dj": dj, "wj": wj, "pj_total": pj_total, "Predecesores": [],
            })

        if normalize_yes(job.get("Req BebidaPostre", "No")) and to_num(job["p BebidaPostre seg"]) > 0:
            op_id = f"{job_id}_Bebidas_Postres"
            prep_ops.append(op_id)
            rows.append({
                "OpID": op_id, "ID Pedido": job_id, "Canal": canal, "TipoPedido": tipo,
                "Etapa": "Preparación", "Estacion": "Bebidas/Postres",
                "DuracionProceso": to_num(job["p BebidaPostre seg"]),
                "rj": rj, "dj": dj, "wj": wj, "pj_total": pj_total, "Predecesores": [],
            })

        ensamble_op = None
        if normalize_yes(job.get("Req Ensamble", "No")) and to_num(job["p Ensamble seg"]) > 0:
            ensamble_op = f"{job_id}_Ensamble"
            rows.append({
                "OpID": ensamble_op, "ID Pedido": job_id, "Canal": canal, "TipoPedido": tipo,
                "Etapa": "Ensamble", "Estacion": "Ensamble",
                "DuracionProceso": to_num(job["p Ensamble seg"]),
                "rj": rj, "dj": dj, "wj": wj, "pj_total": pj_total,
                "Predecesores": prep_ops.copy(),
            })

        if normalize_yes(job.get("Req Staging", "No")) and to_num(job["p Staging seg"]) > 0:
            staging_pred = [ensamble_op] if ensamble_op is not None else prep_ops.copy()
            rows.append({
                "OpID": f"{job_id}_Staging_Bolseo", "ID Pedido": job_id, "Canal": canal,
                "TipoPedido": tipo, "Etapa": "Staging/Bolseo", "Estacion": "Staging/Bolseo",
                "DuracionProceso": to_num(job["p Staging seg"]),
                "rj": rj, "dj": dj, "wj": wj, "pj_total": pj_total,
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
                "Estacion": [station], "Capacidad aplicada": [1],
                "Desde seg": [0], "Hasta seg": [horizon_seg],
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
# SCHEDULER DINÁMICO
# ============================================================

def schedule_operations(
    operations: pd.DataFrame,
    calendars: Dict[str, List[Tuple[float, float]]],
    rule: str,
    store_decisions: bool = True,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
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

    order_counter = 1
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
        durations_now = [max(float(r["pij"]), 1.0) for r in candidate_rows]
        p_bar = max(sum(durations_now) / len(durations_now), 1.0) if durations_now else 1.0
        k = 2.0

        for row in candidate_rows:
            row["TrabajoPosterior"] = max(float(row["P_remanente"]) - float(row["pij"]), 0.0)
            row["d_op"] = float(row["dj"]) - row["TrabajoPosterior"]
            row["Holgura_ATC"] = max(row["d_op"] - float(row["pij"]) - current_time, 0.0)
            exponente = min(700.0, -row["Holgura_ATC"] / (k * p_bar))
            row["ATC"] = (row["wj_safe"] / row["pij"]) * math.exp(exponente)
            row["pbar"] = p_bar

            if rule == "Holgura crítica por margen disponible + SPT":
                if row["Holgura"] >= 0:
                    row["GrupoPrioridad"] = 1
                    row["ClavePrioridad"] = row["Holgura"]
                else:
                    row["GrupoPrioridad"] = 2
                    row["ClavePrioridad"] = row["pij"]
                if row["Holgura"] < 0:
                    row["EstadoHolgura"] = "Sin margen de cumplimiento"
                elif abs(row["Holgura"]) <= 1e-9:
                    row["EstadoHolgura"] = "Al límite del SLA"
                else:
                    row["EstadoHolgura"] = "Con margen de cumplimiento"

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
        if rule == "Holgura crítica por margen disponible + SPT":
            return sorted(
                candidate_rows,
                key=lambda x: (
                    x["GrupoPrioridad"], x["ClavePrioridad"], x["pij"],
                    x.get("Disponible_desde", 0.0), x["dj"], -x.get("wj_safe", 1.0),
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
                    resource=resource, earliest_start=earliest,
                    duration=float(op["DuracionProceso"]), calendars=calendars,
                )
                if feasible_start is not None:
                    possible_decisions.append({
                        "Resource": resource, "Station": station,
                        "DecisionTime": float(feasible_start),
                    })

        if not possible_decisions:
            pending_debug = pd.DataFrame(pending_ops)
            raise ValueError(
                "No se pudo programar alguna operación. "
                f"Operaciones pendientes: {len(pending_debug)}."
            )

        selected_decision = sorted(possible_decisions, key=lambda x: (x["DecisionTime"], x["Resource"]))[0]
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
                    resource=resource, earliest_start=decision_time,
                    duration=float(op["DuracionProceso"]), calendars=calendars,
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
            "Regla": rule, "Orden": order_counter, "OpID": selected["OpID"],
            "ID Pedido": selected["ID Pedido"], "Canal": selected["Canal"],
            "TipoPedido": selected["TipoPedido"], "Etapa": selected["Etapa"],
            "Estacion": station, "Recurso": resource,
            "Inicio": start, "InicioProceso": start, "Fin": end,
            "Pij": float(selected["DuracionProceso"]), "Duracion": float(selected["DuracionProceso"]),
            "rj": float(selected["rj"]), "dj": float(selected["dj"]),
            "wj": float(selected["wj"]), "pj": float(selected["pj_total"]),
            "P_remanente": float(selected.get("P_remanente", np.nan)),
            "Holgura": float(selected.get("Holgura", np.nan)),
            "EstadoHolgura": selected.get("EstadoHolgura", ""),
            "GrupoPrioridad": selected.get("GrupoPrioridad", np.nan),
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
        schedule.groupby("ID Pedido", as_index=False)
        .agg(Cj=("Fin", "max"), InicioPrimeraOperacion=("Inicio", "min"))
    )
    base = jobs[["ID Pedido", "Canal", "Tipo Pedido", "r j seg", "d j seg", "w j", "Tiempo total procesamiento"]].copy()
    base = base.rename(columns={
        "Tipo Pedido": "TipoPedido", "r j seg": "rj", "d j seg": "dj",
        "w j": "wj", "Tiempo total procesamiento": "pj",
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
    rows = [summarize_rule(rule, results) for rule, results in all_results.items()]
    comparison = pd.DataFrame(rows)
    numeric_cols = comparison.select_dtypes(include=[np.number]).columns
    comparison[numeric_cols] = comparison[numeric_cols].round(2)
    return comparison


def build_ranking(comparison: pd.DataFrame) -> pd.DataFrame:
    metrics = [
        ("Menor Cmax", "Cmax", False), ("Menor ∑Cj", "∑Cj", False),
        ("Menor ∑Fj", "∑Fj", False), ("Menor ∑Uj", "∑Uj", False),
        ("Menor Tmax", "Tmax", False), ("Menor Lmax", "Lmax", False),
        ("Menor ∑Tj", "∑Tj", False), ("Menor ∑Lj", "∑Lj", False),
        ("Menor ∑wjCj", "∑wjCj", False), ("Menor ∑wjTj", "∑wjTj", False),
        ("Mayor cumplimiento SLA", "Cumplimiento SLA (%)", True),
    ]
    rows = []
    for label, col, higher_is_better in metrics:
        best_value = comparison[col].max() if higher_is_better else comparison[col].min()
        winners = comparison[np.isclose(comparison[col], best_value)]["Regla"].tolist()
        rows.append({
            "Indicador destacado": label,
            "Regla destacada": " | ".join(winners),
            "Valor obtenido": round(best_value, 2),
            "Lectura": "Empate" if len(winners) > 1 else "Mejor valor único",
        })
    return pd.DataFrame(rows)


# ============================================================
# CUELLO DE BOTELLA
# ============================================================

def preliminary_bottleneck(jobs: pd.DataFrame, capacities: pd.DataFrame) -> pd.DataFrame:
    load = pd.DataFrame({
        "Estacion": STATIONS,
        "Carga total seg": [
            jobs["p Parrilla seg"].sum(), jobs["p Freidora seg"].sum(),
            jobs["p BebidaPostre seg"].sum(), jobs["p Ensamble seg"].sum(),
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
    if schedule is None or schedule.empty:
        return pd.DataFrame()
    cap = capacities.copy()
    cap["DuracionIntervalo"] = cap["Hasta seg"] - cap["Desde seg"]
    cap["SegundosMaquina"] = cap["Capacidad aplicada"] * cap["DuracionIntervalo"]
    disponibilidad = (
        cap.groupby("Estacion", as_index=False)["SegundosMaquina"]
        .sum().rename(columns={"SegundosMaquina": "Segundos maquina disponibles"})
    )
    sched = schedule.copy()
    sched["CargaProceso"] = sched["Fin"] - sched["Inicio"]
    carga_real = (
        sched.groupby("Estacion", as_index=False)
        .agg(**{"Carga total": ("CargaProceso", "sum"), "Recursos usados": ("Recurso", "nunique")})
    )
    table = (
        pd.DataFrame({"Estacion": STATIONS})
        .merge(carga_real, on="Estacion", how="left")
        .merge(disponibilidad, on="Estacion", how="left")
        .fillna({"Carga total": 0, "Recursos usados": 0, "Segundos maquina disponibles": 0})
    )
    table["Carga por recurso"] = np.where(
        table["Recursos usados"] > 0,
        table["Carga total"] / table["Recursos usados"], 0,
    )
    table["Saturación"] = np.where(
        table["Segundos maquina disponibles"] > 0,
        table["Carga total"] / table["Segundos maquina disponibles"], 0,
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
    return table.sort_values(["Saturación", "Carga por recurso"], ascending=[False, False]).reset_index(drop=True)


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
                    "Tipo": "Solapamiento", "Recurso": resource,
                    "OpID 1": group.loc[i, "OpID"], "OpID 2": group.loc[i + 1, "OpID"],
                    "Fin 1": current_end, "Inicio 2": next_start,
                })
    return pd.DataFrame(violations)


def validate_release_times(schedule: pd.DataFrame) -> pd.DataFrame:
    invalid = schedule[schedule["Inicio"] < schedule["rj"] - 1e-9].copy()
    if invalid.empty:
        return pd.DataFrame()
    invalid["Tipo"] = "Inicio antes de rj"
    return invalid[["Tipo", "OpID", "ID Pedido", "Inicio", "rj"]]


def validate_resource_calendar(schedule: pd.DataFrame, calendars: Dict[str, List[Tuple[float, float]]]) -> pd.DataFrame:
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
                "Tipo": "Violación calendario recurso", "OpID": row["OpID"],
                "ID Pedido": row["ID Pedido"], "Recurso": resource,
                "Inicio": start, "Fin": end,
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
                    "Tipo": "Violación precedencia", "OpID": op_id,
                    "Predecesor": pred, "Inicio operación": start[op_id],
                    "Fin predecesor": finish[pred],
                })
    return pd.DataFrame(violations)


def validate_schedule(schedule: pd.DataFrame, operations: pd.DataFrame, calendars: Dict[str, List[Tuple[float, float]]]) -> pd.DataFrame:
    validations = [
        validate_no_overlap(schedule), validate_release_times(schedule),
        validate_resource_calendar(schedule, calendars), validate_precedence(schedule, operations),
    ]
    validations = [v for v in validations if not v.empty]
    if not validations:
        return pd.DataFrame({
            "Resultado": ["OK"],
            "Detalle": ["No se detectaron violaciones de solapamiento, rj, precedencia o calendario."],
        })
    return pd.concat(validations, ignore_index=True)


# ============================================================
# GRÁFICAS — GANTT Y TARDANZA (originales)
# ============================================================

def resource_sort_key(resource: str) -> Tuple[int, int, str]:
    station = station_from_resource(resource)
    station_idx = STATIONS.index(station) if station in STATIONS else len(STATIONS)
    try:
        resource_num = int(str(resource).rsplit("_", 1)[1])
    except Exception:
        resource_num = 1
    return station_idx, resource_num, str(resource)


def plot_gantt(schedule: pd.DataFrame, rule: str):
    data = schedule.copy()
    data["Inicio"] = pd.to_numeric(data["Inicio"], errors="coerce").fillna(0.0)
    data["Fin"] = pd.to_numeric(data["Fin"], errors="coerce").fillna(0.0)
    data["Duracion"] = pd.to_numeric(data["Duracion"], errors="coerce").fillna(data["Fin"] - data["Inicio"])
    data["Inicio_txt"] = data["Inicio"].round(0).astype(int).astype(str) + " s"
    data["Fin_txt"] = data["Fin"].round(0).astype(int).astype(str) + " s"
    data["Duracion_txt"] = data["Duracion"].round(0).astype(int).astype(str) + " s"
    data["EtiquetaTrabajo"] = data["ID Pedido"].astype(str)
    resource_order = sorted(data["Recurso"].unique().tolist(), key=resource_sort_key)
    max_finish = float(data["Fin"].max()) if not data.empty else 0.0
    tick_step = 300 if max_finish <= 4200 else 600
    tick_end = int(np.ceil(max_finish / tick_step) * tick_step) if max_finish > 0 else tick_step
    tick_vals = list(range(0, tick_end + tick_step, tick_step))
    tick_text = [f"{v} s" for v in tick_vals]
    fig = px.bar(
        data, x="Duracion", y="Recurso", base="Inicio", orientation="h",
        color="ID Pedido", text="EtiquetaTrabajo",
        category_orders={"Recurso": resource_order},
        custom_data=["OpID", "ID Pedido", "Estacion", "Inicio_txt", "Fin_txt", "Duracion_txt", "rj", "dj", "wj"],
        title=f"Diagrama de Gantt — {rule}",
    )
    fig.update_traces(
        textposition="inside", insidetextanchor="middle", textfont_size=10,
        marker_line_width=0.5, cliponaxis=False,
        hovertemplate=(
            "<b>%{customdata[0]}</b><br>Pedido: %{customdata[1]}<br>"
            "Estación: %{customdata[2]}<br>Inicio: %{customdata[3]}<br>"
            "Fin: %{customdata[4]}<br>Duración: %{customdata[5]}<br>"
            "rj: %{customdata[6]:.0f} s<br>dj: %{customdata[7]:.0f} s<br>"
            "wj: %{customdata[8]:.2f}<extra></extra>"
        ),
    )
    fig.update_yaxes(categoryorder="array", categoryarray=resource_order, autorange="reversed",
                     title="Recursos ordenados por estación")
    fig.update_xaxes(title="Tiempo operativo transcurrido (segundos)", type="linear",
                     tickmode="array", tickvals=tick_vals, ticktext=tick_text,
                     range=[0, max(tick_end, 1)], showgrid=True, zeroline=True)
    fig.update_layout(
        height=max(600, 85 + 62 * len(resource_order)), showlegend=False,
        bargap=0.20, margin=dict(l=165, r=45, t=80, b=80), xaxis_tickangle=0,
    )
    return fig


def plot_tardiness(results: pd.DataFrame, rule: str):
    tardy = results[results["Tj"] > 0].copy()
    if tardy.empty:
        return None
    tardy = tardy.sort_values(["Tj", "w_j T_j"], ascending=[False, False])
    fig = px.bar(
        tardy, x="Tj", y="ID Pedido", orientation="h", color="w_j T_j",
        hover_data=["TipoPedido", "Cj", "dj", "Lj", "Tj", "wj", "w_j T_j"],
        title=f"Pedidos tardíos y tardanza — {rule}",
    )
    fig.update_layout(height=520, xaxis_title="Tardanza Tj (segundos)", yaxis_title="Pedido",
                      margin=dict(l=110, r=30, t=70, b=50))
    return fig


def plot_comparison(comparison: pd.DataFrame):
    data = comparison[["Regla", "∑Uj", "∑Tj"]].melt(
        id_vars="Regla", var_name="Indicador", value_name="Valor"
    )
    fig = px.bar(data, x="Regla", y="Valor", color="Indicador", barmode="group",
                 text="Valor", title="Pedidos tardíos y tardanza total por regla")
    fig.update_layout(height=520, xaxis_title="Regla", yaxis_title="Valor",
                      margin=dict(l=70, r=30, t=70, b=90))
    return fig


def plot_cmax(comparison: pd.DataFrame):
    data = comparison.sort_values("Cmax").copy()
    fig = px.bar(data, x="Cmax", y="Regla", orientation="h", text="Cmax",
                 title="Comparación de Cmax por regla")
    fig.update_layout(height=480, xaxis_title="Cmax", yaxis_title="Regla",
                      margin=dict(l=160, r=30, t=70, b=50))
    return fig


# ============================================================
# GRÁFICAS NUEVAS — ANÁLISIS DE CUELLO DE BOTELLA Y COMPARATIVAS
# ============================================================

def plot_saturation_heatmap(real_bottlenecks: Dict[str, pd.DataFrame]) -> go.Figure:
    """
    Heatmap de saturación: eje X = reglas, eje Y = estaciones.
    Color = saturación real (carga / segundos-máquina disponibles).
    Permite identificar de un vistazo qué estación es cuello de botella
    bajo cada regla y si alguna regla alivia la presión en una estación clave.
    """
    rules = list(real_bottlenecks.keys())
    z = []
    text = []
    for station in STATIONS:
        row_z = []
        row_t = []
        for rule in rules:
            df = real_bottlenecks[rule]
            val = df[df["Estacion"] == station]["Saturación"].values
            sat = float(val[0]) if len(val) > 0 else 0.0
            row_z.append(sat)
            row_t.append(f"{sat:.1%}")
        z.append(row_z)
        text.append(row_t)

    short_rules = [RULE_ALIASES.get(r, r) for r in rules]

    fig = go.Figure(data=go.Heatmap(
        z=z,
        x=short_rules,
        y=STATIONS,
        text=text,
        texttemplate="%{text}",
        textfont={"size": 13, "color": "white"},
        colorscale=[
            [0.0, "#1a4f6e"],
            [0.5, "#e07b39"],
            [0.85, "#c0392b"],
            [1.0, "#7b0000"],
        ],
        colorbar=dict(title="Saturación", tickformat=".0%"),
        hovertemplate="Estación: %{y}<br>Regla: %{x}<br>Saturación: %{text}<extra></extra>",
        zmin=0, zmax=1,
    ))

    fig.update_layout(
        title=dict(text="Saturación real por estación y regla de despacho", font=dict(size=16)),
        xaxis=dict(title="Regla de despacho", tickfont=dict(size=12)),
        yaxis=dict(title="Estación", tickfont=dict(size=12), autorange="reversed"),
        height=400,
        margin=dict(l=140, r=60, t=70, b=60),
        font=dict(family="Georgia, serif"),
    )
    return fig


def plot_saturation_bars_by_rule(real_bottlenecks: Dict[str, pd.DataFrame]) -> go.Figure:
    """
    Barras agrupadas: para cada regla, muestra la saturación de cada estación.
    Permite comparar el perfil de carga de una misma regla entre estaciones
    y ver cómo cambia ese perfil al cambiar de regla.
    """
    rows = []
    for rule, df in real_bottlenecks.items():
        alias = RULE_ALIASES.get(rule, rule)
        for _, r in df.iterrows():
            rows.append({"Regla": alias, "Estacion": r["Estacion"], "Saturación": r["Saturación"]})
    data = pd.DataFrame(rows)

    fig = px.bar(
        data, x="Regla", y="Saturación", color="Estacion",
        barmode="group",
        color_discrete_map={s: STATION_COLORS[s] for s in STATIONS if s in STATION_COLORS},
        text=data["Saturación"].map(lambda v: f"{v:.1%}"),
        title="Saturación por estación y regla (barras agrupadas)",
        labels={"Saturación": "Saturación real", "Regla": "Regla de despacho", "Estacion": "Estación"},
    )
    fig.add_hline(y=0.85, line_dash="dot", line_color="#c0392b", line_width=1.5,
                  annotation_text="Umbral alta presión (85%)", annotation_position="top right",
                  annotation_font_color="#c0392b")
    fig.update_traces(textposition="outside", textfont_size=10, cliponaxis=False)
    fig.update_yaxes(tickformat=".0%", range=[0, 1.15])
    fig.update_layout(
        height=500, margin=dict(l=70, r=30, t=70, b=90),
        legend=dict(title="Estación", orientation="h", yanchor="bottom", y=-0.35, xanchor="center", x=0.5),
        font=dict(family="Georgia, serif"),
    )
    return fig


def plot_load_vs_capacity(jobs: pd.DataFrame, capacities: pd.DataFrame) -> go.Figure:
    """
    Barras apiladas: carga teórica total vs. capacidad disponible total
    por estación. Útil para validar si el dimensionamiento de recursos es
    suficiente antes de ver el resultado de la programación.
    """
    load_vals = [
        jobs["p Parrilla seg"].sum(), jobs["p Freidora seg"].sum(),
        jobs["p BebidaPostre seg"].sum(), jobs["p Ensamble seg"].sum(),
        jobs["p Staging seg"].sum(),
    ]
    cap = capacities.copy()
    cap["DuracionIntervalo"] = cap["Hasta seg"] - cap["Desde seg"]
    cap["SegundosMaquina"] = cap["Capacidad aplicada"] * cap["DuracionIntervalo"]
    cap_vals = []
    for s in STATIONS:
        v = cap[cap["Estacion"] == s]["SegundosMaquina"].sum()
        cap_vals.append(float(v) if v > 0 else 1.0)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Capacidad disponible (seg·máquina)",
        x=STATIONS, y=cap_vals,
        marker_color=["rgba(42,157,143,0.35)"] * 5,
        marker_line_color=["#2a9d8f"] * 5, marker_line_width=2,
    ))
    fig.add_trace(go.Bar(
        name="Carga teórica total (seg)",
        x=STATIONS, y=load_vals,
        marker_color=[STATION_COLORS.get(s, "#888") for s in STATIONS],
        text=[f"{v:.0f} s" for v in load_vals],
        textposition="outside",
    ))
    fig.update_layout(
        barmode="overlay",
        title=dict(text="Carga teórica vs. capacidad disponible por estación", font=dict(size=15)),
        xaxis_title="Estación", yaxis_title="Segundos",
        height=450, margin=dict(l=70, r=30, t=70, b=60),
        legend=dict(orientation="h", yanchor="bottom", y=-0.28, xanchor="center", x=0.5),
        font=dict(family="Georgia, serif"),
    )
    return fig


def plot_radar_rules(comparison: pd.DataFrame) -> go.Figure:
    """
    Radar (spider chart) multidimensional por regla.
    Normaliza cada indicador en [0,1] donde 1 = mejor desempeño relativo.
    Permite ver de forma inmediata qué regla tiene un perfil más equilibrado.
    """
    metrics_radar = ["Cmax", "∑Uj", "∑Tj", "∑wjTj", "Cumplimiento SLA (%)"]
    labels = ["Cmax", "∑Uj", "∑Tj", "∑wjTj", "SLA %"]

    df = comparison[["Regla"] + metrics_radar].copy()

    # Normalizar: para minimizar → invertir; para maximizar → directo.
    # Resultado: 1 = mejor, 0 = peor.
    normalized = {}
    for col in metrics_radar:
        vmin = df[col].min()
        vmax = df[col].max()
        span = max(vmax - vmin, 1e-9)
        if col == "Cumplimiento SLA (%)":
            normalized[col] = (df[col] - vmin) / span
        else:
            normalized[col] = 1 - (df[col] - vmin) / span

    fig = go.Figure()
    for _, row in df.iterrows():
        rule = row["Regla"]
        alias = RULE_ALIASES.get(rule, rule)
        vals = [float(normalized[m][df[df["Regla"] == rule].index[0]]) for m in metrics_radar]
        vals_closed = vals + [vals[0]]
        labels_closed = labels + [labels[0]]
        color = RULE_COLORS.get(rule, "#888888")
        fig.add_trace(go.Scatterpolar(
            r=vals_closed, theta=labels_closed, name=alias,
            line=dict(color=color, width=2),
            fill="toself", fillcolor=color.replace(")", ",0.08)").replace("rgb", "rgba") if "rgb" in color else color + "14",
        ))

    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 1], tickformat=".1f",
                                   gridcolor="#e0e0e0", linecolor="#aaa"),
                   angularaxis=dict(gridcolor="#e0e0e0")),
        title=dict(text="Perfil de desempeño multidimensional por regla<br><sup>1 = mejor desempeño relativo en cada indicador</sup>",
                   font=dict(size=15)),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.22, xanchor="center", x=0.5),
        height=520, margin=dict(l=60, r=60, t=100, b=80),
        font=dict(family="Georgia, serif"),
    )
    return fig


def plot_sla_compliance_bars(comparison: pd.DataFrame) -> go.Figure:
    """
    Barras horizontales de cumplimiento SLA (%) por regla, con semáforo de color.
    Lectura inmediata: cuántos pedidos termina a tiempo cada regla.
    """
    data = comparison[["Regla", "∑Uj", "Cumplimiento SLA (%)"]].copy()
    data["Alias"] = data["Regla"].map(lambda r: RULE_ALIASES.get(r, r))
    data = data.sort_values("Cumplimiento SLA (%)", ascending=True)

    colors = []
    for v in data["Cumplimiento SLA (%)"]:
        if v >= 90:
            colors.append("#2a9d8f")
        elif v >= 70:
            colors.append("#f4a261")
        else:
            colors.append("#e63946")

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=data["Cumplimiento SLA (%)"], y=data["Alias"],
        orientation="h",
        marker_color=colors,
        text=[f"{v:.1f}%  ({int(u)} tardíos)" for v, u in zip(data["Cumplimiento SLA (%)"], data["∑Uj"])],
        textposition="inside",
        insidetextanchor="middle",
        textfont=dict(color="white", size=12),
        hovertemplate="Regla: %{y}<br>Cumplimiento SLA: %{x:.1f}%<extra></extra>",
    ))
    fig.add_vline(x=90, line_dash="dot", line_color="#2a9d8f", line_width=1.5,
                  annotation_text="Meta 90 %", annotation_position="top",
                  annotation_font_color="#2a9d8f")
    fig.update_xaxes(range=[0, 105], ticksuffix="%", title="Cumplimiento del SLA (%)")
    fig.update_yaxes(title="Regla de despacho")
    fig.update_layout(
        title=dict(text="Cumplimiento del SLA por regla de despacho<br><sup>Verde ≥ 90 % · Naranja 70–90 % · Rojo < 70 %</sup>",
                   font=dict(size=15)),
        height=420, margin=dict(l=100, r=40, t=90, b=60),
        font=dict(family="Georgia, serif"),
    )
    return fig


def plot_tardiness_distribution(all_results: Dict[str, pd.DataFrame]) -> go.Figure:
    """
    Box plot de la distribución de tardanza (Tj) por regla.
    Permite ver mediana, dispersión y outliers: una regla puede tener
    ∑Tj bajo pero alta varianza (pedidos muy tardíos puntualmente).
    """
    fig = go.Figure()
    for rule, results in all_results.items():
        alias = RULE_ALIASES.get(rule, rule)
        color = RULE_COLORS.get(rule, "#888888")
        tj_vals = results["Tj"].values
        fig.add_trace(go.Box(
            y=tj_vals, name=alias,
            marker_color=color, line_color=color,
            boxmean="sd",
            hovertemplate=f"Regla: {alias}<br>Tardanza: %{{y:.0f}} s<extra></extra>",
        ))
    fig.update_layout(
        title=dict(text="Distribución de tardanza individual (T_j) por regla<br><sup>Punto central = media · Caja = IQR · Bigotes = 1.5×IQR</sup>",
                   font=dict(size=15)),
        yaxis_title="Tardanza T_j (segundos)",
        xaxis_title="Regla de despacho",
        height=470, margin=dict(l=70, r=30, t=90, b=60),
        showlegend=False,
        font=dict(family="Georgia, serif"),
        plot_bgcolor="#fafafa",
    )
    fig.update_yaxes(gridcolor="#e8e8e8", zeroline=True, zerolinecolor="#aaa")
    return fig


def plot_carga_temporal(all_schedules: Dict[str, pd.DataFrame], bin_size: int = 300) -> go.Figure:
    """
    Perfil de carga temporal (working jobs en el tiempo) por regla.
    Para cada intervalo de tiempo, cuenta cuántas operaciones están activas.
    Permite detectar si la programación concentra el trabajo o lo distribuye.
    """
    max_time = max(sched["Fin"].max() for sched in all_schedules.values() if not sched.empty)
    bins = np.arange(0, max_time + bin_size, bin_size)

    fig = go.Figure()
    for rule, sched in all_schedules.items():
        alias = RULE_ALIASES.get(rule, rule)
        color = RULE_COLORS.get(rule, "#888")
        counts = []
        for t in bins[:-1]:
            active = ((sched["Inicio"] <= t + bin_size) & (sched["Fin"] > t)).sum()
            counts.append(active)
        fig.add_trace(go.Scatter(
            x=bins[:-1], y=counts, mode="lines", name=alias,
            line=dict(color=color, width=2),
            hovertemplate=f"Regla: {alias}<br>t=%{{x:.0f}} s<br>Ops activas: %{{y}}<extra></extra>",
        ))

    fig.update_layout(
        title=dict(text=f"Perfil de carga temporal — operaciones activas por intervalo de {bin_size} s<br><sup>Muestra cuándo se concentra el trabajo en la planta</sup>",
                   font=dict(size=15)),
        xaxis_title="Tiempo (segundos)", yaxis_title="Operaciones activas simultáneas",
        height=430, margin=dict(l=70, r=30, t=90, b=60),
        legend=dict(orientation="h", yanchor="bottom", y=-0.28, xanchor="center", x=0.5),
        font=dict(family="Georgia, serif"),
        plot_bgcolor="#fafafa",
    )
    fig.update_xaxes(showgrid=True, gridcolor="#e8e8e8")
    fig.update_yaxes(showgrid=True, gridcolor="#e8e8e8")
    return fig


def plot_saturation_single_rule(real_cb: pd.DataFrame, rule: str) -> go.Figure:
    """
    Gráfica de barras de saturación por estación para UNA regla.
    Muestra además la carga por recurso para entender utilización individual.
    """
    if real_cb.empty:
        return go.Figure()

    alias = RULE_ALIASES.get(rule, rule)
    df = real_cb.copy()

    colors = []
    for s in df["Saturación"]:
        if s >= 0.85:
            colors.append("#e63946")
        elif s >= 0.6:
            colors.append("#f4a261")
        else:
            colors.append("#2a9d8f")

    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=("Saturación por estación", "Carga real por recurso (seg)"),
        horizontal_spacing=0.12,
    )

    fig.add_trace(go.Bar(
        x=df["Estacion"], y=df["Saturación"],
        marker_color=colors,
        text=[f"{v:.1%}" for v in df["Saturación"]],
        textposition="outside",
        name="Saturación",
        hovertemplate="Estación: %{x}<br>Saturación: %{text}<extra></extra>",
    ), row=1, col=1)

    fig.add_trace(go.Bar(
        x=df["Estacion"], y=df["Carga por recurso"],
        marker_color=[STATION_COLORS.get(s, "#888") for s in df["Estacion"]],
        text=[f"{v:.0f} s" for v in df["Carga por recurso"]],
        textposition="outside",
        name="Carga/recurso",
        hovertemplate="Estación: %{x}<br>Carga por recurso: %{text}<extra></extra>",
    ), row=1, col=2)

    fig.add_hline(y=0.85, line_dash="dot", line_color="#c0392b", row=1, col=1,
                  annotation_text="Alta presión", annotation_position="top right")
    fig.add_hline(y=1.0, line_dash="dash", line_color="#7b0000", row=1, col=1)

    fig.update_yaxes(tickformat=".0%", row=1, col=1, title="Saturación", range=[0, 1.2])
    fig.update_yaxes(title="Carga por recurso (s)", row=1, col=2)
    fig.update_layout(
        title=dict(text=f"Cuello de botella real — {alias}", font=dict(size=15)),
        showlegend=False, height=420,
        margin=dict(l=70, r=40, t=100, b=60),
        font=dict(family="Georgia, serif"),
        plot_bgcolor="#fafafa",
    )
    return fig


def plot_weighted_tardiness(comparison: pd.DataFrame) -> go.Figure:
    """
    Comparativa de ∑wjTj (tardanza ponderada) y ∑Tj (tardanza total).
    Doble barra para ver si las reglas priorizan bien los pedidos de mayor peso.
    """
    data = comparison[["Regla", "∑Tj", "∑wjTj"]].copy()
    data["Alias"] = data["Regla"].map(lambda r: RULE_ALIASES.get(r, r))

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="∑T_j (tardanza total)", x=data["Alias"], y=data["∑Tj"],
        marker_color="#457b9d", text=data["∑Tj"].round(0).astype(int),
        textposition="outside",
    ))
    fig.add_trace(go.Bar(
        name="∑w_j·T_j (tardanza ponderada)", x=data["Alias"], y=data["∑wjTj"],
        marker_color="#e63946", text=data["∑wjTj"].round(0).astype(int),
        textposition="outside",
    ))
    fig.update_layout(
        barmode="group",
        title=dict(text="Tardanza total vs. tardanza ponderada por prioridad<br><sup>Una brecha grande indica que los pedidos tardíos son de alta prioridad</sup>",
                   font=dict(size=15)),
        xaxis_title="Regla", yaxis_title="Segundos",
        height=460, margin=dict(l=70, r=30, t=90, b=60),
        legend=dict(orientation="h", yanchor="bottom", y=-0.28, xanchor="center", x=0.5),
        font=dict(family="Georgia, serif"),
        plot_bgcolor="#fafafa",
    )
    fig.update_yaxes(gridcolor="#e8e8e8")
    return fig


# ============================================================
# CÓDIGO R EXPORTABLE
# ============================================================

def dataframe_to_r_tribble(df: pd.DataFrame, object_name: str = "comparacion") -> str:
    cols = list(df.columns)
    r = ["library(tibble)", "", f"{object_name} <- tribble(", "  " + ", ".join([f"~`{c}`" for c in cols]) + ","]
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
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return output.getvalue()


# ============================================================
# HELPERS DE EJECUCIÓN POR INSTANCIA
# ============================================================

def compute_instance(excel_source, instance_number: int, selected_rules: List[str], store_decisions: bool) -> Dict:
    jobs_raw, events, capacities_raw, sheet_name = read_instance_excel(excel_source, int(instance_number))
    jobs = prepare_jobs(jobs_raw)
    horizon_seg = calculate_horizon_seg(jobs, capacities_raw)
    capacities = prepare_capacities(capacities_raw, horizon_seg=horizon_seg)
    operations = build_operations(jobs)
    calendars = build_resource_calendar(capacities, horizon_seg=horizon_seg)
    bottleneck_table = preliminary_bottleneck(jobs, capacities)
    calendar_df = make_calendar_table(calendars)

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
            errors.append({"Instancia": instance_number, "Regla": rule, "Error": str(rule_error)})

    if not all_results:
        raise ValueError(f"No se pudo calcular ninguna regla para la Instancia {instance_number}.")

    comparison = build_comparison(all_results)
    ranking = build_ranking(comparison)

    return {
        "instance_number": instance_number, "sheet_name": sheet_name,
        "jobs": jobs, "events": events, "capacities": capacities,
        "operations": operations, "calendars": calendars, "horizon_seg": horizon_seg,
        "bottleneck_table": bottleneck_table, "calendar_df": calendar_df,
        "all_schedules": all_schedules, "all_results": all_results,
        "all_decisions": all_decisions, "real_bottlenecks": real_bottlenecks,
        "comparison": comparison, "ranking": ranking, "errors": errors,
    }


def render_instance_block(data: Dict, show_decisions: bool, excel_label: str):
    instance_number = data["instance_number"]
    sheet_name = data["sheet_name"]
    jobs = data["jobs"]
    events = data["events"]
    capacities = data["capacities"]
    operations = data["operations"]
    horizon_seg = data["horizon_seg"]
    bottleneck_table = data["bottleneck_table"]
    calendar_df = data["calendar_df"]
    all_schedules = data["all_schedules"]
    all_results = data["all_results"]
    all_decisions = data["all_decisions"]
    real_bottlenecks = data["real_bottlenecks"]
    comparison = data["comparison"]
    ranking = data["ranking"]
    errors = data["errors"]

    st.success(f"Instancia {instance_number} calculada desde la hoja: {sheet_name} | Archivo: {excel_label}")

    # ── 1. Datos ──────────────────────────────────────────────────────────────
    st.header("1. Datos de la instancia")
    st.caption(
        f"Horizonte de calendario usado: {horizon_seg:.0f} segundos. "
        "La regla de holgura crítica separa pedidos con margen de cumplimiento frente al SLA "
        "y pedidos sin margen, aplicando SPT en este último grupo."
    )

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: st.metric("Pedidos", len(jobs))
    with c2: st.metric("Operaciones", len(operations))
    with c3: st.metric("Carga total", f"{jobs['Tiempo total procesamiento'].sum():.0f} seg")
    with c4:
        ventana_llegada = jobs["r j seg"].max() - jobs["r j seg"].min()
        st.metric("Ventana llegada", f"{ventana_llegada:.0f} seg")
    with c5:
        dominant_channel = jobs["Canal"].value_counts().idxmax()
        st.metric("Canal dominante", dominant_channel)

    with st.expander("Ver pedidos originales preparados", expanded=False):
        st.dataframe(jobs, use_container_width=True)

    col_l, col_r = st.columns(2)
    with col_l:
        st.subheader("Capacidades")
        st.dataframe(capacities, use_container_width=True)
    with col_r:
        st.subheader("Eventos")
        if events.empty:
            st.info("No hay eventos registrados para esta instancia.")
        else:
            st.dataframe(events, use_container_width=True)

    with st.expander("Ver operaciones generadas", expanded=False):
        st.dataframe(operations, use_container_width=True)

    col_l, col_r = st.columns(2)
    with col_l:
        st.subheader("Cuello de botella preliminar")
        st.caption("Diagnóstico teórico previo a la programación. Utiliza capacidad promedio.")
        st.dataframe(bottleneck_table, use_container_width=True)
    with col_r:
        st.subheader("Carga teórica vs. capacidad disponible")
        st.plotly_chart(
            plot_load_vs_capacity(jobs, capacities),
            use_container_width=True,
            key=f"load_cap_{instance_number}",
        )

    # ── 2. Comparación global de reglas ───────────────────────────────────────
    st.header("2. Comparación global de reglas de despacho")
    st.dataframe(comparison, use_container_width=True)

    # Fila 1: SLA + Radar
    col_l, col_r = st.columns(2)
    with col_l:
        st.plotly_chart(
            plot_sla_compliance_bars(comparison),
            use_container_width=True, key=f"sla_bars_{instance_number}",
        )
    with col_r:
        st.plotly_chart(
            plot_radar_rules(comparison),
            use_container_width=True, key=f"radar_{instance_number}",
        )

    # Fila 2: Tardanza total vs. ponderada + distribución
    col_l, col_r = st.columns(2)
    with col_l:
        st.plotly_chart(
            plot_weighted_tardiness(comparison),
            use_container_width=True, key=f"wtardy_{instance_number}",
        )
    with col_r:
        st.plotly_chart(
            plot_tardiness_distribution(all_results),
            use_container_width=True, key=f"tdist_{instance_number}",
        )

    # Fila 3: Cmax + perfil temporal
    col_l, col_r = st.columns(2)
    with col_l:
        st.plotly_chart(
            plot_cmax(comparison),
            use_container_width=True, key=f"cmax_{instance_number}",
        )
    with col_r:
        st.plotly_chart(
            plot_carga_temporal(all_schedules),
            use_container_width=True, key=f"carga_temp_{instance_number}",
        )

    # ── 2b. Análisis comparativo de cuello de botella ─────────────────────────
    st.subheader("Análisis comparativo de cuello de botella entre reglas")
    st.caption(
        "El heatmap muestra la saturación real de cada estación bajo cada regla. "
        "Rojo = alta saturación (cuello de botella). El gráfico de barras desglosa el mismo dato "
        "para facilitar la comparación entre estaciones dentro de una misma regla."
    )

    if real_bottlenecks:
        st.plotly_chart(
            plot_saturation_heatmap(real_bottlenecks),
            use_container_width=True, key=f"heatmap_{instance_number}",
        )
        st.plotly_chart(
            plot_saturation_bars_by_rule(real_bottlenecks),
            use_container_width=True, key=f"sat_bars_{instance_number}",
        )

    # Ranking
    st.subheader("Ranking por indicador")
    st.dataframe(ranking, use_container_width=True)

    if errors:
        st.warning("Algunas reglas no pudieron calcularse.")
        st.dataframe(pd.DataFrame(errors), use_container_width=True)

    # ── 3. Resultados por regla ───────────────────────────────────────────────
    st.header("3. Resultados por regla")
    rule_tabs = st.tabs(list(all_schedules.keys()))

    for rule_tab, rule in zip(rule_tabs, all_schedules.keys()):
        with rule_tab:
            st.subheader(rule)

            st.markdown("#### Gantt ordenado por estación")
            st.plotly_chart(
                plot_gantt(all_schedules[rule], rule),
                use_container_width=True, key=f"gantt_{instance_number}_{rule}",
            )

            # Cuello de botella individual de esta regla
            st.markdown("#### Cuello de botella real — saturación post-simulación")
            real_cb = real_bottlenecks.get(rule, pd.DataFrame())
            if real_cb.empty:
                st.info("No hay operaciones programadas para calcular saturación.")
            else:
                st.plotly_chart(
                    plot_saturation_single_rule(real_cb, rule),
                    use_container_width=True, key=f"sat_single_{instance_number}_{rule}",
                )
                st.dataframe(real_cb, use_container_width=True)

            # Tardanza individual
            st.markdown("#### Pedidos tardíos")
            tardy_fig = plot_tardiness(all_results[rule], rule)
            if tardy_fig is None:
                st.info(f"Bajo la regla {rule} no se registran pedidos tardíos.")
            else:
                st.plotly_chart(tardy_fig, use_container_width=True,
                                key=f"tardy_{instance_number}_{rule}")

            st.markdown("#### Resultados por pedido")
            st.dataframe(all_results[rule], use_container_width=True)
            st.markdown("#### Operaciones programadas")
            st.dataframe(all_schedules[rule], use_container_width=True)

            with st.expander("Ver tabla de decisiones"):
                if all_decisions[rule].empty:
                    st.info("Tabla de decisiones desactivada o vacía.")
                else:
                    st.dataframe(all_decisions[rule], use_container_width=True)

    # ── 4. Descarga ───────────────────────────────────────────────────────────
    st.header("4. Descargar resultados de esta instancia")
    output_dfs = {
        "comparacion": comparison, "ranking": ranking, "pedidos": jobs,
        "operaciones_generadas": operations, "capacidades": capacities,
        "eventos": events if not events.empty else pd.DataFrame({"Detalle": ["Sin eventos"]}),
        "cuello_preliminar": bottleneck_table, "calendario_recursos": calendar_df,
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
    rows = []
    for inst_num, data in all_instances_data.items():
        comp = data["comparison"]
        for _, r in comp.iterrows():
            rows.append({
                "Instancia": inst_num, "Regla": r["Regla"],
                "Cmax": r["Cmax"], "∑Tj": r["∑Tj"], "∑Uj": r["∑Uj"],
                "∑wjTj": r["∑wjTj"], "Cumplimiento SLA (%)": r["Cumplimiento SLA (%)"],
            })
    return pd.DataFrame(rows).sort_values(["Instancia", "∑Uj", "∑Tj", "Cmax", "Regla"]).reset_index(drop=True)


def build_winners_table(master: pd.DataFrame) -> pd.DataFrame:
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
        rows.append({
            "Instancia": inst_num,
            "Mejor Cmax": f"{cmax_r} ({cmax_v})",
            "Mejor ∑Tj": f"{tj_r} ({tj_v})",
            "Mejor ∑Uj": f"{uj_r} ({uj_v})",
            "Mejor ∑wjTj": f"{wjtj_r} ({wjtj_v})",
            "Mejor SLA %": f"{sla_r} ({sla_v})",
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
    La herramienta admite dos modalidades de trabajo: cargar un archivo Excel con la estructura
    definida para el proyecto, o utilizar automáticamente el archivo local disponible en el
    directorio de ejecución. El sistema detecta las hojas de instancia presentes en el archivo
    y permite seleccionar la instancia a resolver.

    ---

    **Criterio de desempeño principal: ∑U_j — número de pedidos que incumplen el SLA**

    El indicador central de análisis es la sumatoria de pedidos tardíos (∑U_j), definida como
    la cantidad de pedidos cuyo tiempo de finalización supera la fecha de entrega comprometida (d_j).
    Este criterio se justifica desde la perspectiva del servicio al cliente en restaurantes de
    comida rápida: el incumplimiento del tiempo de atención prometido —el SLA— afecta directamente
    la percepción del cliente, genera insatisfacción y deteriora indicadores operativos clave
    como el tiempo promedio de atención, la tasa de cumplimiento de órdenes y la retención de
    clientes en los distintos canales de servicio (AutoMac, Mostrador, Delivery y Pickup).
    Por esta razón, la comparación entre reglas de despacho se interpreta indicador por indicador,
    con énfasis en ∑U_j, ∑T_j (tardanza total acumulada), ∑w_jT_j (tardanza ponderada por
    prioridad de pedido) y el porcentaje de cumplimiento del SLA.
    """
)

uploaded_excel = st.file_uploader(
    "Cargar archivo Excel de configuración compatible",
    type=["xlsx"],
    help=(
        "El archivo debe contener hojas de instancia con nombres como 'Instancia 1', 'Instancia_1' o '1'. "
        "Opcionalmente puede incluir hojas 'Capacidades' y 'Eventos' con columna Instancia."
    ),
)

if uploaded_excel is not None:
    excel_source = uploaded_excel.getvalue()
    excel_label = uploaded_excel.name
    st.success(f"Archivo cargado: {excel_label}")
else:
    try:
        excel_path = find_project_excel()
        excel_source = excel_path
        excel_label = excel_path.name
        st.info(f"No se cargó un archivo; se utilizará automáticamente el archivo local: {excel_label}")
        st.caption(f"Ruta detectada: `{excel_path}`")
    except Exception as e:
        st.error("No se encontró un archivo Excel local. Cargue un archivo compatible o coloque el Excel en la carpeta del archivo .py.")
        st.exception(e)
        st.stop()

try:
    detected_instances = detect_instance_numbers(excel_source)
except Exception as e:
    st.error("No fue posible leer las hojas del archivo Excel seleccionado.")
    st.exception(e)
    st.stop()

if not detected_instances:
    st.error("No se detectaron hojas de instancia. Verifique que los nombres sean del tipo 'Instancia 1', 'Instancia_1' o '1'.")
    st.stop()

modo = st.radio(
    "Modo de cálculo",
    options=["Una instancia", "Todas las instancias"],
    horizontal=True,
    help=(
        "**Una instancia**: seleccione la instancia a calcular (más rápido, mayor detalle). "
        "**Todas las instancias**: calcula todas las detectadas y genera una tabla maestra comparativa."
    ),
)

if modo == "Una instancia":
    col_a, col_b, col_c = st.columns([1, 2, 1])
    with col_a:
        instance_number = st.selectbox(
            "Instancia a resolver", options=detected_instances, index=0,
            format_func=lambda x: f"Instancia {x}",
        )
    with col_b:
        selected_rules = st.multiselect("Reglas a calcular", options=RULES, default=RULES)
    with col_c:
        show_decisions = st.checkbox("Mostrar decisiones", value=True,
                                     help="Tabla de candidatos por decisión.")
    instances_to_run = [int(instance_number)]
else:
    col_a, col_b = st.columns([2, 1])
    with col_a:
        selected_rules = st.multiselect(
            "Reglas a calcular (aplica a todas las instancias)", options=RULES, default=RULES,
        )
    with col_b:
        show_decisions = st.checkbox("Mostrar decisiones", value=False,
                                     help="Desactivado por defecto en modo múltiple para mayor velocidad.")
    st.caption(
        f"Se calcularán **{len(detected_instances)} instancias**: "
        f"{', '.join(str(x) for x in detected_instances)}."
    )
    instances_to_run = list(detected_instances)

run_button = st.button("Calcular programación", type="primary")

if not run_button:
    st.info("Seleccione la configuración y presione 'Calcular programación'.")
    st.stop()

try:
    if not selected_rules:
        st.warning("Seleccione al menos una regla para calcular.")
        st.stop()

    all_instances_data: Dict[int, Dict] = {}
    instance_errors: List[Dict] = []
    progress = st.progress(0)
    status = st.empty()

    for idx, inst_num in enumerate(instances_to_run):
        status.text(f"Calculando Instancia {inst_num} ({idx + 1}/{len(instances_to_run)})...")
        try:
            data = compute_instance(
                excel_source=excel_source, instance_number=int(inst_num),
                selected_rules=selected_rules, store_decisions=show_decisions,
            )
            all_instances_data[int(inst_num)] = data
        except Exception as inst_err:
            instance_errors.append({"Instancia": inst_num, "Error": str(inst_err)})
        progress.progress((idx + 1) / len(instances_to_run))

    status.empty()
    progress.empty()

    if not all_instances_data:
        raise ValueError("No se pudo calcular ninguna instancia.")

    if modo == "Una instancia":
        data = list(all_instances_data.values())[0]
        render_instance_block(data, show_decisions, excel_label)

    else:
        st.success(f"Se calcularon {len(all_instances_data)} instancia(s) | Archivo: {excel_label}")

        if instance_errors:
            st.warning("Algunas instancias no pudieron calcularse:")
            st.dataframe(pd.DataFrame(instance_errors), use_container_width=True)

        st.header("Tabla maestra: todas las instancias × todas las reglas")
        master = build_master_summary(all_instances_data)
        st.caption(
            "Una fila por combinación instancia × regla. ∑Uj mide pedidos que incumplen el SLA, "
            "∑Tj mide la tardanza total, ∑wjTj pondera la tardanza por prioridad y "
            "Cumplimiento SLA (%) resume el servicio al cliente."
        )
        st.dataframe(master, use_container_width=True)

        st.subheader("Reglas ganadoras por instancia")
        winners = build_winners_table(master)
        st.dataframe(winners, use_container_width=True)

        st.subheader("Descarga consolidada")
        master_output = {"maestra_instancia_regla": master, "reglas_ganadoras": winners}
        if instance_errors:
            master_output["errores_instancia"] = pd.DataFrame(instance_errors)
        for inst_num, data in all_instances_data.items():
            master_output[clean_sheet_name(f"i{inst_num}_comparacion")] = data["comparison"]
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

        st.header("Detalle por instancia")
        inst_tabs = st.tabs([f"Instancia {n}" for n in all_instances_data.keys()])
        for tab, (inst_num, data) in zip(inst_tabs, all_instances_data.items()):
            with tab:
                render_instance_block(data, show_decisions, excel_label)

except Exception as e:
    st.error("Ocurrió un error al calcular la programación.")
    st.exception(e)
