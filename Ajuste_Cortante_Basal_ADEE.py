#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ajuste de Cortante Basal NSR-10 A.5.4.5 — ETABS (ADEE)
======================================================
Aplicación de escritorio UNIFICADA en Python (Tkinter) que automatiza
el ajuste de cortante basal según NSR-10, replicando la lógica validada
de hoja de cálculo y reemplazando la consola por un flujo de 3 vistas.

Vista 1 — Conexión y Configuración Inicial
  • Botón Conectar al modelo activo de ETABS (comtypes)
  • Selectores dinámicos extraídos de ETABS: Casos dinámicos X/Y, Caso Modal, Función Espectro
  • Inputs: Ta X/Y [s], Coeficiente Cu
  • Toggle Regular/Irregular → define 0.80 / 0.90 (NSR-10 A.5.4.5)

Vista 2 — Extracción y Verificación de Datos (Tablas)
  • Tabla 1: Mass Summary by Story (Piso, Diafragma, Masa X/Y) + fila TOTAL en kg
  • Tabla 2: Modal Participating Mass Ratios — resalta fila Tx (máx UX) y Ty (máx UY)

Vista 3 — Cálculos Paso a Paso y Resultados Finales
  • Paso 1: Tmax = Ta*Cu, Tmodal, T_ajustado = min(Tmax,Tmodal)
  • Paso 2: Sa = numpy.interp(T_ajustado, curva ETABS)
  • Paso 3: Vs = M_total * Sa * g / 1000  [kN]  (g=9.80665)
  • Paso 4: Vt = |FX|max / |FY|max de BaseReact para casos dinámicos
  • Paso 5: ¿Vt < porcentaje*Vs? → SÍ/NO
  • Resultado Final: Factor = (porcentaje*Vs)/Vt ; si no requiere → 1.0 / No Aplica

Backend
  • Unidades: SapModel.SetPresentUnits(6) → kN, m, C (masa en Ton = kN·s²/m → *1000 = kg)
  • pandas para DataFrames ETABS, numpy.interp para espectro
  • Manejo robusto de API ETABS (variantes de retorno) + Mock para pruebas sin ETABS

Uso
  Windows con ETABS abierto:  python Ajuste_Cortante_Basal_ADEE.py
  Pruebas sin ETABS (Linux/Mac): python Ajuste_Cortante_Basal_ADEE.py --mock
  Dependencias: pip install pandas numpy comtypes openpyxl  (comtypes solo Windows)

Autor: Muse Spark — versión unificada, corrige problemas previos (Qt xcb, imports, índices API)
Ruta destino: C:/Luis estudio/Civil codigos importantes/Uso del agente para ajuste ADEE
"""
from __future__ import annotations
import sys
import os
import traceback
import datetime
from pathlib import Path
from dataclasses import dataclass
from typing import List, Tuple, Any

# ---------------------------------------------------------------------------
# 0) Verificación de dependencias críticas con mensajes claros (corrige crash al correr)
# ---------------------------------------------------------------------------
MISSING_DEPS = []

try:
    import pandas as pd
except ImportError:
    pd = None  # type: ignore
    MISSING_DEPS.append("pandas")

try:
    import numpy as np
except ImportError:
    np = None  # type: ignore
    MISSING_DEPS.append("numpy")

# Tkinter es estándar; si falta, se caerá a modo consola
try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
    TK_AVAILABLE = True
except ImportError:
    TK_AVAILABLE = False
    tk = None  # type: ignore

GRAVEDAD = 9.80665

# ---------------------------------------------------------------------------
# 1) Helpers ETABS (robustos ante variantes de retorno COM)
#     Corrección crítica: ETABS 18/19/20/21 devuelven GetTableForDisplayArray con
#     orden distinto (retCode al inicio/final, Fields en índice variable). Se usa
#     parser que busca por contenido, no por índice fijo → evita 'int object not iterable'
# ---------------------------------------------------------------------------
import difflib
import re

def _safe_get_name_list(obj) -> List[str]:
    """Extrae lista de nombres probando firmas comtypes GetNameList."""
    # Firma 1: GetNameList(0, [])
    try:
        ret = obj.GetNameList(0, [])
        if isinstance(ret, (list, tuple)):
            for elem in ret:
                if isinstance(elem, (list, tuple)) and len(elem) > 0 and isinstance(elem[0], str):
                    return list(elem)
            if len(ret) >= 2 and isinstance(ret[1], (list, tuple)):
                return list(ret[1])
    except Exception:
        pass
    # Firma 2: GetNameList() sin args
    try:
        ret = obj.GetNameList()
        if isinstance(ret, (list, tuple)):
            for elem in ret:
                if isinstance(elem, (list, tuple)) and elem and isinstance(elem[0], str):
                    return list(elem)
    except Exception:
        pass
    return []


def _parse_display_table_ret(ret, debug_key: str = "") -> Tuple[List[str], List[str], int]:
    """
    Parser robusto para SapModel.DatabaseTables.GetTableForDisplayArray
    Corrige el error 'int object is not iterable' del reporte.
    Busca por contenido, no por índice fijo:
      - Fields: lista corta (<30) de strings que contiene palabras clave (Story, Mass, Name, Period...)
      - Data: lista plana larga (>Fields) con valores
      - retCode: entero 0=OK, otro=error (puede estar en ret[0], ret[4] o ret[-1])
    Retorna (fields, data, ret_code). Si no se encuentra, lanza con dump útil.
    """
    if ret is None:
        raise RuntimeError(f"[{debug_key}] GetTableForDisplayArray retornó None")
    if not isinstance(ret, (list, tuple)):
        raise RuntimeError(f"[{debug_key}] Retorno no es lista/tupla: {type(ret)} = {ret}")

    # Dump para debug si falla
    def _dump():
        parts = []
        for i, e in enumerate(ret):
            t = type(e).__name__
            if isinstance(e, (list, tuple)):
                preview = str(e[:6]) + ("..." if len(e) > 6 else "")
                parts.append(f"  ret[{i}] ({t}, len={len(e)}): {preview}")
            else:
                parts.append(f"  ret[{i}] ({t}): {repr(e)[:200]}")
        return "\n".join(parts)

    # Detectar retCode: buscar int 0/1 pequeño al inicio o final
    ret_code = 0
    # ETABS suele poner retCode = 0/1 al final, pero en comtypes a veces en índice 4
    candidates = []
    for idx in [ -1, 0, 4, len(ret)-1, len(ret)-2 ]:
        try:
            v = ret[idx] if idx >= 0 else ret[len(ret)+idx]
            if isinstance(v, int) and v in (0, 1):
                # Si es 0 casi siempre es OK; priorizar
                candidates.append(v)
        except Exception:
            continue
    # Heurística: si algún elemento es int y los demás son listas, ese es retCode
    # Buscamos el primer int que no sea parte de Data length
    int_indices = [i for i, e in enumerate(ret) if isinstance(e, int)]
    if int_indices:
        # Si hay un 0 al final, es retCode
        if isinstance(ret[-1], int) and ret[-1] in (0, 1):
            ret_code = int(ret[-1])
        elif isinstance(ret[0], int) and ret[0] in (0, 1) and len(ret) > 3:
            # Algunas versiones ponen retCode primero
            ret_code = int(ret[0])
        elif len(ret) > 4 and isinstance(ret[4], int):
            ret_code = int(ret[4])

    # Buscar fields: lista de strings corta con palabras clave
    fields = None
    # Palabras clave típicas por tabla
    field_keywords = ["story", "diaphragm", "mass", "name", "period", "value", "ux", "uy", "px", "elevation", "height"]
    best_score = -1
    for elem in ret:
        if isinstance(elem, (list, tuple)) and 1 <= len(elem) <= 40 and all(isinstance(x, str) for x in elem):
            # Score por contener keywords y longitud moderada
            lname = " ".join(elem).lower()
            score = sum(1 for kw in field_keywords if kw in lname)
            # Bonus si todos son encabezados razonables (<20 chars cada uno)
            if all(len(s) < 30 for s in elem):
                score += 0.5
            if score > best_score:
                best_score = score
                fields = list(elem)

    # Buscar data: lista plana larga (>50% fields len) con valores mixtos strings/números
    data = None
    longest = -1
    for elem in ret:
        if isinstance(elem, (list, tuple)) and len(elem) > 0:
            # Data típica es flat list de strings, longitud >> fields
            # Puede contener números como strings, no todos deben ser strings header-like
            if fields is not None and elem is fields:
                continue
            if len(elem) > longest and len(elem) >= (len(fields) if fields else 3) * 2:
                # Verificar que no sea lista de nombres cortos (como list de funciones)
                # Data suele ser muy larga (> n_records * n_cols)
                data = list(elem)
                longest = len(elem)

    # Fallbacks tradicionales por índice si no se encontró
    if fields is None:
        for try_idx in [1, 2, 0]:
            try:
                cand = ret[try_idx]
                if isinstance(cand, (list, tuple)) and len(cand) > 0 and isinstance(cand[0], str):
                    fields = list(cand)
                    break
            except Exception:
                continue
    if data is None:
        for try_idx in [3, 2, 4]:
            try:
                cand = ret[try_idx]
                if isinstance(cand, (list, tuple)) and len(cand) > 5:
                    data = list(cand)
                    break
            except Exception:
                continue

    if fields is None or data is None:
        raise RuntimeError(
            f"[{debug_key}] No se pudo parsear GetTableForDisplayArray.\n"
            f"ret_code detectado={ret_code}\nDump ret:\n{_dump()}\n"
            f"Sugerencia: verifique que el modelo esté corrido (Run Analysis) y que la tabla exista. "
            f"En ETABS pruebe Display > Show Tables > Mass Summary by Story para confirmar."
        )
    return fields, data, ret_code


def _resolve_spectrum_name(SapModel, typed: str, available: List[str]) -> Tuple[str, str]:
    """
    Manejo difuso para nombre de función espectro (segundo requerimiento).
    - typed: lo escrito por usuario (puede tener espacios, mayúsculas, sin guiones)
    - available: lista exacta de ETABS (GetNameList)
    Retorna (nombre_resuelto, mensaje).
    Lógica:
      1) Match exacto case-sensitive → OK
      2) Match exacto case-insensitive + strip → OK (auto-corrige)
      3) Normalizado (sin espacios/guiones/acentos, lower) → OK
      4) difflib close_matches (cutoff 0.6) → sugiere más cercano, auto-selecciona si cutoff 0.8
      5) Si no, retorna typed original + warning con lista disponibles
    """
    if not typed:
        return typed, "Nombre vacío"
    if not available:
        return typed, "No hay funciones en el modelo para comparar"
    # 1) exacto
    if typed in available:
        return typed, "Coincidencia exacta"
    # 2) case-insensitive + strip
    t_norm = typed.strip()
    for av in available:
        if av.strip().lower() == t_norm.lower():
            return av, f"Corregido mayúsculas/espacios: '{typed}' → '{av}'"
    # 3) normalizado: quitar espacios, guiones, guiones bajos, lower
    def _norm(s): return re.sub(r"[\s_\-]+", "", s.lower())
    t_n = _norm(t_norm)
    for av in available:
        if _norm(av) == t_n:
            return av, f"Corregido normalizado: '{typed}' → '{av}'"
    # 4) difflib
    # Probar con lower para mejor match
    lower_map = {av.lower(): av for av in available}
    matches = difflib.get_close_matches(t_norm.lower(), [av.lower() for av in available], n=1, cutoff=0.6)
    if matches:
        suggested_lower = matches[0]
        suggested = lower_map[suggested_lower]
        # Calcular ratio
        ratio = difflib.SequenceMatcher(None, t_norm.lower(), suggested_lower).ratio()
        if ratio >= 0.8:
            return suggested, f"Corregido difuso (ratio {ratio:.2f}): '{typed}' → '{suggested}' (auto)"
        else:
            # Si ratio medio, igual sugerir pero avisar
            return suggested, f"Sugerido difuso (ratio {ratio:.2f}): '{typed}' → '{suggested}' — verifique si es la función deseada. Disponibles: {available}"
    # 5) no match
    return typed, f"No se encontró coincidencia para '{typed}'. Disponibles: {available}. Escriba exacto o seleccione de la lista."


# --- Helpers para altura y Ta automático (tercer requerimiento) ---
NSR10_CT_ALPHA = {
    # Tabla A.4.2-1 NSR-10 + correcciones comunes
    "Pórticos concreto (Ct=0.047, α=0.90) — NSR-10 A.4.2-1 #1": (0.047, 0.90),
    "Pórticos acero (Ct=0.072, α=0.80) — NSR-10 A.4.2-1 #2": (0.072, 0.80),
    "Pórticos acero arriostrados (Ct=0.073, α=0.75) — NSR-10": (0.073, 0.75),
    "Muros concreto/mampostería (Ct=0.049, α=0.75) — NSR-10": (0.049, 0.75),
    "Dual / Muros + pórticos (Ct=0.049, α=0.75)": (0.049, 0.75),
    "Muros mampostería reforzada (Ct=0.047, α=0.90)": (0.047, 0.90),
}

def get_building_height_and_stories(SapModel) -> Tuple[float, int, float]:
    """
    Extrae altura total h [m] y número de pisos N del modelo ETABS.
    Intenta: Story.GetStories() → DatabaseTables 'Story Definitions'
    Retorna (h_total_m, N_pisos, h_prom_m). Lanza si no puede.
    """
    # Intento 1: API Story.GetStories / GetNameList
    try:
        ret = SapModel.Story.GetStories()
        if isinstance(ret, (list, tuple)) and len(ret) >= 2:
            names = None
            # Recolectar todas las listas de floats candidatas
            float_lists = []
            for elem in ret:
                if isinstance(elem, (list, tuple)) and elem and isinstance(elem[0], str):
                    if any("story" in str(x).lower() or "piso" in str(x).lower() for x in elem):
                        names = list(elem)
                    elif names is None and len(elem) > 1 and isinstance(elem[0], str):
                        names = list(elem)
                if isinstance(elem, (list, tuple)) and elem and isinstance(elem[0], (int, float)):
                    try:
                        vals = [float(x) for x in elem]
                        # Considerar listas con valores plausibles de altura/elev (0-300m)
                        if vals and max(vals) >= 0 and max(vals) <= 500 and len(vals) >= 2:
                            float_lists.append(vals)
                    except Exception:
                        continue
            # Elegir elevs como la lista de floats con mayor max (elevaciones > alturas)
            elevs = None
            if float_lists:
                # Ordenar por max descendente; elevaciones tendrán max mayor (9 vs 3)
                float_lists.sort(key=lambda v: max(v), reverse=True)
                # Tomar la primera que tenga variación (no todas iguales)
                for cand in float_lists:
                    if len(set(cand)) > 1:  # tiene variación, es elevación
                        elevs = cand
                        break
                if elevs is None:
                    elevs = float_lists[0]
            if names and elevs and len(names) == len(elevs):
                h_total = float(max(elevs) - min(elevs))
                if h_total < 0.1:
                    h_total = float(max(elevs))
                # Contar pisos excluyendo Base
                base_count = sum(1 for n in names if "base" in str(n).lower())
                N_pisos = len(names) - base_count if base_count else len(names)
                # Si base estaba incluida pero h_total incluía base (0), N_pisos correcto; si no, usar len
                if N_pisos <= 0:
                    N_pisos = len(names)
                h_prom = h_total / N_pisos if N_pisos else 0
                return h_total, N_pisos, h_prom
    except Exception:
        pass

    # Intento 2: DatabaseTables Story Definitions / Story Data
    for table_key in ["Story Definitions", "Story Data", "Stories"]:
        try:
            ret = SapModel.DatabaseTables.GetTableForDisplayArray(table_key, [], "")
            fields, data, rc = _parse_display_table_ret(ret, debug_key=table_key)
            if rc != 0:
                continue
            # Buscar columnas: Name/Story, Elevation, Height
            col_idx = {h.lower(): i for i, h in enumerate(fields)}
            # Encontrar índice de elevación
            elev_idx = None
            for k in col_idx:
                if "elevation" in k:
                    elev_idx = col_idx[k]; break
            if elev_idx is None:
                # Probar Height
                for k in col_idx:
                    if "height" in k:
                        elev_idx = col_idx[k]; break
            if elev_idx is None:
                continue
            name_idx = None
            for k in col_idx:
                if "story" in k or "name" in k:
                    name_idx = col_idx[k]; break
            ncol = len(fields)
            rows = [data[i:i+ncol] for i in range(0, len(data), ncol)]
            names = [r[name_idx] for r in rows] if name_idx is not None else [f"S{i}" for i in range(len(rows))]
            elevs = []
            for r in rows:
                try:
                    elevs.append(float(str(r[elev_idx]).replace(",", "")))
                except Exception:
                    elevs.append(0.0)
            if elevs:
                h_total = float(max(elevs) - min(elevs))
                if h_total < 1.0:
                    h_total = float(max(elevs))
                N = len(names)
                # Filtrar "Base"
                # Si primer story es Base con elev 0, N_pisos = N-1
                base_count = sum(1 for n in names if "base" in n.lower())
                N_pisos = N - base_count if base_count else N
                h_prom = h_total / N_pisos if N_pisos else h_total / N if N else 0
                return h_total, N_pisos if N_pisos else N, h_prom
        except Exception:
            continue

    # Intento 3: fallback — preguntar usuario (no hay datos)
    raise RuntimeError(
        "No se pudo extraer altura del modelo automáticamente.\n"
        "Intente: en ETABS verifique Define > Story Data o Display > Show Tables > Story Definitions.\n"
        "Como alternativa, ingrese la altura total h [m] y número de pisos N manualmente en el diálogo de cálculo automático."
    )


def calc_Ta_nsr10(h: float, sistema: str, N: int | None = None) -> Tuple[float, str]:
    """
    Calcula Ta = Ct * h^α según NSR-10 A.4.2-3, o Ta = 0.1*N alternativa.
    Retorna (Ta, detalle_formula).
    """
    if h <= 0:
        raise ValueError("Altura h debe ser >0")
    if sistema not in NSR10_CT_ALPHA:
        # Intentar buscar por substring
        for k in NSR10_CT_ALPHA:
            if sistema.lower() in k.lower():
                sistema = k; break
    Ct, alpha = NSR10_CT_ALPHA.get(sistema, (0.047, 0.90))
    Ta = Ct * (h ** alpha)
    # Alternativa 0.1*N si N <=12 y hp <=3m (se informa como referencia)
    alt_text = f"Ta = Ct·h^α = {Ct}·{h:.2f}^{alpha} = {Ta:.4f} s  ({sistema})"
    if N and N <= 12:
        Ta_alt = 0.1 * N
        # Calcular hp promedio
        hp = h / N if N else 0
        if hp <= 3.1:
            alt_text += f"\nAlternativa NSR-10 A.4.2-5 (N≤12, hp={hp:.2f}m≤3m): Ta = 0.1·N = 0.1·{N} = {Ta_alt:.4f} s"
        else:
            alt_text += f"\nAlternativa 0.1·N = {Ta_alt:.4f}s no aplica (hp={hp:.2f}m >3m)"
    return Ta, alt_text


def calc_Cu_nsr10(Av: float, Fv: float) -> Tuple[float, str]:
    """
    NSR-10 A.4.2-2: Cu = 1.75 - 1.2·Av·Fv, limitado 1.2 ≤ Cu ≤ 1.4 (algunas versiones 1.75)
    Usamos límites 1.2 ≤ Cu ≤ 1.75 según texto oficial, con advertencia si >1.4.
    """
    Cu_raw = 1.75 - 1.2 * Av * Fv
    Cu = max(1.2, min(Cu_raw, 1.75))
    detalle = f"Cu = 1.75 - 1.2·Av·Fv = 1.75 - 1.2·{Av}·{Fv} = {Cu_raw:.4f} → limitado 1.2–1.75 = {Cu:.4f} s"
    if Cu > 1.4:
        detalle += " (NSR-10 algunas interpretaciones limitan a 1.4; verifique con su revisora)"
    return Cu, detalle


def get_all_load_cases(SapModel) -> List[str]:
    try:
        return _safe_get_name_list(SapModel.LoadCases)
    except Exception:
        return []


def get_modal_cases(SapModel) -> List[str]:
    all_cases = get_all_load_cases(SapModel)
    mods: List[str] = []
    for lc in all_cases:
        try:
            ret = SapModel.LoadCases.GetTypeOAPI(lc)
            tipo = None
            if isinstance(ret, (list, tuple)):
                for v in ret:
                    if isinstance(v, int) and v in (1, 2, 3, 4, 5, 6, 7, 8):
                        tipo = v
                        break
            if tipo == 3:
                mods.append(lc)
        except Exception:
            continue
    if not mods:
        for lc in all_cases:
            if "modal" in lc.lower():
                mods.append(lc)
        if not mods:
            return all_cases
    return mods


def get_response_spectrum_cases(SapModel) -> List[str]:
    all_cases = get_all_load_cases(SapModel)
    rs: List[str] = []
    for lc in all_cases:
        try:
            ret = SapModel.LoadCases.GetTypeOAPI(lc)
            tipo = None
            if isinstance(ret, (list, tuple)):
                for v in ret:
                    if isinstance(v, int) and v in (1, 2, 3, 4, 5, 6, 7, 8):
                        tipo = v
                        break
            if tipo == 4:
                rs.append(lc)
        except Exception:
            continue
    if not rs:
        mods = set(get_modal_cases(SapModel))
        rs = [c for c in all_cases if c not in mods]
        if not rs:
            rs = all_cases
    return rs


def _get_available_table_keys(SapModel) -> List[str]:
    """Intenta listar todas las TableKeys disponibles vía DatabaseTables.GetAvailableTables."""
    try:
        # Intento 1: sin args (comtypes devuelve tupla)
        ret = SapModel.DatabaseTables.GetAvailableTables()
        if isinstance(ret, (list, tuple)):
            for elem in ret:
                if isinstance(elem, (list, tuple)) and elem and isinstance(elem[0], str):
                    # Filtrar que parezca TableKey (contiene espacios y mayúsculas)
                    if len(elem) > 5:
                        return list(elem)
            # Si ret[1] es lista
            if len(ret) >= 2 and isinstance(ret[1], (list, tuple)):
                return list(ret[1])
    except Exception:
        pass
    try:
        # Intento 2: con args por referencia
        ret = SapModel.DatabaseTables.GetAvailableTables(0, [], 0, [])
        if isinstance(ret, (list, tuple)):
            for elem in ret:
                if isinstance(elem, (list, tuple)) and elem and isinstance(elem[0], str):
                    if len(elem) > 5:
                        return list(elem)
    except Exception:
        pass
    # Fallback: lista hardcodeada de candidatas conocidas (se prueba luego)
    return []


def get_spectrum_functions(SapModel) -> List[str]:
    """
    Lista TODAS las funciones de espectro RS visibles en ETABS.
    Estrategia en 3 capas para que la lista desplegable aparezca completa
    (corrige bug donde solo aparecía 'Zona 1 Dosq.' y faltaba 'nsr_10_Mzl_E.'):
      1) Func.FuncRS.GetNameList con varias firmas / tipos
      2) DatabaseTables.GetAvailableTables -> filtrar keys con 'Response Spectrum'
      3) Tabla hardcodeada + DatabaseTables fallback
    Combina resultados de todas las fuentes en un set.
    """
    found = set()

    # --- Capa 1: FuncRS.GetNameList con múltiples variantes ---
    # ETABS separa User vs Code vs NSR-10; probar con distintos índices
    func_rs_obj = None
    try:
        func_rs_obj = SapModel.Func.FuncRS
    except Exception:
        pass
    if func_rs_obj is not None:
        for args in [(0, []), (1, []), (2, []), (6, []), (0,), (1,), (), ([],), ([""],)]:
            try:
                if isinstance(args, tuple) and len(args) == 2:
                    ret = func_rs_obj.GetNameList(args[0], args[1])  # type: ignore
                elif isinstance(args, tuple) and len(args) == 1 and isinstance(args[0], int):
                    ret = func_rs_obj.GetNameList(args[0])  # type: ignore
                elif args == ():
                    ret = func_rs_obj.GetNameList()  # type: ignore
                else:
                    continue
                # Extraer lista de nombres del retorno
                lst = []
                if isinstance(ret, (list, tuple)):
                    for elem in ret:
                        if isinstance(elem, (list, tuple)) and elem and isinstance(elem[0], str):
                            lst = list(elem)
                            break
                    if not lst and len(ret) >= 2 and isinstance(ret[1], (list, tuple)):
                        lst = list(ret[1])
                if lst:
                    for n in lst:
                        if n and str(n).strip():
                            found.add(str(n).strip())
            except Exception:
                continue
        # También probar _safe_get_name_list genérico
        try:
            lst2 = _safe_get_name_list(func_rs_obj)
            for n in lst2:
                found.add(str(n).strip())
        except Exception:
            pass

    # --- Capa 1b: Func genérico (algunas versiones listan RS vía Func.GetNameList) ---
    try:
        generic = _safe_get_name_list(SapModel.Func)
        for n in generic:
            if "espectro" in n.lower() or "spectrum" in n.lower() or "nsr" in n.lower() or "zona" in n.lower():
                found.add(str(n).strip())
    except Exception:
        pass

    # --- Capa 2: Descubrir tablas vía GetAvailableTables ---
    try:
        avail_keys = _get_available_table_keys(SapModel)
        # Filtrar candidatas relevantes
        cand_keys = [k for k in avail_keys if any(kw in k.lower() for kw in ["response spectrum", "spectrum", "espectro", "function"])]
        # Si no encontró nada relevante, probar todas (limitado)
        if not cand_keys:
            cand_keys = avail_keys[:40]
        for table_key in cand_keys:
            # Solo intentar si parece función/espectro
            if not any(kw in table_key.lower() for kw in ["response", "spectrum", "function", "espectro", "rs"]):
                continue
            try:
                ret = SapModel.DatabaseTables.GetTableForDisplayArray(table_key, [], "")
                fields, data, rc = _parse_display_table_ret(ret, debug_key=table_key)
                if rc != 0 or not fields or not data:
                    continue
                # Buscar columna Name
                low_fields = [f.strip().lower() for f in fields]
                idx = -1
                for i, h in enumerate(low_fields):
                    if h in ("name", "function", "function name", "rs name", "spectrum name", "func name"):
                        idx = i; break
                if idx == -1:
                    idx = 0
                ncol = len(fields)
                rows = [data[i:i+ncol] for i in range(0, len(data), ncol)]
                for r in rows:
                    if len(r) <= idx: continue
                    name = str(r[idx]).strip()
                    if name and name.lower() != "none":
                        found.add(name)
            except Exception:
                continue
    except Exception:
        pass

    # --- Capa 3: Hardcode + fallback directo (por si GetAvailableTables falla) ---
    hardcoded_keys = [
        "Functions - Response Spectrum - User Defined",
        "Functions: Response Spectrum - User Defined",
        "Functions - Response Spectrum - User",
        "Functions - Response Spectrum - Code",
        "Functions - Response Spectrum - NSR-10",
        "Functions - Response Spectrum - Code - NSR-10",
        "Functions - Response Spectrum",
        "Function Definitions - Response Spectrum",
        "Response Spectrum Functions",
        "Functions - Response Spectrum - Auto",
        "Functions  - Response Spectrum - User Defined",
        "Funciones - Espectro de Respuesta - Definido por el Usuario",
        "Funciones - Espectro de Respuesta",
        "Espectro de Respuesta",
    ]
    for table_key in hardcoded_keys:
        try:
            ret = SapModel.DatabaseTables.GetTableForDisplayArray(table_key, [], "")
            fields, data, rc = _parse_display_table_ret(ret, debug_key=table_key)
            if rc != 0 or not fields or not data:
                continue
            low_fields = [f.strip().lower() for f in fields]
            idx = -1
            for i, h in enumerate(low_fields):
                if h in ("name", "function", "function name", "rs name", "spectrum name", "func name"):
                    idx = i; break
            if idx == -1: idx = 0
            ncol = len(fields)
            rows = [data[i:i+ncol] for i in range(0, len(data), ncol)]
            for r in rows:
                if len(r) <= idx: continue
                name = str(r[idx]).strip()
                if name and name.lower() != "none":
                    found.add(name)
        except Exception:
            continue

    # Si aún vacío, devolver lo que haya (puede ser solo Zona 1 Dosq.)
    # Normalizar y ordenar: primero las que parecen espectro
    result = sorted(found, key=lambda s: (0 if "zona" in s.lower() or "nsr" in s.lower() or "espectro" in s.lower() or "spectrum" in s.lower() else 1, s.lower()))
    return result


def connect_to_etabs():
    try:
        import comtypes.client  # type: ignore
    except ImportError as e:
        raise RuntimeError("comtypes no instalado. Ejecute: pip install comtypes  (solo Windows)") from e
    try:
        etabs_obj = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
        return etabs_obj.SapModel, etabs_obj
    except Exception as e:
        raise RuntimeError("No se encontró ETABS abierto. Abra ETABS con un modelo activo.") from e

# ---------------------------------------------------------------------------
# 2) Extracción (unidades kN-m-C, pandas, numpy)
# ---------------------------------------------------------------------------
@dataclass
class MassResult:
    df: Any
    masa_total_ton: float
    masa_total_kg: float
    col_masa: str

@dataclass
class ModalResult:
    df: Any
    tx_modal: float
    ty_modal: float
    idx_max_ux: int
    idx_max_uy: int
    periods: Any
    ux: Any
    uy: Any

@dataclass
class BaseResult:
    df: Any
    VtX: float
    VtY: float

@dataclass
class CalculoResult:
    Ta_x: float; Ta_y: float; Cu: float; porcentaje: float; tipo: str
    masa_total_kg: float; masa_total_ton: float
    Tmax_x: float; Tmax_y: float; Tx_modal: float; Ty_modal: float
    T_aju_x: float; T_aju_y: float; Sa_x: float; Sa_y: float
    periods: Any; sa: Any; func_name: str
    VsX: float; VsY: float; VtX: float; VtY: float
    req_X: float; req_Y: float
    necesita_x: bool; necesita_y: bool
    factor_x: Any; factor_y: Any
    factor_x_num: float; factor_y_num: float
    caso_dx: str; caso_dy: str


def extract_mass_summary(SapModel) -> MassResult:
    """
    Extrae Mass Summary by Story con parser robusto que evita 'int object not iterable'.
    Muestra dump útil si falla para diagnóstico del usuario.
    """
    if pd is None or np is None:
        raise RuntimeError("Faltan pandas/numpy. Instale: pip install pandas numpy")
    try:
        SapModel.SetPresentUnits(6)
    except Exception:
        pass
    # Intentar primero tabla estándar, luego variantes por idioma/versión
    last_exc = None
    for table_key in ["Mass Summary by Story", "Mass Summary by Story - GUI", "Masa por Piso"]:
        try:
            ret = SapModel.DatabaseTables.GetTableForDisplayArray(table_key, [], "")
            headers, data, ret_code = _parse_display_table_ret(ret, debug_key=table_key)
            if ret_code != 0:
                last_exc = RuntimeError(f"ETABS retornó código {ret_code} para tabla '{table_key}'")
                continue
            ncol = len(headers)
            if ncol == 0:
                last_exc = RuntimeError(f"Tabla '{table_key}' sin columnas")
                continue
            if len(data) == 0:
                last_exc = RuntimeError(f"Tabla '{table_key}' vacía — verifique que el modelo esté corrido (Run) y tenga masas definidas (Define > Mass Source)")
                continue
            # Validar que headers contengan al menos Story/Mass
            low_headers = " ".join(headers).lower()
            if "mass" not in low_headers and "story" not in low_headers:
                last_exc = RuntimeError(f"Tabla '{table_key}' no parece ser Mass Summary. Headers: {headers}")
                continue
            rows = [data[i:i+ncol] for i in range(0, len(data), ncol)]
            # Si data no cuadra exacto, avisar
            if len(rows) * ncol != len(data):
                # Algunos ETABS devuelven datos con fila de totales extra; truncar
                rows = [data[i:i+ncol] for i in range(0, len(data) - len(data) % ncol, ncol)]
            df = pd.DataFrame(rows, columns=headers)
            # Éxito
            break
        except Exception as e:
            last_exc = e
            continue
    else:
        # Si sale del for sin break, propagar último error con ayuda
        raise RuntimeError(
            f"No se pudo extraer Mass Summary by Story.\n"
            f"Último error: {last_exc}\n\n"
            f"Causas comunes:\n"
            f" • El modelo no está corrido: ejecute Analyze > Run Analysis.\n"
            f" • No hay Mass Source definido: Define > Mass Source.\n"
            f" • Nombre de tabla cambia por versión/idioma de ETABS. Revise Display > Show Tables.\n"
            f" • Pruebe activar 'Modo Mock' para ver formato esperado."
        ) from last_exc

    # Columna masa: busca MassX exacto, luego fallback
    col_masa = None
    for col in df.columns:
        if col.replace(" ", "").lower() == "massx":
            col_masa = col
            break
    if col_masa is None:
        for col in df.columns:
            if "mass" in col.lower():
                col_masa = col
                break
    if col_masa is None:
        col_masa = df.columns[2] if len(df.columns) > 2 else df.columns[0]

    df[col_masa] = pd.to_numeric(df[col_masa], errors="coerce").fillna(0)
    masa_ton = float(df[col_masa].sum())
    masa_kg = masa_ton * 1000.0
    return MassResult(df=df, masa_total_ton=masa_ton, masa_total_kg=masa_kg, col_masa=col_masa)


def extract_modal_data(SapModel, caso_modal: str) -> ModalResult:
    if pd is None or np is None:
        raise RuntimeError("Faltan pandas/numpy")
    try:
        SapModel.SetPresentUnits(6)
    except Exception:
        pass
    try:
        SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
        SapModel.Results.Setup.SetCaseSelectedForOutput(caso_modal)
    except Exception as e:
        raise RuntimeError(f"Error seleccionando caso modal '{caso_modal}': {e}") from e

    ret = SapModel.Results.ModalParticipatingMassRatios()
    if not ret or len(ret) < 7:
        raise RuntimeError(f"ModalParticipatingMassRatios retorno inesperado: {ret}")
    try:
        periods = pd.to_numeric(ret[4], errors="coerce")
        ux = pd.to_numeric(ret[5], errors="coerce")
        uy = pd.to_numeric(ret[6], errors="coerce")
    except Exception as e:
        raise RuntimeError(f"Error parseando modal: {e}") from e

    periods = np.array(periods, dtype=float)
    ux = np.array(ux, dtype=float)
    uy = np.array(uy, dtype=float)
    if periods.ndim == 0:
        periods = np.array([periods]); ux = np.array([ux]); uy = np.array([uy])
    if len(ux) == 0:
        raise RuntimeError("ModalParticipatingMassRatios vacío")

    idx_x = int(np.argmax(ux)); idx_y = int(np.argmax(uy))
    tx = float(periods[idx_x]); ty = float(periods[idx_y])

    df = pd.DataFrame({"Mode": np.arange(1, len(periods)+1), "Period": periods, "UX": ux, "UY": uy})
    try:
        if len(ret) > 8:
            df["SumUX"] = pd.to_numeric(ret[8], errors="coerce")
            df["SumUY"] = pd.to_numeric(ret[9], errors="coerce")
    except Exception:
        pass
    return ModalResult(df=df, tx_modal=tx, ty_modal=ty, idx_max_ux=idx_x, idx_max_uy=idx_y, periods=periods, ux=ux, uy=uy)


def _extract_spectrum_via_table(SapModel, func_name: str):
    """
    Fallback cuando SapModel.Func.FuncRS.GetCurve no existe (AttributeError)
    o falla por versión de ETABS. Lee directamente la tabla
    'Functions - Response Spectrum - User Defined' via DatabaseTables.
    Es el método más compatible entre ETABS 18/19/20/21/22.
    Ahora también descubre tablas disponibles dinámicamente, por lo que
    encuentra 'nsr_10_Mzl_E.' aunque no esté en la lista hardcodeada.
    """
    last_exc = None
    # Probar varios nombres de tabla que usa ETABS según versión/idioma
    # Primero descubrir todas las tablas disponibles que contengan espectro/función
    discovered = []
    try:
        avail = _get_available_table_keys(SapModel)
        # Filtrar relevantes
        for k in avail:
            lk = k.lower()
            if any(x in lk for x in ["response spectrum", "spectrum", "espectro", "response spectra"]):
                discovered.append(k)
        # Si no encontró nada relevante, usar las primeras 15 como fallback
        if not discovered and avail:
            discovered = avail[:15]
    except Exception:
        pass

    hardcoded = [
        "Functions - Response Spectrum - User Defined",
        "Functions: Response Spectrum - User Defined",
        "Functions - Response Spectrum - User",
        "Functions - Response Spectrum - Code",
        "Functions - Response Spectrum - NSR-10",
        "Functions - Response Spectrum - Code - NSR-10",
        "Functions - Response Spectrum",
        "Function Definitions - Response Spectrum",
        "Response Spectrum Functions",
        "Functions  - Response Spectrum - User Defined",
        "Funciones - Espectro de Respuesta - Definido por el Usuario",
        "Funciones - Espectro de Respuesta",
        "Espectro de Respuesta",
        "Functions - Response Spectrum - Auto",
    ]
    # Unir descubiertas primero (más probable), luego hardcode sin duplicar
    table_candidates = []
    for k in discovered + hardcoded:
        if k not in table_candidates:
            table_candidates.append(k)
    for table_key in table_candidates:
        try:
            ret = SapModel.DatabaseTables.GetTableForDisplayArray(table_key, [], "")
            fields, data, rc = _parse_display_table_ret(ret, debug_key=table_key)
            if rc != 0:
                last_exc = RuntimeError(f"Tabla '{table_key}' retCode={rc}")
                continue
            if not fields or not data:
                continue
            # Normalizar encabezados
            low_fields = [f.strip().lower() for f in fields]
            # Buscar columna Name (puede ser 'Name', 'Function', 'Function Name')
            idx_name = -1
            for i, h in enumerate(low_fields):
                if h in ("name", "function", "function name", "spectrum name", "func name"):
                    idx_name = i; break
            if idx_name == -1:
                # Si no hay Name, asumir primera columna
                idx_name = 0
            # Buscar Period
            idx_per = -1
            for i, h in enumerate(low_fields):
                if "period" in h or "per" == h or "t [" in h:
                    idx_per = i; break
            if idx_per == -1:
                # Si no hay Period, intentar segunda columna
                idx_per = 1 if len(fields) > 1 else -1
            # Buscar Value / Accel / Sa / Ordinate
            idx_val = -1
            for i, h in enumerate(low_fields):
                if any(k in h for k in ["value", "accel", "ordinate", "sa", "acceleration", "ampl"]):
                    # Evitar col de Name
                    if i != idx_name:
                        idx_val = i; break
            if idx_val == -1:
                idx_val = 2 if len(fields) > 2 else 1

            ncol = len(fields)
            rows = [data[i:i+ncol] for i in range(0, len(data), ncol)]
            # Filtrar solo filas donde Name == func_name (comparación tolerante)
            per_vals = []
            sa_vals = []
            # Normalizar func_name para comparar
            target_norm = func_name.strip().lower()
            for r in rows:
                if len(r) <= max(idx_name, idx_per, idx_val):
                    continue
                row_name = str(r[idx_name]).strip()
                if row_name.lower() == target_norm:
                    try:
                        p = float(str(r[idx_per]).replace(",", "."))
                        v = float(str(r[idx_val]).replace(",", "."))
                        per_vals.append(p)
                        sa_vals.append(v)
                    except Exception:
                        continue
            # Si no encontró con igualdad exacta, intentar difuso dentro de tabla
            if not per_vals:
                # Buscar filas donde row_name contiene target o viceversa
                for r in rows:
                    row_name = str(r[idx_name]).strip()
                    if not row_name:
                        continue
                    # Comparación difusa leve: si target está contenido o ratio alto
                    if target_norm in row_name.lower() or row_name.lower() in target_norm:
                        try:
                            p = float(str(r[idx_per]).replace(",", "."))
                            v = float(str(r[idx_val]).replace(",", "."))
                            per_vals.append(p); sa_vals.append(v)
                        except Exception:
                            continue
            if per_vals and sa_vals:
                per_arr = np.array(per_vals, dtype=float)
                sa_arr = np.array(sa_vals, dtype=float)
                # Ordenar y limpiar duplicados
                idx = np.argsort(per_arr)
                return per_arr[idx], sa_arr[idx]
        except Exception as e:
            last_exc = e
            continue
    raise RuntimeError(f"No se pudo leer espectro '{func_name}' vía tablas. Último error: {last_exc}. "
                       f"Verifique en ETABS: Display > Show Tables > Functions > Response Spectrum.") from last_exc


def extract_spectrum_curve(SapModel, func_name: str):
    """
    Extrae curva Sa vs T con manejo difuso de nombre y fallback robusto.
    1) Resuelve nombre difuso (mayúsculas/espacios/guiones)
    2) Intenta SapModel.Func.FuncRS.GetCurve (varias firmas)
    3) Si falla por AttributeError (ETABS 22 / API distinta) → fallback vía DatabaseTables
    """
    if np is None:
        raise RuntimeError("Falta numpy")
    try:
        SapModel.SetPresentUnits(6)
    except Exception:
        pass

    # --- Resolver nombre difuso antes de llamar API ---
    typed_original = func_name
    try:
        available = get_spectrum_functions(SapModel)
        if available:
            resolved, msg = _resolve_spectrum_name(SapModel, func_name, available)
            if resolved != func_name:
                print(f"[Spectrum] {msg}")
                func_name = resolved
                extract_spectrum_curve._last_resolve_msg = msg  # type: ignore
                extract_spectrum_curve._last_resolved = resolved  # type: ignore
            else:
                extract_spectrum_curve._last_resolve_msg = "OK"  # type: ignore
                extract_spectrum_curve._last_resolved = func_name  # type: ignore
        else:
            extract_spectrum_curve._last_resolve_msg = "OK (sin lista)"  # type: ignore
            extract_spectrum_curve._last_resolved = func_name  # type: ignore
    except Exception:
        extract_spectrum_curve._last_resolve_msg = "No se pudo resolver (sin lista disponible)"  # type: ignore
        extract_spectrum_curve._last_resolved = typed_original  # type: ignore

    # --- Intentar vía API GetCurve con múltiples firmas ---
    last_exc = None
    tried_via_api = False
    for try_name in [func_name, typed_original] if func_name != typed_original else [func_name]:
        tried_via_api = True
        # Probar varios nombres de método que existen según versión ETABS
        for method_name in ["GetCurve", "GetFuncCurve", "GetUserCurve", "GetRS_Curve"]:
            try:
                func_rs = SapModel.Func.FuncRS
                # Verificar que método existe (evita el KeyError de comtypes)
                if not hasattr(func_rs, method_name):
                    # Intentar obtener via __getattr__ capturando AttributeError limpio
                    try:
                        meth = getattr(func_rs, method_name)
                    except AttributeError as ae:
                        raise AttributeError(f"{method_name} no existe") from ae
                else:
                    meth = getattr(func_rs, method_name)

                # Algunas firmas requieren argumentos diferentes
                # Firma 1: GetCurve(Name) -> (Periods, Values, ret)
                # Firma 2: GetCurve(Name, NumberItems, Period[], Value[]) por ref
                # Probamos comtypes genérico: llamada con un solo arg
                ret = meth(try_name)
                if not isinstance(ret, (list, tuple)):
                    raise RuntimeError(f"{method_name} no retornó lista: {ret}")
                periods = sa = ret_code = None
                if len(ret) >= 4:
                    periods, sa, ret_code = ret[1], ret[2], ret[3]
                elif len(ret) == 3:
                    # Puede ser (Periods, Values, ret) o (NumberItems, Periods, Values)
                    # Heurística: si ret[0] es int y ret[1] es lista → es NumberItems
                    if isinstance(ret[0], int) and isinstance(ret[1], (list, tuple)):
                        periods, sa, ret_code = ret[1], ret[2], 0
                    else:
                        periods, sa, ret_code = ret[0], ret[1], ret[2]
                else:
                    raise RuntimeError(f"{method_name} retorno inesperado: {ret}")
                if isinstance(ret_code, int) and ret_code != 0:
                    raise RuntimeError(f"Error {method_name} '{try_name}' código {ret_code}")
                periods = np.array(list(periods), dtype=float); sa = np.array(list(sa), dtype=float)
                if len(periods) == 0:
                    raise RuntimeError(f"Curva '{try_name}' vacía")
                idx = np.argsort(periods)
                extract_spectrum_curve._last_used_name = try_name  # type: ignore
                return periods[idx], sa[idx]
            except AttributeError as ae:
                # Este método no existe, probar siguiente
                last_exc = ae
                continue
            except KeyError as ke:
                # comtypes lanza KeyError interno por name.lower() -> convertir a AttributeError
                last_exc = AttributeError(f"{method_name} no mapeado (KeyError {ke})")
                continue
            except Exception as e:
                # Otro error (ej. código !=0) → probar siguiente nombre/caso
                last_exc = e
                # Si fue error de código, no tiene sentido probar otros method_name para mismo try_name
                # pero igual probamos fallback tabla
                break
        # Si salió del loop de method_name sin éxito por AttributeError, probar fallback tabla antes de siguiente try_name
        if last_exc and isinstance(last_exc, AttributeError):
            continue

    # --- Fallback vía DatabaseTables (más compatible) ---
    # Esto es lo que soluciona el error de la imagen: AttributeError: GetCurve para 'Zona 1 Dosq.'
    try:
        per_arr, sa_arr = _extract_spectrum_via_table(SapModel, func_name if func_name != typed_original else typed_original)
        extract_spectrum_curve._last_used_name = func_name  # type: ignore
        extract_spectrum_curve._last_resolve_msg = (getattr(extract_spectrum_curve, "_last_resolve_msg", "") + " (vía tabla)")  # type: ignore
        return per_arr, sa_arr
    except Exception as e_table:
        last_exc = e_table
        # Intentar también con nombre original si era distinto
        if func_name != typed_original:
            try:
                per_arr, sa_arr = _extract_spectrum_via_table(SapModel, typed_original)
                extract_spectrum_curve._last_used_name = typed_original  # type: ignore
                return per_arr, sa_arr
            except Exception as e2:
                last_exc = e2
                pass

    # Si todo falla, mostrar lista disponibles y sugerencia
    try:
        avail_str = ", ".join(get_spectrum_functions(SapModel)[:12])
        if not avail_str:
            avail_str = "(vacía — verifique Define > Functions > Response Spectrum)"
    except Exception:
        avail_str = "(no se pudo listar)"
    raise RuntimeError(
        f"No se pudo extraer curva espectro para '{typed_original}'.\n"
        f"Último error API: {last_exc}\n"
        f"Último error tabla: {last_exc}\n"
        f"Funciones disponibles en ETABS: {avail_str}\n"
        f"Tip: verifique Define > Functions > Response Spectrum, copie el nombre exacto, "
        f"o seleccione de la lista desplegable. Si el método GetCurve falla por versión de ETABS, "
        f"la app ahora lee automáticamente la tabla 'Functions - Response Spectrum - User Defined'."
    ) from last_exc


def extract_base_reactions(SapModel, caso_x: str, caso_y: str) -> BaseResult:
    if pd is None or np is None:
        raise RuntimeError("Faltan pandas/numpy")
    try:
        SapModel.SetPresentUnits(6)
    except Exception:
        pass
    try:
        SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
        SapModel.Results.Setup.SetCaseSelectedForOutput(caso_x)
        SapModel.Results.Setup.SetCaseSelectedForOutput(caso_y)
    except Exception as e:
        raise RuntimeError(f"Error seleccionando casos dinámicos: {e}") from e

    ret = SapModel.Results.BaseReact()
    if not ret or len(ret) < 6:
        raise RuntimeError(f"BaseReact retorno inesperado: {ret}")
    try:
        load_cases = list(ret[1])
        fx = pd.to_numeric(ret[4], errors="coerce")
        fy = pd.to_numeric(ret[5], errors="coerce")
    except Exception as e:
        raise RuntimeError(f"Error parseando BaseReact: {e}") from e

    df = pd.DataFrame({"Load Case": load_cases, "FX": fx, "FY": fy})
    # Evitar crash con .str cuando hay NaN o ints: convertir a str
    def _vt(df_, caso, col):
        # Filtro exacto y case-insensitive seguro
        mask_exact = df_["Load Case"].astype(str) == str(caso)
        sub = df_[mask_exact]
        if sub.empty:
            mask_lower = df_["Load Case"].astype(str).str.lower() == str(caso).lower()
            sub = df_[mask_lower]
        if sub.empty:
            raise RuntimeError(f"No hay reacción para '{caso}'. Disponibles: {list(df_['Load Case'].astype(str).unique())}")
        vals = pd.to_numeric(sub[col], errors="coerce").abs()
        return float(vals.max())
    return BaseResult(df=df, VtX=_vt(df, caso_x, "FX"), VtY=_vt(df, caso_y, "FY"))


def calcular_ajuste(Ta_x, Ta_y, Cu, porcentaje, tipo, masa_kg, Tx, Ty, Sa_x, Sa_y, VtX, VtY, periods, sa, func_name, caso_dx, caso_dy) -> CalculoResult:
    Tmax_x = Ta_x * Cu; Tmax_y = Ta_y * Cu
    Taju_x = min(Tmax_x, Tx); Taju_y = min(Tmax_y, Ty)
    masa_ton = masa_kg / 1000.0
    VsX = masa_kg * Sa_x * GRAVEDAD / 1000.0
    VsY = masa_kg * Sa_y * GRAVEDAD / 1000.0
    req_X = porcentaje * VsX; req_Y = porcentaje * VsY
    necesita_x = VtX < req_X; necesita_y = VtY < req_Y
    factor_x = (req_X / VtX if VtX != 0 else float("inf")) if necesita_x else "No Aplica"
    factor_y = (req_Y / VtY if VtY != 0 else float("inf")) if necesita_y else "No Aplica"
    fx_num = float(factor_x) if necesita_x else 1.0
    fy_num = float(factor_y) if necesita_y else 1.0
    return CalculoResult(
        Ta_x=Ta_x, Ta_y=Ta_y, Cu=Cu, porcentaje=porcentaje, tipo=tipo,
        masa_total_kg=masa_kg, masa_total_ton=masa_ton,
        Tmax_x=Tmax_x, Tmax_y=Tmax_y, Tx_modal=Tx, Ty_modal=Ty,
        T_aju_x=Taju_x, T_aju_y=Taju_y, Sa_x=Sa_x, Sa_y=Sa_y,
        periods=periods, sa=sa, func_name=func_name,
        VsX=VsX, VsY=VsY, VtX=VtX, VtY=VtY, req_X=req_X, req_Y=req_Y,
        necesita_x=necesita_x, necesita_y=necesita_y,
        factor_x=factor_x, factor_y=factor_y, factor_x_num=fx_num, factor_y_num=fy_num,
        caso_dx=caso_dx, caso_dy=caso_dy
    )

# ---------------------------------------------------------------------------
# 3) Mock para pruebas sin ETABS
# ---------------------------------------------------------------------------
class MockSapModel:
    def __init__(self):
        if pd is None or np is None:
            raise RuntimeError("Mock requiere pandas/numpy")
        self._mass_df = pd.DataFrame({
            "Story": ["Story1", "Story2", "Story3"],
            "Diaphragm": ["D1", "D1", "D1"],
            "MassSource": ["M1", "M1", "M1"],
            "MassX": [150.5, 120.3, 80.2],
            "MassY": [150.5, 120.3, 80.2],
        })
        self._periods = np.array([0.85, 0.65, 0.45, 0.30, 0.20])
        self._ux = np.array([0.05, 0.65, 0.10, 0.15, 0.05])
        self._uy = np.array([0.60, 0.05, 0.20, 0.05, 0.10])
        self._spec_p = np.array([0.0, 0.1, 0.2, 0.3, 0.5, 0.75, 1.0, 1.5, 2.0])
        self._spec_sa = np.array([0.60, 0.60, 0.60, 0.55, 0.45, 0.35, 0.28, 0.18, 0.12])
        self._cases = ["Modal", "Fsx_ADE", "Fsy_ADE", "Dead", "Live"]
        self._funcs = ["Espectro_NSR-10", "Espectro_ADEE"]
        self._cur = []
        # Datos de pisos para Ta automático (h, N)
        self._story_names = ["Story3", "Story2", "Story1", "Base"]
        self._story_elevs = [9.0, 6.0, 3.0, 0.0]
        self._story_heights = [3.0, 3.0, 3.0, 0.0]
        self.DatabaseTables = self._DB(self)
        self.Results = self._Results(self)
        self.Func = self._Func(self)
        self.LoadCases = self._LoadCases(self)
        self.File = self._File(self)
        self.Story = self._Story(self)
    def SetPresentUnits(self, u): self._units = u; return 0
    class _DB:
        def __init__(self, p): self.p = p
        def GetAvailableTables(self, *args):
            # Simula lista de tablas disponibles para discovery
            keys = [
                "Mass Summary by Story",
                "Story Definitions",
                "Functions - Response Spectrum - User Defined",
                "Functions - Response Spectrum",
                "Modal Participating Mass Ratios",
            ]
            # Retorna formato (NumberTables, TableKeys, ret) o similar
            return (len(keys), keys, 0)
        def GetTableForDisplayArray(self, t, f, g):
            if "Mass Summary" in t:
                h = list(self.p._mass_df.columns); d = []
                for _, r in self.p._mass_df.iterrows():
                    for c in h: d.append(str(r[c]))
                return ["", h, len(self.p._mass_df), d, 0]
            if "Functions" in t:
                h = ["Name", "Period", "Value"]; d = []
                for n in self.p._funcs: d.extend([n, "0.2", "0.6"])
                return ["", h, len(d)//len(h), d, 0]
            if "Story" in t:
                # Simula Story Definitions con Name, Elevation, Height
                h = ["Story", "Elevation", "Height"]
                d = []
                for n, e, ht in zip(self.p._story_names, self.p._story_elevs, self.p._story_heights):
                    d.extend([n, str(e), str(ht)])
                return ["", h, len(self.p._story_names), d, 0]
            # Para cualquier tabla de espectro que contenga datos reales, simular espectro detallado
            if "Response Spectrum" in t:
                # Si piden tabla específica de espectro que no es User Defined, devolver igual datos
                # Esto permite que fallback vía tabla funcione para nsr_10_Mzl_E.
                h = ["Name", "Period", "Value", "Damping"]
                d = []
                for n in self.p._funcs:
                    # Generar curva más detallada para cada función
                    for per, val in zip(self.p._spec_p, self.p._spec_sa):
                        d.extend([n, str(per), str(val), "0.05"])
                return ["", h, len(d)//len(h), d, 0]
            return ["", [], 0, [], -96]
    class _Story:
        def __init__(self, p): self.p = p
        def GetStories(self):
            # Retorna tupla mezclada simulando ETABS API real
            # Incluimos listas de nombres y elevaciones en posiciones variables
            return (len(self.p._story_names), self.p._story_names, self.p._story_elevs, self.p._story_heights, 0)
    class _Results:
        def __init__(self, p): self.p = p; self.Setup = self._Setup(p)
        class _Setup:
            def __init__(self, p): self.p = p
            def DeselectAllCasesAndCombosForOutput(self): self.p._cur = []; return 0
            def SetCaseSelectedForOutput(self, c): self.p._cur.append(c); return 0
        def ModalParticipatingMassRatios(self):
            n = len(self.p._periods)
            return [n, ["Modal"]*n, ["Mode"]*n, list(range(1, n+1)), list(self.p._periods), list(self.p._ux), list(self.p._uy), [0]*n, [0]*n, [0]*n, [0]*n, [0]*n, [0]*n, 0]
        def BaseReact(self):
            cases = self.p._cur if self.p._cur else ["Fsx_ADE", "Fsy_ADE"]
            fx=[]; fy=[]
            for c in cases:
                if "x" in c.lower(): fx.append(800.0); fy.append(20.0)
                elif "y" in c.lower(): fx.append(15.0); fy.append(750.0)
                else: fx.append(100.0); fy.append(100.0)
            return [len(cases), cases, ["Max"]*len(cases), [0]*len(cases), fx, fy, [0]*len(cases), [0]*len(cases), [0]*len(cases), [0]*len(cases), 0]
    class _Func:
        def __init__(self, p): self.p = p; self.FuncRS = self._FR(p)
        class _FR:
            def __init__(self, p): self.p = p
            def GetCurve(self, name):
                if name not in self.p._funcs: return [name, [], [], 1]
                return [name, list(self.p._spec_p), list(self.p._spec_sa), 0]
            def GetNameList(self, n, names): return [len(self.p._funcs), list(self.p._funcs), 0]
    class _LoadCases:
        def __init__(self, p): self.p = p
        def GetNameList(self, n, names): return [len(self.p._cases), list(self.p._cases), 0]
        def GetTypeOAPI(self, case):
            if "modal" in case.lower(): return [3, 0]
            if "fs" in case.lower() or "ade" in case.lower(): return [4, 0]
            return [1, 0]
    class _File:
        def __init__(self, p): self.p = p
        def GetModelFilename(self, *a): return r"C:\Mock\Edificio_Ejemplo.edb"

# ---------------------------------------------------------------------------
# 4) GUI Tkinter — 3 Vistas (corrección de bugs previos)
# ---------------------------------------------------------------------------
if TK_AVAILABLE:
    class App(tk.Tk):
        def __init__(self, start_mock=False):
            super().__init__()
            self.title("Ajuste de Cortante Basal NSR-10 — ETABS | ADEE")
            self.geometry("1220x780")
            self.minsize(1100, 720)
            # DPI awareness Windows
            try:
                from ctypes import windll; windll.shcore.SetProcessDpiAwareness(1)  # type: ignore
            except Exception:
                pass

            self.SapModel = None
            self.is_mock = start_mock
            self.connected = False
            self.mass_res: MassResult | None = None
            self.modal_res: ModalResult | None = None
            self.spec_p = None; self.spec_sa = None
            self.base_res: BaseResult | None = None
            self.calculo: CalculoResult | None = None

            # Estilo
            style = ttk.Style(self)
            try: style.theme_use("clam")
            except Exception: pass
            style.configure("Title.TLabel", font=("Segoe UI", 13, "bold"), foreground="#1a3a5f")
            style.configure("Subtitle.TLabel", font=("Segoe UI", 8), foreground="#5a6d8a")
            style.configure("Card.TLabelframe", background="white")
            style.configure("Card.TLabelframe.Label", font=("Segoe UI", 9, "bold"), foreground="#1a3a5f")

            self._build_ui()
            if MISSING_DEPS:
                self.after(500, lambda: messagebox.showwarning(
                    "Dependencias faltantes",
                    "Faltan: " + ", ".join(MISSING_DEPS) + "\n\nInstale con:\n pip install pandas numpy openpyxl comtypes\n\nLa app seguirá en modo limitado."))

            if start_mock:
                self.var_mock.set(True)
                self.after(400, self.on_connect)

        # ---------- UI ----------
        def _build_ui(self):
            # Header
            hdr = ttk.Frame(self, padding=10); hdr.pack(fill="x")
            left = ttk.Frame(hdr); left.pack(side="left")
            ttk.Label(left, text="Ajuste de Cortante Basal — NSR-10 A.5.4.5", style="Title.TLabel").pack(anchor="w")
            ttk.Label(left, text="Tkinter • ETABS API (comtypes) • pandas + numpy • Unidades kN-m-C (SetPresentUnits=6)", style="Subtitle.TLabel").pack(anchor="w")
            right = ttk.Frame(hdr); right.pack(side="right")
            self.var_mock = tk.BooleanVar(value=self.is_mock)
            ttk.Checkbutton(right, text="Modo Mock (sin ETABS)", variable=self.var_mock, command=self.on_mock_toggle).pack(side="left", padx=6)
            self.lbl_status = ttk.Label(right, text="● Desconectado", foreground="#c0392b", font=("Segoe UI", 9, "bold"))
            self.lbl_status.pack(side="left", padx=8)
            self.btn_connect = ttk.Button(right, text="🔌 Conectar a ETABS", command=self.on_connect)
            self.btn_connect.pack(side="left")

            # Notebook 3 vistas
            self.nb = ttk.Notebook(self); self.nb.pack(fill="both", expand=True, padx=12, pady=6)
            self.tab1 = ttk.Frame(self.nb, padding=12); self.tab2 = ttk.Frame(self.nb, padding=12); self.tab3 = ttk.Frame(self.nb, padding=12)
            self.nb.add(self.tab1, text="  Vista 1: Conexión y Configuración  ")
            self.nb.add(self.tab2, text="  Vista 2: Tablas Extraídas  ")
            self.nb.add(self.tab3, text="  Vista 3: Cálculos y Resultados  ")
            self._tabs_enabled = [True, False, False]
            self.nb.bind("<<NotebookTabChanged>>", self._on_tab_change)

            self._build_tab1(); self._build_tab2(); self._build_tab3()

            # Log
            logf = ttk.LabelFrame(self, text=" Consola / Log ", padding=6); logf.pack(fill="x", padx=12, pady=(0,8))
            self.txt_log = tk.Text(logf, height=5, font=("Consolas", 8), bg="#0f172a", fg="#e2e8f0", wrap="word")
            self.txt_log.pack(fill="x"); self.txt_log.configure(state="disabled")
            self.log("App iniciada. Conéctese a ETABS para habilitar vistas.")

        def _on_tab_change(self, e):
            try:
                idx = self.nb.index(self.nb.select())
                if not self._tabs_enabled[idx]:
                    self.nb.select(0)
                    # Evitar spam: solo avisar si intenta ir a 2/3 sin extraer
                    if idx in (1,2) and not self.connected:
                        messagebox.showwarning("Bloqueado", "Conéctese y extraiga datos primero.")
            except Exception:
                pass

        def _build_tab1(self):
            f = self.tab1
            # Hacer que Vista1 sea scrolleable (evita que se corte en pantallas pequeñas)
            canvas = tk.Canvas(f, highlightthickness=0)
            sb = ttk.Scrollbar(f, orient="vertical", command=canvas.yview)
            canvas.configure(yscrollcommand=sb.set)
            sb.pack(side="right", fill="y")
            canvas.pack(side="left", fill="both", expand=True)
            inner = ttk.Frame(canvas)
            canvas.create_window((0, 0), window=inner, anchor="nw")
            def _on_conf(e): canvas.configure(scrollregion=canvas.bbox("all"))
            inner.bind("<Configure>", _on_conf)

            # Conexión
            g = ttk.LabelFrame(inner, text=" Conexión ", padding=10); g.pack(fill="x", pady=4)
            self.lbl_model = ttk.Label(g, text="Modelo: — (desconectado)", font=("Segoe UI", 9, "bold")); self.lbl_model.pack(anchor="w")
            ttk.Label(g, text="Unidades API se fijan a kN-m-C (6) automáticamente antes de cada extracción.", font=("Segoe UI", 8), foreground="#64748b").pack(anchor="w")
            ttk.Button(g, text="⟳ Refrescar listas (casos/funciones)", command=self.refresh_combos).pack(anchor="e", pady=4)

            # Selectores (con validación difusa espectro)
            sel = ttk.LabelFrame(inner, text=" Selectores Dinámicos (extraídos de ETABS) ", padding=10); sel.pack(fill="x", pady=6)
            ttk.Label(sel, text="Caso Dinámico X:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
            self.cbo_dx = ttk.Combobox(sel, width=24); self.cbo_dx.grid(row=0, column=1, padx=6, pady=4); self.cbo_dx.set("Fsx_ADE")
            ttk.Label(sel, text="Caso Dinámico Y:").grid(row=1, column=0, sticky="w", padx=6, pady=4)
            self.cbo_dy = ttk.Combobox(sel, width=24); self.cbo_dy.grid(row=1, column=1, padx=6, pady=4); self.cbo_dy.set("Fsy_ADE")
            ttk.Label(sel, text="Caso Modal:").grid(row=0, column=2, sticky="w", padx=6, pady=4)
            self.cbo_modal = ttk.Combobox(sel, width=20); self.cbo_modal.grid(row=0, column=3, padx=6, pady=4); self.cbo_modal.set("Modal")
            ttk.Label(sel, text="Función Espectro:").grid(row=1, column=2, sticky="w", padx=6, pady=4)
            self.cbo_func = ttk.Combobox(sel, width=20); self.cbo_func.grid(row=1, column=3, padx=6, pady=4); self.cbo_func.set("Espectro_NSR-10")
            ttk.Label(sel, text="↳ Las listas se llenan al conectar; ahora tolera mayúsculas/espacios/guiones. Si escribe 'espectro nsr10' lo corrige a 'Espectro_NSR-10'.", font=("Segoe UI", 7, "italic"), foreground="#0d8a4a").grid(row=2, column=0, columnspan=4, sticky="w", padx=6)
            self.lbl_func_info = ttk.Label(sel, text="Función: — (se validará al extraer)", font=("Segoe UI", 7), foreground="#64748b"); self.lbl_func_info.grid(row=3, column=0, columnspan=4, sticky="w", padx=6)
            # Validación difusa en vivo para espectro
            def _on_func_change(e=None):
                typed = self.cbo_func.get().strip()
                if not self.connected or not typed:
                    return
                try:
                    avail = get_spectrum_functions(self.SapModel) if self.SapModel else []
                    if avail:
                        resolved, msg = _resolve_spectrum_name(self.SapModel, typed, avail)
                        if resolved != typed:
                            self.lbl_func_info.configure(text=f"⚠ {msg}", foreground="#b45309")
                        else:
                            self.lbl_func_info.configure(text=f"✓ {msg} — se usará '{resolved}'", foreground="#0d8a4a")
                    else:
                        self.lbl_func_info.configure(text="No hay funciones en modelo para comparar", foreground="#64748b")
                except Exception:
                    pass
                self.validate_tab1()
            self.cbo_func.bind("<<ComboboxSelected>>", _on_func_change)
            self.cbo_func.bind("<KeyRelease>", _on_func_change)
            self.cbo_func.bind("<FocusOut>", _on_func_change)

            # Parámetros manuales (aún editables, pero ahora pueden llenarse automáticamente)
            inp = ttk.LabelFrame(inner, text=" Parámetros de Periodo Aproximado (puede llenarse manual o con cálculo NSR-10) ", padding=10); inp.pack(fill="x", pady=6)
            ttk.Label(inp, text="Ta en X [s]").grid(row=0, column=0, sticky="w", padx=6, pady=4)
            self.ent_ta_x = ttk.Entry(inp, width=10); self.ent_ta_x.grid(row=0, column=1, padx=6, pady=4); self.ent_ta_x.insert(0, "0.746")
            ttk.Label(inp, text="Ta en Y [s]").grid(row=0, column=2, sticky="w", padx=6, pady=4)
            self.ent_ta_y = ttk.Entry(inp, width=10); self.ent_ta_y.grid(row=0, column=3, padx=6, pady=4); self.ent_ta_y.insert(0, "0.746")
            ttk.Label(inp, text="Coeficiente Cu (NSR-10 A.4.2.2)").grid(row=1, column=0, sticky="w", padx=6, pady=4)
            self.ent_cu = ttk.Entry(inp, width=10); self.ent_cu.grid(row=1, column=1, padx=6, pady=4); self.ent_cu.insert(0, "1.2")
            ttk.Label(inp, text="Ej.: 0.746 / 1.2 (o calcule abajo)").grid(row=1, column=2, columnspan=2, sticky="w", padx=6)

            # --- NUEVO: Cálculo Automático NSR-10 A.4.2 ---
            auto = ttk.LabelFrame(inner, text=" Cálculo Automático de Ta y Cu — NSR-10 A.4.2 (extrae altura del modelo) ", padding=10); auto.pack(fill="x", pady=6)
            ttk.Label(auto, text="Este panel calcula Ta = Ct·h^α (A.4.2-3) y Cu = 1.75-1.2·Av·Fv (A.4.2-2) usando datos del modelo. h y N se extraen de ETABS; Ud. elige sistema estructural y Av/Fv.", font=("Segoe UI", 7, "italic"), foreground="#334155", wraplength=700, justify="left").grid(row=0, column=0, columnspan=4, sticky="w", padx=6, pady=(0,6))

            # Fila altura / N
            ttk.Label(auto, text="Altura total h [m]:").grid(row=1, column=0, sticky="w", padx=6, pady=3)
            self.ent_h = ttk.Entry(auto, width=10); self.ent_h.grid(row=1, column=1, sticky="w", padx=6, pady=3)
            self.ent_h.insert(0, "")
            self.ent_h.configure(state="normal")
            ttk.Label(auto, text="N pisos:").grid(row=1, column=2, sticky="w", padx=6, pady=3)
            self.ent_N = ttk.Entry(auto, width=8); self.ent_N.grid(row=1, column=3, sticky="w", padx=6, pady=3)
            self.btn_extract_h = ttk.Button(auto, text="↻ Extraer h y N del modelo", command=self.on_extract_height); self.btn_extract_h.grid(row=1, column=4, padx=6, pady=3)
            ttk.Label(auto, text="hp prom = h/N").grid(row=1, column=5, sticky="w", padx=6)
            self.lbl_hp = ttk.Label(auto, text="—", font=("Segoe UI", 8, "bold"), foreground="#1a3a5f"); self.lbl_hp.grid(row=1, column=6, padx=6)

            # Sistema estructural X/Y
            ttk.Label(auto, text="Sistema X:").grid(row=2, column=0, sticky="w", padx=6, pady=3)
            self.cbo_sistema_x = ttk.Combobox(auto, width=38, values=list(NSR10_CT_ALPHA.keys())); self.cbo_sistema_x.grid(row=2, column=1, columnspan=2, sticky="ew", padx=6, pady=3)
            self.cbo_sistema_x.set("Pórticos concreto (Ct=0.047, α=0.90) — NSR-10 A.4.2-1 #1")
            ttk.Label(auto, text="Sistema Y:").grid(row=2, column=3, sticky="w", padx=6, pady=3)
            self.cbo_sistema_y = ttk.Combobox(auto, width=38, values=list(NSR10_CT_ALPHA.keys())); self.cbo_sistema_y.grid(row=2, column=4, columnspan=3, sticky="ew", padx=6, pady=3)
            self.cbo_sistema_y.set("Pórticos concreto (Ct=0.047, α=0.90) — NSR-10 A.4.2-1 #1")

            # Av / Fv para Cu
            ttk.Label(auto, text="Av (acel. horiz. pico efectiva):").grid(row=3, column=0, sticky="w", padx=6, pady=3)
            self.ent_Av = ttk.Entry(auto, width=10); self.ent_Av.grid(row=3, column=1, sticky="w", padx=6, pady=3); self.ent_Av.insert(0, "")
            ttk.Label(auto, text="Fv (coef. amplif. zona):").grid(row=3, column=2, sticky="w", padx=6, pady=3)
            self.ent_Fv = ttk.Entry(auto, width=10); self.ent_Fv.grid(row=3, column=3, sticky="w", padx=6, pady=3); self.ent_Fv.insert(0, "")
            ttk.Label(auto, text="Ej. Av=0.20, Fv=3.5 (zona alta) → Cu≈1.2").grid(row=3, column=4, columnspan=3, sticky="w", padx=6)

            # Botón calcular
            self.btn_calc_Ta = ttk.Button(auto, text="🧮 Calcular Ta (Ct·h^α) y Cu → llenar arriba", command=self.on_calc_Ta_auto); self.btn_calc_Ta.grid(row=4, column=0, columnspan=2, pady=8, padx=6, sticky="ew")
            self.btn_calc_Ta_02 = ttk.Button(auto, text="Usar Ta = 0.1·N (alt. ≤12 pisos)", command=lambda: self.on_calc_Ta_auto(use_alt=True)); self.btn_calc_Ta_02.grid(row=4, column=2, columnspan=2, pady=8, padx=6, sticky="ew")
            self.lbl_Ta_info = ttk.Label(auto, text="Ta: —  |  Cu: —", font=("Consolas", 8), background="#f8fafc", relief="solid", borderwidth=1, padding=6, wraplength=700, justify="left"); self.lbl_Ta_info.grid(row=5, column=0, columnspan=7, sticky="ew", padx=6, pady=4)
            # Hacer que la columna 1 se expanda
            auto.columnconfigure(1, weight=1); auto.columnconfigure(3, weight=1)

            # Regularidad
            reg = ttk.LabelFrame(inner, text=" Regularidad (NSR-10 A.5.4.5) ", padding=10); reg.pack(fill="x", pady=6)
            self.var_reg = tk.StringVar(value="regular")
            ttk.Radiobutton(reg, text="Edificio Regular  →  80% Vs (0.80)", variable=self.var_reg, value="regular").pack(anchor="w")
            ttk.Radiobutton(reg, text="Edificio Irregular →  90% Vs (0.90)", variable=self.var_reg, value="irregular").pack(anchor="w")
            ttk.Label(reg, text="Define el porcentaje de ajuste para el criterio Vt < porcentaje×Vs.", font=("Segoe UI", 7), foreground="#64748b").pack(anchor="w")

            # Botón continuar
            bf = ttk.Frame(inner); bf.pack(fill="x", pady=10)
            self.lbl_ready = ttk.Label(bf, text="Conéctese y complete campos para continuar →", foreground="#64748b", font=("Segoe UI", 8, "italic")); self.lbl_ready.pack(side="left")
            self.btn_to_tab2 = ttk.Button(bf, text="Extraer Datos → Vista 2: Ver Tablas", command=self.on_extract); self.btn_to_tab2.pack(side="right")
            self.btn_to_tab2.configure(state="disabled")

            # Validación en vivo
            for cbo in (self.cbo_dx, self.cbo_dy, self.cbo_modal, self.cbo_func):
                cbo.bind("<<ComboboxSelected>>", lambda e: self.validate_tab1())
                cbo.bind("<KeyRelease>", lambda e: self.validate_tab1())
            for ent in (self.ent_ta_x, self.ent_ta_y, self.ent_cu):
                ent.bind("<KeyRelease>", lambda e: self.validate_tab1())
            self.after(600, self.validate_tab1)

        def _build_tab2(self):
            f = self.tab2
            ttk.Label(f, text="Verifique las tablas extraídas antes de calcular. Amarillo = máx UX (Tx) • Naranja = máx UY (Ty) • Fila TOTAL en azul.", font=("Segoe UI", 8, "italic"), foreground="#475569").pack(anchor="w", pady=4)
            paned = ttk.PanedWindow(f, orient="horizontal"); paned.pack(fill="both", expand=True, pady=6)

            # Masa
            fr1 = ttk.LabelFrame(paned, text=" Tabla 1 — Mass Summary by Story ", padding=6); paned.add(fr1, weight=1)
            self.tree_masa = ttk.Treeview(fr1, show="headings", height=11); self.tree_masa.pack(fill="both", expand=True, side="left")
            sb1 = ttk.Scrollbar(fr1, orient="vertical", command=self.tree_masa.yview); sb1.pack(side="right", fill="y"); self.tree_masa.configure(yscrollcommand=sb1.set)
            sb1h = ttk.Scrollbar(fr1, orient="horizontal", command=self.tree_masa.xview); sb1h.pack(side="bottom", fill="x"); self.tree_masa.configure(xscrollcommand=sb1h.set)
            self.lbl_masa_total = ttk.Label(fr1, text="Masa total: —", font=("Segoe UI", 9, "bold"), background="#eef4ff", anchor="center"); self.lbl_masa_total.pack(fill="x", pady=4)

            # Modal
            fr2 = ttk.LabelFrame(paned, text=" Tabla 2 — Modal Participating Mass Ratios ", padding=6); paned.add(fr2, weight=1)
            self.tree_modal = ttk.Treeview(fr2, show="headings", height=11); self.tree_modal.pack(fill="both", expand=True, side="left")
            sb2 = ttk.Scrollbar(fr2, orient="vertical", command=self.tree_modal.yview); sb2.pack(side="right", fill="y"); self.tree_modal.configure(yscrollcommand=sb2.set)
            sb2h = ttk.Scrollbar(fr2, orient="horizontal", command=self.tree_modal.xview); sb2h.pack(side="bottom", fill="x"); self.tree_modal.configure(xscrollcommand=sb2h.set)
            self.lbl_modal_info = ttk.Label(fr2, text="Tx (máx UX) = —  |  Ty (máx UY) = —", font=("Segoe UI", 9, "bold"), background="#fff3e0", anchor="center"); self.lbl_modal_info.pack(fill="x", pady=4)

            bf = ttk.Frame(f); bf.pack(fill="x", pady=6)
            ttk.Button(bf, text="⟳ Re-extraer Tablas", command=self.on_extract).pack(side="left")
            self.btn_to_tab3 = ttk.Button(bf, text="Calcular Ajuste → Vista 3: Resultados", command=self.on_calculate); self.btn_to_tab3.pack(side="right"); self.btn_to_tab3.configure(state="disabled")

        def _build_tab3(self):
            f = self.tab3
            canvas = tk.Canvas(f, highlightthickness=0); sb = ttk.Scrollbar(f, orient="vertical", command=canvas.yview)
            canvas.configure(yscrollcommand=sb.set); sb.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
            inner = ttk.Frame(canvas); canvas.create_window((0, 0), window=inner, anchor="nw")
            def _conf(e): canvas.configure(scrollregion=canvas.bbox("all"))
            inner.bind("<Configure>", _conf)

            ttk.Label(inner, text="Cálculos Paso a Paso — Trazabilidad tipo Memoria de Cálculo", font=("Segoe UI", 11, "bold"), foreground="#1a3a5f").pack(anchor="w", pady=4)
            ttk.Label(inner, text="Cada paso replica la hoja validada. Verifique fórmulas y valores antes de aplicar el factor en ETABS.", font=("Segoe UI", 8), foreground="#475569").pack(anchor="w")

            self.lbl_steps = {}
            for title in [
                "Paso 1 — Periodos (Tmax = Ta × Cu,  T = min(Tmax, Tmodal))",
                "Paso 2 — Aceleración Espectral Sa (numpy.interp)",
                "Paso 3 — Cortante Estático Vs = Masa × Sa × g / 1000",
                "Paso 4 — Cortante Dinámico Vt (BaseReact ETABS)",
                "Paso 5 — Criterio de Ajuste (NSR-10 A.5.4.5)",
            ]:
                lf = ttk.LabelFrame(inner, text=f" {title} ", padding=8); lf.pack(fill="x", pady=6, padx=4)
                lbl = ttk.Label(lf, text="—", font=("Consolas", 8), justify="left", wraplength=680); lbl.pack(anchor="w", fill="x")
                self.lbl_steps[title] = lbl

            # Resultado final destacado
            res = ttk.LabelFrame(inner, text=" 🎯  RESULTADO FINAL — FACTOR DE ESCALA ", padding=12); res.pack(fill="x", pady=10, padx=4)
            self.lbl_factor_x = ttk.Label(res, text="Factor X: —", font=("Segoe UI", 12, "bold"), foreground="#1a3a5f", anchor="center", background="#fff7e6", padding=8); self.lbl_factor_x.pack(fill="x", pady=3)
            self.lbl_factor_y = ttk.Label(res, text="Factor Y: —", font=("Segoe UI", 12, "bold"), foreground="#1a3a5f", anchor="center", background="#fff7e6", padding=8); self.lbl_factor_y.pack(fill="x", pady=3)
            ttk.Label(res, text="Ingrese el factor en ETABS: Define > Load Cases > Response Spectrum > Scale Factor. Si “No Aplica” use 1.0.", font=("Segoe UI", 7), foreground="#57534e", wraplength=650, justify="center").pack(fill="x", pady=3)
            self.lbl_resumen = ttk.Label(res, text="—", font=("Segoe UI", 8), wraplength=650, justify="left", background="#fef3c7", padding=6); self.lbl_resumen.pack(fill="x", pady=4)

            exp = ttk.Frame(inner); exp.pack(fill="x", pady=6)
            self.btn_export_txt = ttk.Button(exp, text="📄 Exportar Memoria (.txt)", command=self.export_txt); self.btn_export_txt.pack(side="left", padx=4); self.btn_export_txt.configure(state="disabled")
            self.btn_export_excel = ttk.Button(exp, text="📊 Exportar Tablas (.xlsx)", command=self.export_excel); self.btn_export_excel.pack(side="left", padx=4); self.btn_export_excel.configure(state="disabled")
            self.btn_copy = ttk.Button(exp, text="📋 Copiar Factores", command=self.copy_factors); self.btn_copy.pack(side="left", padx=4); self.btn_copy.configure(state="disabled")

            # Info unidades
            uni = ttk.LabelFrame(inner, text=" Unidades y Trazabilidad ", padding=8); uni.pack(fill="x", padx=4, pady=6)
            ttk.Label(uni, text="• Unidades API fijadas a kN-m-C (SetPresentUnits=6) • Masa ETABS: kN·s²/m (Ton) → ×1000 = kg • Vs[kN]=M[kg]×Sa[g]×9.80665/1000 • Sa con numpy.interp • T_aju=min(Ta×Cu,Tmodal) • Factor=(porcentaje×Vs)/Vt (0.80 regular, 0.90 irregular)", font=("Segoe UI", 7), foreground="#334155", wraplength=680, justify="left").pack(anchor="w")

            bf = ttk.Frame(inner); bf.pack(fill="x", pady=4)
            self.btn_recalc = ttk.Button(bf, text="⟳ Recalcular con nuevos Ta/Cu", command=self.on_calculate); self.btn_recalc.pack(side="right"); self.btn_recalc.configure(state="disabled")

        # ---------- Log ----------
        def log(self, msg):
            ts = datetime.datetime.now().strftime("%H:%M:%S")
            self.txt_log.configure(state="normal"); self.txt_log.insert("end", f"[{ts}] {msg}\n"); self.txt_log.see("end"); self.txt_log.configure(state="disabled")
            print(f"[{ts}] {msg}")

        # ---------- Eventos ----------
        def on_mock_toggle(self):
            if self.var_mock.get(): self.log("Modo Mock activado (simulado, sin ETABS).")
            else: self.log("Modo Mock desactivado (ETABS real).")
            if self.connected and self.is_mock != self.var_mock.get():
                self.connected = False; self.SapModel = None
                self.lbl_status.configure(text="● Desconectado", foreground="#c0392b")
                self._tabs_enabled[1] = self._tabs_enabled[2] = False
                self.btn_to_tab2.configure(state="disabled"); self.lbl_model.configure(text="Modelo: — (desconectado)")

        def on_connect(self):
            mock = self.var_mock.get(); self.is_mock = mock
            if mock:
                if pd is None or np is None:
                    messagebox.showerror("Faltan dependencias", "Modo Mock requiere pandas y numpy.\n\npip install pandas numpy")
                    return
                try:
                    self.SapModel = MockSapModel()
                    self.connected = True
                    self.lbl_status.configure(text="● Conectado (Mock)", foreground="#0d8a4a")
                    self.lbl_model.configure(text="Modelo: C:\\Mock\\Edificio_Ejemplo.edb (Mock)")
                    self.log("✅ Conectado Mock — unidades kN-m-C (6)."); self.refresh_combos()
                except Exception as e:
                    self.log(f"❌ Mock error: {e}"); messagebox.showerror("Error Mock", f"{e}\n\n{traceback.format_exc()}")
                return
            # Conexión real
            if pd is None or np is None:
                messagebox.showerror("Faltan dependencias", "Requiere pandas y numpy.\n\npip install pandas numpy")
                return
            try:
                self.log("Conectando a ETABS (CSI.ETABS.API.ETABSObject)...")
                SapModel, EtabsObject = connect_to_etabs()
                self.SapModel = SapModel; self.connected = True
                try:
                    ruta = SapModel.GetModelFilename()
                    ruta = str(ruta[1] if isinstance(ruta, (list, tuple)) and len(ruta) > 1 else ruta)
                except Exception: ruta = "—"
                try: SapModel.SetPresentUnits(6)
                except Exception: pass
                self.lbl_model.configure(text=f"Modelo: {ruta}")
                self.lbl_status.configure(text="● Conectado a ETABS", foreground="#0d8a4a")
                self.log(f"✅ Conectado ETABS: {ruta}"); self.refresh_combos()
            except RuntimeError as e:
                self.log(f"❌ {e}"); messagebox.showerror("Error de Conexión", str(e) + "\n\nActive 'Modo Mock' para pruebas sin ETABS.")
            except Exception as e:
                self.log(f"❌ Error inesperado: {e}"); messagebox.showerror("Error", f"{e}\n\n{traceback.format_exc()}")

        def refresh_combos(self):
            if not self.connected or self.SapModel is None:
                self.log("No conectado, no se pueden refrescar listas.")
                return
            try:
                rs = get_response_spectrum_cases(self.SapModel)
                mods = get_modal_cases(self.SapModel)
                allc = get_all_load_cases(self.SapModel)
                funcs = get_spectrum_functions(self.SapModel)
                self.log(f"Casos Totales: {allc}")
                self.log(f"RS: {rs} | Modal: {mods} | Funciones: {funcs}")
                def set_combo(cbo, lst):
                    cur = cbo.get()
                    cbo["values"] = lst if lst else []
                    if cur in lst: cbo.set(cur)
                    elif lst: cbo.set(lst[0])
                set_combo(self.cbo_dx, rs if rs else allc)
                set_combo(self.cbo_dy, rs if rs else allc)
                set_combo(self.cbo_modal, mods if mods else allc)
                set_combo(self.cbo_func, funcs)
                # Heurística defaults
                if self.is_mock:
                    if not self.cbo_dx.get(): self.cbo_dx.set("Fsx_ADE")
                    if not self.cbo_dy.get(): self.cbo_dy.set("Fsy_ADE")
                    if not self.cbo_modal.get(): self.cbo_modal.set("Modal")
                    if not self.cbo_func.get(): self.cbo_func.set("Espectro_NSR-10")
                else:
                    # Intentar asignar X/Y automáticamente si hay varios
                    if rs and len(rs) >= 2:
                        # Si los combos quedaron iguales, separar
                        if self.cbo_dx.get() == self.cbo_dy.get():
                            xs = [c for c in rs if "x" in c.lower()]
                            ys = [c for c in rs if "y" in c.lower()]
                            if xs: self.cbo_dx.set(xs[0])
                            if ys: self.cbo_dy.set(ys[0])
                self.validate_tab1()
                self.log("✓ Comboboxes actualizados.")
            except Exception as e:
                self.log(f"❌ Error refrescando: {e}\n{traceback.format_exc()}")

        def validate_tab1(self):
            if not self.connected:
                self.btn_to_tab2.configure(state="disabled"); self.lbl_ready.configure(text="Conéctese a ETABS para continuar →", foreground="#64748b")
                return
            ok = all([self.cbo_dx.get().strip(), self.cbo_dy.get().strip(), self.cbo_modal.get().strip(), self.cbo_func.get().strip()])
            try:
                ta_x = float(self.ent_ta_x.get()); ta_y = float(self.ent_ta_y.get()); cu = float(self.ent_cu.get())
                ok = ok and ta_x > 0 and ta_y > 0 and cu > 0
            except Exception: ok = False
            self.btn_to_tab2.configure(state="normal" if ok else "disabled")
            if ok: self.lbl_ready.configure(text="✓ Listo para extraer datos →", foreground="#0d8a4a")
            else: self.lbl_ready.configure(text="Complete todos los campos (casos, función, Ta, Cu) →", foreground="#64748b")

        # ---------- Cálculo automático Ta NSR-10 ----------
        def on_extract_height(self):
            """Extrae h y N del modelo ETABS (Story Data) y los pone en los entries."""
            if not self.connected or self.SapModel is None:
                messagebox.showwarning("No conectado", "Conéctese a ETABS primero (o use Mock).")
                return
            try:
                self.log("Extrayendo altura y número de pisos del modelo...")
                h, N, hp = get_building_height_and_stories(self.SapModel)
                # Actualizar entries
                for ent, val in [(self.ent_h, f"{h:.2f}"), (self.ent_N, f"{N}")]:
                    ent.delete(0, tk.END); ent.insert(0, val)
                self.lbl_hp.configure(text=f"{hp:.2f} m")
                self.log(f"✅ Altura extraída: h={h:.2f} m, N={N} pisos, hp={hp:.2f} m")
                # Mostrar detalle en Ta_info
                self.lbl_Ta_info.configure(text=f"h={h:.2f}m, N={N}, hp={hp:.2f}m — seleccione sistema y calcule Ta", foreground="#1a3a5f")
                messagebox.showinfo("Altura extraída",
                    f"Altura total h = {h:.2f} m\n"
                    f"Número de pisos N = {N}\n"
                    f"Altura promedio hp = {hp:.2f} m\n\n"
                    f"Revise y corrija manualmente si es necesario (por ejemplo, si ETABS cuenta Base como piso).\n"
                    f"Luego seleccione sistema estructural y pulse 'Calcular Ta'.")
            except Exception as e:
                self.log(f"❌ No se pudo extraer altura: {e}")
                # Permitir ingreso manual
                messagebox.showerror("Error extrayendo altura",
                    f"{e}\n\n"
                    f"Ingrese h [m] y N manualmente en los campos y luego calcule Ta.\n"
                    f"Tip: en ETABS verifique Story Data (elevaciones) o Display > Show Tables > Story Definitions.")
                # Enfocar entry h
                try: self.ent_h.focus_set()
                except Exception: pass

        def on_calc_Ta_auto(self, use_alt: bool = False):
            """Calcula Ta = Ct·h^α (y Cu) y lo vuelca a los entries Ta/Cu."""
            # Validar h/N
            try:
                h_txt = self.ent_h.get().strip().replace(",", ".")
                N_txt = self.ent_N.get().strip()
                if not h_txt:
                    raise ValueError("Ingrese h o use 'Extraer h y N del modelo'")
                h = float(h_txt)
                N = int(float(N_txt)) if N_txt else None
                if h <= 0:
                    raise ValueError("h debe ser >0")
                if N is not None and N <= 0:
                    N = None
            except Exception as e:
                messagebox.showwarning("Datos inválidos", f"Altura/N inválidos: {e}\n\nIngrese h [m] numérico y N entero.")
                return

            # Sistemas
            sis_x = self.cbo_sistema_x.get().strip()
            sis_y = self.cbo_sistema_y.get().strip()
            if not sis_x or not sis_y:
                messagebox.showwarning("Falta sistema", "Seleccione sistema estructural para X y Y.")
                return

            try:
                if use_alt:
                    # Alternativa 0.1*N (NSR-10 A.4.2-5) — solo si N<=12 y hp<=3m
                    if N is None:
                        raise ValueError("Para Ta=0.1·N necesita N (número de pisos)")
                    hp = h / N if N else 0
                    if N > 12:
                        messagebox.showwarning("No aplica", f"Alternativa 0.1·N solo para N≤12 (actual N={N})")
                        return
                    if hp > 3.05:
                        if not messagebox.askyesno("Advertencia hp", f"hp promedio = {hp:.2f}m >3.0m, NSR-10 A.4.2-5 no aplica estrictamente. ¿Usar igual 0.1·N?"):
                            return
                    Ta_x = Ta_y = 0.1 * N
                    det_x = f"Ta = 0.1·N = 0.1·{N} = {Ta_x:.4f} s  (A.4.2-5, hp={hp:.2f}m)"
                    det_y = det_x
                else:
                    Ta_x, det_x = calc_Ta_nsr10(h, sis_x, N)
                    Ta_y, det_y = calc_Ta_nsr10(h, sis_y, N)
                # Actualizar entries Ta
                for ent, val in [(self.ent_ta_x, f"{Ta_x:.4f}"), (self.ent_ta_y, f"{Ta_y:.4f}")]:
                    ent.delete(0, tk.END); ent.insert(0, val)

                # Cu: si Av/Fv provistos, calcular; si no, mantener manual
                Av_txt = self.ent_Av.get().strip().replace(",", ".")
                Fv_txt = self.ent_Fv.get().strip().replace(",", ".")
                cu_msg = ""
                if Av_txt and Fv_txt:
                    try:
                        Av = float(Av_txt); Fv = float(Fv_txt)
                        Cu, cu_det = calc_Cu_nsr10(Av, Fv)
                        self.ent_cu.delete(0, tk.END); self.ent_cu.insert(0, f"{Cu:.4f}")
                        cu_msg = f"\n{cu_det}"
                    except Exception as e:
                        cu_msg = f"\nCu: error Av/Fv ({e}) — manteniendo Cu manual"
                else:
                    # Si no hay Av/Fv, intentar sugerir Cu típico? Dejar manual
                    cu_msg = "\nCu: manteniendo valor manual (ingrese Av/Fv para calcular Cu = 1.75-1.2·Av·Fv)"

                info = f"Ta_X = {Ta_x:.4f}s  → {det_x}\nTa_Y = {Ta_y:.4f}s  → {det_y}{cu_msg}\n\nLos valores se han copiado a los campos Ta X/Y y Cu arriba. Verifique y ajuste si su sistema es diferente por dirección."
                self.lbl_Ta_info.configure(text=info, foreground="#0d4b2b")
                self.log(f"✅ Ta calculado: X={Ta_x:.4f}s ({sis_x}), Y={Ta_y:.4f}s ({sis_y}), h={h:.2f}m, N={N}{cu_msg.split(chr(10))[0] if cu_msg else ''}")
                self.validate_tab1()
                # Resaltar que se actualizó
                self.lbl_ready.configure(text="✓ Ta/Cu actualizados automáticamente — verifique y extraiga tablas →", foreground="#0d8a4a")
            except Exception as e:
                self.log(f"❌ Error calculando Ta: {e}\n{traceback.format_exc()}")
                messagebox.showerror("Error Ta", f"{e}\n\n{traceback.format_exc()}")

        # ---------- Tree helpers ----------
        def _fill_tree(self, tree, df):
            # Limpiar
            for c in tree["columns"]: tree.heading(c, text="")
            tree.delete(*tree.get_children())
            if df is None or df.empty:
                tree["columns"] = []
                return
            tree["columns"] = list(df.columns)
            for col in df.columns:
                tree.heading(col, text=str(col))
                # Ancho según contenido
                tree.column(col, width=120, anchor="center", stretch=True)
            for _, row in df.iterrows():
                vals = []
                for v in row.values:
                    if isinstance(v, (int, float)) or (np is not None and isinstance(v, np.floating)):
                        # Formateo
                        try:
                            if abs(float(v)) >= 1000: vals.append(f"{float(v):,.2f}")
                            else: vals.append(f"{float(v):.4f}")
                        except Exception: vals.append(str(v))
                    else:
                        vals.append(str(v))
                tree.insert("", "end", values=vals)
            # Ajustar ancho
            for col in df.columns:
                tree.column(col, width=110)

        # ---------- Vista1 -> Vista2 ----------
        def on_extract(self):
            if not self.connected or self.SapModel is None:
                messagebox.showwarning("No conectado", "Conéctese a ETABS primero."); return
            if pd is None or np is None:
                messagebox.showerror("Faltan dependencias", "Requiere pandas y numpy."); return
            caso_dx = self.cbo_dx.get().strip(); caso_dy = self.cbo_dy.get().strip()
            caso_m = self.cbo_modal.get().strip(); func = self.cbo_func.get().strip()
            if not all([caso_dx, caso_dy, caso_m, func]):
                messagebox.showwarning("Faltan datos", "Seleccione todos los casos y la función."); return
            try:
                self.log("Extrayendo masa (Mass Summary by Story)...")
                self.btn_to_tab2.configure(state="disabled", text="⏳ Extrayendo...")
                self.update_idletasks()
                mass = extract_mass_summary(self.SapModel); self.mass_res = mass
                self.log(f"✅ Masa total: {mass.masa_total_ton:.4f} Ton ({mass.masa_total_kg:,.2f} kg)")
                self._fill_tree(self.tree_masa, mass.df)
                # Fila TOTAL azul
                tot_vals = []
                for col in mass.df.columns:
                    if col == mass.df.columns[0]: tot_vals.append("TOTAL")
                    elif col == mass.col_masa: tot_vals.append(f"{mass.masa_total_ton:.4f} Ton / {mass.masa_total_kg:,.2f} kg")
                    else:
                        if "mass" in col.lower():
                            try: tot_vals.append(f"{pd.to_numeric(mass.df[col], errors='coerce').sum():.2f}")
                            except Exception: tot_vals.append("")
                        else: tot_vals.append("")
                iid = self.tree_masa.insert("", "end", values=tot_vals, tags=("total",))
                self.tree_masa.tag_configure("total", background="#dceffd", font=("Segoe UI", 9, "bold"))
                self.lbl_masa_total.configure(text=f"Masa total: {mass.masa_total_ton:.4f} Ton  ({mass.masa_total_kg:,.2f} kg)  | Columna: {mass.col_masa}")

                self.log(f"Extrayendo modal '{caso_m}'...")
                modal = extract_modal_data(self.SapModel, caso_m); self.modal_res = modal
                self.log(f"✅ Modal Tx={modal.tx_modal:.4f}s (UX {modal.ux[modal.idx_max_ux]:.3f}), Ty={modal.ty_modal:.4f}s (UY {modal.uy[modal.idx_max_uy]:.3f})")
                self._fill_tree(self.tree_modal, modal.df)
                for i, iid in enumerate(self.tree_modal.get_children()):
                    if i == modal.idx_max_ux and i == modal.idx_max_uy:
                        self.tree_modal.item(iid, tags=("both",))
                    elif i == modal.idx_max_ux:
                        self.tree_modal.item(iid, tags=("max_ux",))
                    elif i == modal.idx_max_uy:
                        self.tree_modal.item(iid, tags=("max_uy",))
                self.tree_modal.tag_configure("max_ux", background="#fff3cd")
                self.tree_modal.tag_configure("max_uy", background="#ffe9c6")
                self.tree_modal.tag_configure("both", background="#ffd699")

                self.lbl_modal_info.configure(text=f"Tx (máx UX) = {modal.tx_modal:.4f}s  (Modo {modal.idx_max_ux+1}, UX={modal.ux[modal.idx_max_ux]:.4f})  |  Ty (máx UY) = {modal.ty_modal:.4f}s  (Modo {modal.idx_max_uy+1}, UY={modal.uy[modal.idx_max_uy]:.4f})")

                # Cache espectro — con manejo difuso y reporte de corrección
                try:
                    p, sa = extract_spectrum_curve(self.SapModel, func)
                    self.spec_p, self.spec_sa = p, sa
                    used = getattr(extract_spectrum_curve, "_last_used_name", func)
                    msg = getattr(extract_spectrum_curve, "_last_resolve_msg", "")
                    if used != func:
                        self.log(f"⚠ Función espectro corregida difusa: '{func}' → '{used}' ({msg})")
                        # Actualizar combobox para que el cálculo use el nombre real de ETABS
                        try: self.cbo_func.set(used)
                        except Exception: pass
                        self.lbl_func_info.configure(text=f"⚠ Corregido: '{func}' → '{used}'", foreground="#b45309")
                        func = used  # usar el corregido para resto del flujo
                        self.spec_p, self.spec_sa = p, sa
                    else:
                        self.log(f"✅ Curva '{used}': {len(p)} puntos (0–{p.max():.2f}s) — {msg}")
                        self.lbl_func_info.configure(text=f"✓ Función OK: '{used}'", foreground="#0d8a4a")
                    # Guardar nombre usado para cálculos
                    self._last_func_used = used
                except Exception as e:
                    self.log(f"⚠ Curva no cargada: {e}")
                    # Mostrar disponibles para ayudar
                    try:
                        avail = get_spectrum_functions(self.SapModel)
                        if avail:
                            self.log(f"  Funciones disponibles: {avail}")
                            self.lbl_func_info.configure(text=f"❌ '{func}' no encontrada. Disponibles: {', '.join(avail[:3])}", foreground="#c0392b")
                    except Exception:
                        pass

                self._tabs_enabled[1] = True; self.nb.select(self.tab2)
                self.btn_to_tab3.configure(state="normal")
                self.log("✓ Extracción completa. Verifique tablas en Vista 2.")
            except Exception as e:
                self.log(f"❌ Error extrayendo: {e}\n{traceback.format_exc()}")
                messagebox.showerror("Error de Extracción", f"{e}\n\n{traceback.format_exc()}")
            finally:
                try: self.btn_to_tab2.configure(text="Extraer Datos → Vista 2: Ver Tablas")
                except Exception: pass
                self.validate_tab1()

        # ---------- Vista2 -> Vista3 ----------
        def on_calculate(self):
            if self.mass_res is None or self.modal_res is None:
                messagebox.showwarning("Faltan tablas", "Extraiga primero las tablas en Vista 2."); return
            if pd is None or np is None:
                messagebox.showerror("Faltan dependencias", "Requiere pandas/numpy"); return
            try:
                Ta_x = float(self.ent_ta_x.get()); Ta_y = float(self.ent_ta_y.get()); Cu = float(self.ent_cu.get())
            except Exception:
                messagebox.showwarning("Datos inválidos", "Ta y Cu deben ser numéricos >0."); return
            porcentaje = 0.80 if self.var_reg.get() == "regular" else 0.90
            tipo = "Regular" if porcentaje == 0.80 else "Irregular"
            caso_dx = self.cbo_dx.get().strip(); caso_dy = self.cbo_dy.get().strip(); func = self.cbo_func.get().strip()

            try:
                self.log(f"Iniciando cálculos: Ta_x={Ta_x}, Ta_y={Ta_y}, Cu={Cu}, {tipo} ({porcentaje*100:.0f}%)")
                self.btn_to_tab3.configure(state="disabled", text="⏳ Calculando..."); self.update_idletasks()

                # Espectro: recargar siempre con manejo difuso (asegura nombre correcto)
                # Si func fue corregida en on_extract, también se corrige aquí
                try:
                    p, sa = extract_spectrum_curve(self.SapModel, func)
                    used = getattr(extract_spectrum_curve, "_last_used_name", func)
                    msg = getattr(extract_spectrum_curve, "_last_resolve_msg", "")
                    if used != func:
                        self.log(f"⚠ Función espectro corregida en cálculo: '{func}' → '{used}' ({msg})")
                        try: self.cbo_func.set(used)
                        except Exception: pass
                        func = used
                    self.spec_p, self.spec_sa = p, sa
                    self._last_func = func
                    self._last_func_used = func
                except Exception as e:
                    # Si falla, intentar con cache previo si existe
                    if self.spec_p is not None and self.spec_sa is not None:
                        self.log(f"⚠ No se pudo recargar curva '{func}': {e} — usando cache previo")
                        p, sa = self.spec_p, self.spec_sa
                    else:
                        raise

                Txm = float(self.modal_res.tx_modal); Tym = float(self.modal_res.ty_modal)
                Tmax_x = Ta_x * Cu; Tmax_y = Ta_y * Cu
                Taju_x = min(Tmax_x, Txm); Taju_y = min(Tmax_y, Tym)
                Sa_x = float(np.interp(Taju_x, p, sa)); Sa_y = float(np.interp(Taju_y, p, sa))
                self.log(f"Sa: Sa_x={Sa_x:.5f}g @ {Taju_x:.4f}s, Sa_y={Sa_y:.5f}g @ {Taju_y:.4f}s")

                self.log(f"Extrayendo BaseReact: {caso_dx} (FX), {caso_dy} (FY)...")
                base = extract_base_reactions(self.SapModel, caso_dx, caso_dy); self.base_res = base
                self.log(f"✅ VtX={base.VtX:.2f} kN, VtY={base.VtY:.2f} kN")

                calc = calcular_ajuste(Ta_x, Ta_y, Cu, porcentaje, tipo, float(self.mass_res.masa_total_kg), Txm, Tym, Sa_x, Sa_y, float(base.VtX), float(base.VtY), p, sa, func, caso_dx, caso_dy)
                self.calculo = calc
                self.log(f"Cálculo: VsX={calc.VsX:.1f} kN, VsY={calc.VsY:.1f} kN, FactorX={calc.factor_x}, FactorY={calc.factor_y}")
                self.show_calculations(calc)
                self._tabs_enabled[2] = True; self.nb.select(self.tab3)
                self.btn_export_txt.configure(state="normal"); self.btn_export_excel.configure(state="normal"); self.btn_copy.configure(state="normal"); self.btn_recalc.configure(state="normal")
            except Exception as e:
                self.log(f"❌ Error cálculo: {e}\n{traceback.format_exc()}")
                messagebox.showerror("Error de Cálculo", f"{e}\n\n{traceback.format_exc()}")
            finally:
                try: self.btn_to_tab3.configure(text="Calcular Ajuste → Vista 3: Resultados", state="normal")
                except Exception: pass

        def show_calculations(self, c: CalculoResult):
            # Paso1
            self.lbl_steps["Paso 1 — Periodos (Tmax = Ta × Cu,  T = min(Tmax, Tmodal))"].configure(text=
                f"Ta_x={c.Ta_x:.4f}s, Ta_y={c.Ta_y:.4f}s, Cu={c.Cu:.4f}\n"
                f"Tmax_x=Ta_x×Cu={c.Ta_x:.4f}×{c.Cu:.4f}={c.Tmax_x:.4f}s\n"
                f"Tmax_y=Ta_y×Cu={c.Ta_y:.4f}×{c.Cu:.4f}={c.Tmax_y:.4f}s\n"
                f"Tx_modal (máx UX)={c.Tx_modal:.5f}s  (Modo {self.modal_res.idx_max_ux+1}, UX={self.modal_res.ux[self.modal_res.idx_max_ux]:.4f})\n"
                f"Ty_modal (máx UY)={c.Ty_modal:.5f}s  (Modo {self.modal_res.idx_max_uy+1}, UY={self.modal_res.uy[self.modal_res.idx_max_uy]:.4f})\n"
                f"T_ajustado_x = min({c.Tmax_x:.4f}, {c.Tx_modal:.5f}) = {c.T_aju_x:.5f}s\n"
                f"T_ajustado_y = min({c.Tmax_y:.4f}, {c.Ty_modal:.5f}) = {c.T_aju_y:.5f}s"
            )
            # Paso2
            # Vecinos para trazabilidad
            def _vecinos(T):
                idx = int(np.searchsorted(c.periods, T))
                vals=[]
                for j in (idx-1, idx, idx+1):
                    if 0 <= j < len(c.periods):
                        vals.append(f"({c.periods[j]:.3f}s,{c.sa[j]:.4f}g)")
                return " ".join(vals)
            self.lbl_steps["Paso 2 — Aceleración Espectral Sa (numpy.interp)"].configure(text=
                f"Función ETABS: {c.func_name}  |  {len(c.periods)} puntos  T[{c.periods.min():.2f}-{c.periods.max():.2f}]s\n"
                f"Método: numpy.interp(T_ajustado, periods, sa)\n"
                f"Sa_x = interp({c.T_aju_x:.5f}s) = {c.Sa_x:.6f} g   vecinos: {_vecinos(c.T_aju_x)}\n"
                f"Sa_y = interp({c.T_aju_y:.5f}s) = {c.Sa_y:.6f} g   vecinos: {_vecinos(c.T_aju_y)}"
            )
            # Paso3
            self.lbl_steps["Paso 3 — Cortante Estático Vs = Masa × Sa × g / 1000"].configure(text=
                f"Masa total = {c.masa_total_kg:,.2f} kg  ({c.masa_total_ton:.4f} Ton)\n"
                f"g = 9.80665 m/s²\n"
                f"VsX = {c.masa_total_kg:,.2f} × {c.Sa_x:.6f} × 9.80665 /1000 = {c.VsX:,.4f} kN\n"
                f"VsY = {c.masa_total_kg:,.2f} × {c.Sa_y:.6f} × 9.80665 /1000 = {c.VsY:,.4f} kN"
            )
            # Paso4
            df = self.base_res.df if self.base_res else None
            det = ""
            if df is not None and not df.empty:
                det = "\n".join(f"{str(r['Load Case'])}: FX={float(r['FX']):.1f} kN, FY={float(r['FY']):.1f} kN" for _, r in df.iterrows())
            self.lbl_steps["Paso 4 — Cortante Dinámico Vt (BaseReact ETABS)"].configure(text=
                f"Caso X: {c.caso_dx} → VtX = |FX|max = {c.VtX:,.4f} kN\n"
                f"Caso Y: {c.caso_dy} → VtY = |FY|max = {c.VtY:,.4f} kN\n"
                f"Detalle BaseReact (todas las filas extraídas):\n{det}\n"
                f"Unidades: kN (SetPresentUnits=6)."
            )
            # Paso5
            self.lbl_steps["Paso 5 — Criterio de Ajuste (NSR-10 A.5.4.5)"].configure(text=
                f"Tipo: {c.tipo}  →  porcentaje = {c.porcentaje*100:.0f}% ({c.porcentaje:.2f})\n"
                f"req_X = {c.porcentaje:.2f} × VsX = {c.porcentaje:.2f}×{c.VsX:,.2f} = {c.req_X:,.4f} kN\n"
                f"req_Y = {c.porcentaje:.2f} × VsY = {c.porcentaje:.2f}×{c.VsY:,.2f} = {c.req_Y:,.4f} kN\n"
                f"¿VtX ({c.VtX:,.2f}) < req_X ({c.req_X:,.2f})?  →  {'SÍ — REQUIERE AJUSTE' if c.necesita_x else 'NO — NO REQUIERE'}\n"
                f"¿VtY ({c.VtY:,.2f}) < req_Y ({c.req_Y:,.2f})?  →  {'SÍ — REQUIERE AJUSTE' if c.necesita_y else 'NO — NO REQUIERE'}\n"
                f"Fórmula: Factor = (porcentaje × Vs) / Vt"
            )
            fx = f"{c.factor_x:.4f}" if not isinstance(c.factor_x, str) else "No Aplica (1.0)"
            fy = f"{c.factor_y:.4f}" if not isinstance(c.factor_y, str) else "No Aplica (1.0)"
            self.lbl_factor_x.configure(text=f"Factor X ({c.caso_dx}):  {fx}")
            self.lbl_factor_y.configure(text=f"Factor Y ({c.caso_dy}):  {fy}")
            self.lbl_resumen.configure(text=
                f"Resumen: X Vt {c.VtX:,.0f} kN vs {c.porcentaje:.0%}×Vs {c.req_X:,.0f} kN → {'AJUSTAR' if c.necesita_x else 'OK'} (Factor {fx})  |  "
                f"Y Vt {c.VtY:,.0f} kN vs {c.porcentaje:.0%}×Vs {c.req_Y:,.0f} kN → {'AJUSTAR' if c.necesita_y else 'OK'} (Factor {fy})\n"
                f"Función: {c.func_name}  •  Masa: {c.masa_total_kg:,.0f} kg  •  T_aju X {c.T_aju_x:.3f}s / Y {c.T_aju_y:.3f}s"
            )

        # ---------- Export ----------
        def export_txt(self):
            if self.calculo is None: messagebox.showwarning("Sin datos", "Calcule primero."); return
            c = self.calculo
            path = filedialog.asksaveasfilename(defaultextension=".txt", initialfile=f"Memoria_ADEE_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.txt", filetypes=[("Texto","*.txt")])
            if not path: return
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write("="*68 + "\n MEMORIA DE CÁLCULO — AJUSTE CORTANTE BASAL NSR-10 A.5.4.5\n" + "="*68 + "\n")
                    f.write(f"Fecha: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                    f.write(f"Modelo: {self.lbl_model.cget('text')}\n")
                    f.write(f"Función espectro: {c.func_name}\n")
                    f.write(f"Casos: X={c.caso_dx}, Y={c.caso_dy}, Modal={self.cbo_modal.get().strip()}\n")
                    f.write(f"Tipo: {c.tipo} ({c.porcentaje*100:.0f}%)\n\n")
                    f.write("I. PERIODOS\n" + "-"*68 + "\n")
                    f.write(f"Ta_x={c.Ta_x:.6f}s, Ta_y={c.Ta_y:.6f}s, Cu={c.Cu:.4f}\n")
                    f.write(f"Tmax_x={c.Tmax_x:.6f}s, Tmax_y={c.Tmax_y:.6f}s\n")
                    f.write(f"Tx_modal={c.Tx_modal:.5f}s (UX max), Ty_modal={c.Ty_modal:.5f}s (UY max)\n")
                    f.write(f"T_aju_x={c.T_aju_x:.5f}s, T_aju_y={c.T_aju_y:.5f}s\n\n")
                    f.write("II. Sa (numpy.interp)\n" + "-"*68 + "\n")
                    f.write(f"Sa_x={c.Sa_x:.6f}g @ {c.T_aju_x:.5f}s\nSa_y={c.Sa_y:.6f}g @ {c.T_aju_y:.5f}s\n")
                    f.write(f"Curva: {len(c.periods)} puntos T[{c.periods.min():.3f}-{c.periods.max():.3f}]s\n\n")
                    f.write("III. Vs = M*Sa*g/1000\n" + "-"*68 + "\n")
                    f.write(f"Masa {c.masa_total_kg:,.2f} kg ({c.masa_total_ton:.4f} Ton)\n")
                    f.write(f"VsX={c.VsX:,.4f} kN, VsY={c.VsY:,.4f} kN\n\n")
                    f.write("IV. Vt (BaseReact)\n" + "-"*68 + "\n")
                    f.write(f"VtX ({c.caso_dx})={c.VtX:,.4f} kN\nVtY ({c.caso_dy})={c.VtY:,.4f} kN\n\n")
                    f.write("V. CRITERIO\n" + "-"*68 + "\n")
                    f.write(f"req_X={c.req_X:,.4f} kN  →  VtX < req_X ? {'SÍ' if c.necesita_x else 'NO'}\n")
                    f.write(f"req_Y={c.req_Y:,.4f} kN  →  VtY < req_Y ? {'SÍ' if c.necesita_y else 'NO'}\n\n")
                    f.write("VI. FACTORES DE ESCALA (ingresar en ETABS)\n" + "-"*68 + "\n")
                    fx = f"{c.factor_x:.4f}" if not isinstance(c.factor_x, str) else "No Aplica (1.0)"
                    fy = f"{c.factor_y:.4f}" if not isinstance(c.factor_y, str) else "No Aplica (1.0)"
                    f.write(f"Factor X ({c.caso_dx}): {fx}\n")
                    f.write(f"Factor Y ({c.caso_dy}): {fy}\n")
                    if not isinstance(c.factor_x, str): f.write(f"  =({c.porcentaje:.2f}*{c.VsX:.2f})/{c.VtX:.2f}\n")
                    else: f.write("  X no requiere ajuste → deje Scale Factor = 1.0\n")
                    if not isinstance(c.factor_y, str): f.write(f"  =({c.porcentaje:.2f}*{c.VsY:.2f})/{c.VtY:.2f}\n")
                    else: f.write("  Y no requiere ajuste → deje Scale Factor = 1.0\n")
                    f.write("\nUnidades: kN-m-C (6), g=9.80665, Sa interp numpy.interp, T_aju=min(Ta*Cu,Tmodal)\n")
                    f.write("="*68 + "\n")
                self.log(f"✓ Memoria TXT guardada: {path}"); messagebox.showinfo("Exportado", f"Guardado en:\n{path}")
            except Exception as e:
                self.log(f"❌ TXT error: {e}"); messagebox.showerror("Error", str(e))

        def export_excel(self):
            if self.mass_res is None or self.modal_res is None or self.calculo is None:
                messagebox.showwarning("Faltan datos", "Extraiga y calcule antes de exportar."); return
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"Tablas_ADEE_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", filetypes=[("Excel","*.xlsx")])
            if not path: return
            try:
                import openpyxl  # noqa: F401
            except ImportError:
                messagebox.showerror("Falta openpyxl", "Instale openpyxl:\n\npip install openpyxl"); return
            try:
                with pd.ExcelWriter(path, engine="openpyxl") as w:
                    df_m = self.mass_res.df.copy()
                    # Fila TOTAL
                    total_row = {col: "" for col in df_m.columns}
                    total_row[df_m.columns[0]] = "TOTAL"
                    total_row[self.mass_res.col_masa] = self.mass_res.masa_total_ton
                    df_m.loc[len(df_m)] = total_row  # type: ignore
                    df_m.to_excel(w, sheet_name="Masa_Summary", index=False)
                    self.modal_res.df.to_excel(w, sheet_name="Modal_Masas", index=False)
                    c = self.calculo
                    dfc = pd.DataFrame([
                        ["Parámetro","Valor","Unidad","Nota"],
                        ["Ta_x",c.Ta_x,"s","Aprox X"],["Ta_y",c.Ta_y,"s","Aprox Y"],["Cu",c.Cu,"","A.4.2.2"],
                        ["Tmax_x",c.Tmax_x,"s","Ta*Cu"],["Tmax_y",c.Tmax_y,"s","Ta*Cu"],
                        ["Tx_modal",c.Tx_modal,"s",f"Modo {self.modal_res.idx_max_ux+1} max UX"],["Ty_modal",c.Ty_modal,"s",f"Modo {self.modal_res.idx_max_uy+1} max UY"],
                        ["T_aju_x",c.T_aju_x,"s","min(Tmax,Tx)"],["T_aju_y",c.T_aju_y,"s","min(Tmax,Ty)"],
                        ["Sa_x",c.Sa_x,"g",f"@ {c.T_aju_x:.4f}s"],["Sa_y",c.Sa_y,"g",f"@ {c.T_aju_y:.4f}s"],
                        ["Masa",c.masa_total_kg,"kg",f"{c.masa_total_ton:.2f} Ton"],
                        ["VsX",c.VsX,"kN","M*Sa*g/1000"],["VsY",c.VsY,"kN","M*Sa*g/1000"],
                        ["VtX",c.VtX,"kN",c.caso_dx],["VtY",c.VtY,"kN",c.caso_dy],
                        ["Porcentaje",c.porcentaje,"",c.tipo],["req_X",c.req_X,"kN","porc*VsX"],["req_Y",c.req_Y,"kN","porc*VsY"],
                        ["Necesita X","SÍ" if c.necesita_x else "NO","","Vt<req?"],["Necesita Y","SÍ" if c.necesita_y else "NO","","Vt<req?"],
                        ["Factor X",c.factor_x if not isinstance(c.factor_x,str) else "No Aplica (1.0)","","ETABS"],["Factor Y",c.factor_y if not isinstance(c.factor_y,str) else "No Aplica (1.0)","","ETABS"],
                    ])
                    dfc.to_excel(w, sheet_name="Calculo_ADEE", index=False, header=False)
                    pd.DataFrame({"Period [s]":c.periods, "Sa [g]":c.sa}).to_excel(w, sheet_name="Curva_Espectro", index=False)
                    if self.base_res is not None: self.base_res.df.to_excel(w, sheet_name="BaseReact", index=False)
                self.log(f"✓ Excel guardado: {path}"); messagebox.showinfo("Exportado", f"Guardado en:\n{path}")
            except PermissionError:
                messagebox.showerror("Permiso denegado", f"No se pudo escribir:\n{path}\n\nCierre el archivo si está abierto.")
            except Exception as e:
                self.log(f"❌ Excel error: {e}\n{traceback.format_exc()}"); messagebox.showerror("Error", f"{e}\n\n{traceback.format_exc()}")

        def copy_factors(self):
            if self.calculo is None: return
            c = self.calculo
            fx = f"{c.factor_x:.4f}" if not isinstance(c.factor_x,str) else "1.0 (No Aplica)"
            fy = f"{c.factor_y:.4f}" if not isinstance(c.factor_y,str) else "1.0 (No Aplica)"
            txt = f"Factor X ({c.caso_dx}): {fx}\nFactor Y ({c.caso_dy}): {fy}"
            self.clipboard_clear(); self.clipboard_append(txt); self.update()
            self.log(f"📋 Copiado: {txt.replace(chr(10),' | ')}"); messagebox.showinfo("Copiado", txt)

# ---------------------------------------------------------------------------
# 5) Fallback consola (si no hay GUI)
# ---------------------------------------------------------------------------
def modo_consola():
    print("="*68)
    print(" MODO CONSOLA — Ajuste Cortante Basal NSR-10 (sin GUI)")
    print("="*68)
    if MISSING_DEPS:
        print(f"Faltan dependencias: {', '.join(MISSING_DEPS)}")
        print("Instale: pip install pandas numpy comtypes openpyxl")
        return
    # Usar mock si --mock
    use_mock = "--mock" in sys.argv
    if use_mock:
        SapModel = MockSapModel(); print("→ Usando MockSapModel (pruebas sin ETABS)")
    else:
        try:
            SapModel, _ = connect_to_etabs()
            SapModel.SetPresentUnits(6)
            print("✅ Conectado a ETABS")
        except Exception as e:
            print(f"❌ {e}")
            print("Tip: ejecute con --mock para probar sin ETABS")
            return
    # Inputs
    try:
        reg = input("¿Edificio Regular? (S/N) [S]: ").strip().upper() or "S"
        porcentaje = 0.80 if reg == "S" else 0.90
        tipo = "Regular" if reg == "S" else "Irregular"
        func = input("Función espectro [Espectro_NSR-10]: ").strip() or "Espectro_NSR-10"
        Ta_x = float(input("Ta X [s] (ej. 0.746): ") or "0.746")
        Ta_y = float(input("Ta Y [s] (ej. 0.746): ") or "0.746")
        Cu = float(input("Cu [ej. 1.2]: ") or "1.2")
        caso_dx = input("Caso dinámico X [Fsx_ADE]: ").strip() or "Fsx_ADE"
        caso_dy = input("Caso dinámico Y [Fsy_ADE]: ").strip() or "Fsy_ADE"
        caso_m = input("Caso Modal [Modal]: ").strip() or "Modal"
    except Exception as e:
        print(f"Entrada inválida: {e}"); return

    try:
        print("\nExtrayendo masas...")
        mass = extract_mass_summary(SapModel)
        print(f"Masa: {mass.masa_total_kg:,.2f} kg ({mass.masa_total_ton:.4f} Ton)")
        print(mass.df.to_string(index=False))

        print(f"\nExtrayendo modal '{caso_m}'...")
        modal = extract_modal_data(SapModel, caso_m)
        print(modal.df.to_string(index=False))
        print(f"Tx={modal.tx_modal:.4f}s (UX {modal.ux[modal.idx_max_ux]:.3f}) Ty={modal.ty_modal:.4f}s (UY {modal.uy[modal.idx_max_uy]:.3f})")

        print(f"\nCurva '{func}'...")
        p, sa = extract_spectrum_curve(SapModel, func)
        print(f"  {len(p)} puntos T[{p.min():.3f}-{p.max():.3f}]s")

        Tmax_x = Ta_x*Cu; Tmax_y = Ta_y*Cu
        Taju_x = min(Tmax_x, modal.tx_modal); Taju_y = min(Tmax_y, modal.ty_modal)
        Sa_x = float(np.interp(Taju_x, p, sa)); Sa_y = float(np.interp(Taju_y, p, sa))
        print(f"Sa_x={Sa_x:.6f}g @ {Taju_x:.4f}s, Sa_y={Sa_y:.6f}g @ {Taju_y:.4f}s")

        print(f"\nBaseReact {caso_dx}/{caso_dy}...")
        base = extract_base_reactions(SapModel, caso_dx, caso_dy)
        print(base.df.to_string(index=False))
        print(f"VtX={base.VtX:.2f} kN VtY={base.VtY:.2f} kN")

        calc = calcular_ajuste(Ta_x, Ta_y, Cu, porcentaje, tipo, mass.masa_total_kg, modal.tx_modal, modal.ty_modal, Sa_x, Sa_y, base.VtX, base.VtY, p, sa, func, caso_dx, caso_dy)
        print("\n" + "="*68)
        print(f" VsX={calc.VsX:,.2f} kN VsY={calc.VsY:,.2f} kN")
        print(f" req_X={calc.req_X:,.2f} req_Y={calc.req_Y:,.2f}")
        print(f" ¿VtX<req? {calc.VtX:.2f}<{calc.req_X:.2f} → {'SÍ' if calc.necesita_x else 'NO'}  Factor X={calc.factor_x}")
        print(f" ¿VtY<req? {calc.VtY:.2f}<{calc.req_Y:.2f} → {'SÍ' if calc.necesita_y else 'NO'}  Factor Y={calc.factor_y}")
        print("="*68)
    except Exception as e:
        print(f"❌ Error: {e}\n{traceback.format_exc()}")

# ---------------------------------------------------------------------------
# 6) Main — decide GUI vs Consola, maneja errores de PyQt/Tkinter
# ---------------------------------------------------------------------------
def main():
    # Permitir --mock desde cualquier modo
    start_mock = "--mock" in sys.argv or "-m" in sys.argv

    # Si no hay Tkinter ni se puede iniciar GUI → consola
    if not TK_AVAILABLE:
        print("Tkinter no disponible. Intentando modo consola...")
        modo_consola()
        return

    # Intentar iniciar Tkinter; capturar error de display (p.ej. sin X11)
    try:
        # En Windows, evitar escalado borroso
        if sys.platform == "win32":
            try: from ctypes import windll; windll.shcore.SetProcessDpiAwareness(1)  # type: ignore
            except Exception: pass
        app = App(start_mock=start_mock)
        # Centrar ventana
        try:
            app.update_idletasks()
            w = app.winfo_width(); h = app.winfo_height()
            ws = app.winfo_screenwidth(); hs = app.winfo_screenheight()
            x = (ws//2) - (w//2); y = (hs//2) - (h//2) - 20
            app.geometry(f"+{x}+{y}")
        except Exception:
            pass
        app.mainloop()
    except tk.TclError as e:
        # Error típico: no display (Linux sin X) o Qt xcb previo
        print(f"Error iniciando GUI Tkinter: {e}")
        print("Cayendo a modo consola...")
        traceback.print_exc()
        modo_consola()
    except Exception as e:
        print(f"Error fatal GUI: {e}\n{traceback.format_exc()}")
        try: messagebox.showerror("Error fatal", f"{e}\n\n{traceback.format_exc()}")
        except Exception: pass
        modo_consola()

if __name__ == "__main__":
    main()
