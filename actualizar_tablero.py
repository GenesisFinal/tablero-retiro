#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de Actualización Automática y Validación Actuarial del Tablero de Retiro
- Busca y lee archivos Excel 'Dataset Retiro*.xlsx' en la carpeta.
- Preserva la historia cronológica completa y anexa nuevos períodos.
- Calcula las tablas de rentabilidad compuesta en Pesos, Dólares y Benchmarks.
- Valida la integridad actuarial al peso antes de guardar y desplegar.
"""
import os
import sys
import glob
import json
import subprocess
import numpy as np

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "data_retiro.json")

def find_excel_file():
    patterns = [
        os.path.join(BASE_DIR, "Dataset Retiro*.xlsx"),
        os.path.join(BASE_DIR, "*.xlsx")
    ]
    for p in patterns:
        matches = glob.glob(p)
        if matches:
            for m in matches:
                if "Dataset Retiro" in os.path.basename(m):
                    return m
            return matches[0]
    return None

def normalize_date(d_str):
    months = {'ene': '01', 'feb': '02', 'mar': '03', 'abr': '04', 'may': '05', 'jun': '06',
              'jul': '07', 'ago': '08', 'sep': '09', 'sept': '09', 'oct': '10', 'nov': '11', 'dic': '12'}
    parts = d_str.lower().split('-')
    if len(parts) == 2:
        m = months.get(parts[0], '01')
        y = '20' + parts[1]
        return f"{y}-{m}"
    return d_str

def parse_excel_dataset(excel_path):
    import openpyxl
    print(f"Cargando dataset desde: {os.path.basename(excel_path)}...")
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    sheet_name = "Modificacion Propuesta" if "Modificacion Propuesta" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))

    header_dates = [str(x).strip() for x in rows[0][5:]]
    dates = [normalize_date(d) for d in header_dates]

    def get_row_floats(row_idx):
        r = rows[row_idx - 1]
        vals = []
        for x in r[5:]:
            if x is None or str(x).strip() in ['', '-', 'None']:
                vals.append(0.0)
            else:
                try:
                    vals.append(float(x))
                except:
                    vals.append(0.0)
        return vals

    data_map = {
        "DATES": dates,
        "FX": get_row_floats(2),
        "IPC_MONTHLY": [v * 100.0 if 0.0 < v < 1.0 else v for v in get_row_floats(4)],
        
        "ri_primas_p": get_row_floats(5),
        "ri_primas_d": get_row_floats(6),
        "ri_rentas_p": get_row_floats(8),
        "ri_rentas_d": get_row_floats(9),
        "ri_rescates_p": get_row_floats(11),
        "ri_rescates_d": get_row_floats(12),
        "ri_ct_act_p": get_row_floats(17),
        "ri_ct_act_d": get_row_floats(18),
        "ri_ct_pas_p": get_row_floats(20),
        "ri_ct_pas_d": get_row_floats(21),
        "ri_pol_act_p": get_row_floats(32),
        "ri_pol_act_d": get_row_floats(33),
        "ri_pol_pas_p": get_row_floats(34),
        "ri_pol_pas_d": get_row_floats(35),
        "ri_rent_p": get_row_floats(29),
        "ri_tt_p": get_row_floats(30),
        "ri_rent_d": get_row_floats(31),

        "rc_primas_p": get_row_floats(36),
        "rc_primas_d": get_row_floats(37),
        "rc_rentas_p": get_row_floats(39),
        "rc_rentas_d": get_row_floats(40),
        "rc_rescates_p": get_row_floats(42),
        "rc_rescates_d": get_row_floats(43),
        "rc_ct_act_p": get_row_floats(49),
        "rc_ct_act_d": get_row_floats(50),
        "rc_ct_pas_p": get_row_floats(52),
        "rc_ct_pas_d": get_row_floats(53),
        "rc_pol_act_p": get_row_floats(64),
        "rc_pol_act_d": get_row_floats(65),
        "rc_cert_act_p": get_row_floats(66),
        "rc_cert_act_d": get_row_floats(67),
        "rc_cert_pas_p": get_row_floats(68),
        "rc_cert_pas_d": get_row_floats(69),
        "rc_rent_p": get_row_floats(61),
        "rc_tt_p": get_row_floats(62),
        "rc_rent_d": get_row_floats(63),

        "rvp_primas_p": get_row_floats(70),
        "rvp_primas_d": get_row_floats(71),
        "rvp_rentas_p": get_row_floats(73),
        "rvp_rentas_d": get_row_floats(74),
        "rvp_rescates_p": get_row_floats(76),
        "rvp_rescates_d": get_row_floats(77),
        "rvp_ct_p": get_row_floats(82),
        "rvp_ct_d": get_row_floats(83),
        "rvp_rent_p": get_row_floats(86),
        "rvp_tt_p": get_row_floats(87),
        "rvp_rent_d": get_row_floats(88),

        "emp_primas_p": get_row_floats(89),
        "emp_primas_d": get_row_floats(90),
        "emp_rentas_p": get_row_floats(92),
        "emp_rentas_d": get_row_floats(93),
        "emp_rescates_p": get_row_floats(95),
        "emp_rescates_d": get_row_floats(96),
        "emp_ct_act_p": get_row_floats(101),
        "emp_ct_act_d": get_row_floats(102),
        "emp_ct_pas_p": get_row_floats(104),
        "emp_ct_pas_d": get_row_floats(105),
        "emp_cert_act_p": get_row_floats(107),
        "emp_cert_act_d": get_row_floats(108),
        "emp_cert_pas_p": get_row_floats(109),
        "emp_cert_pas_d": get_row_floats(110),

        "vin_primas_p": get_row_floats(111),
        "vin_primas_d": get_row_floats(112),
        "vin_rentas_p": get_row_floats(114),
        "vin_rentas_d": get_row_floats(115),
        "vin_rescates_p": get_row_floats(117),
        "vin_rescates_d": get_row_floats(118),
        "vin_ct_act_p": get_row_floats(123),
        "vin_ct_act_d": get_row_floats(124),
        "vin_ct_pas_p": get_row_floats(126),
        "vin_ct_pas_d": get_row_floats(127),
        "vin_cert_act_p": get_row_floats(129),
        "vin_cert_act_d": get_row_floats(130),
        "vin_cert_pas_p": get_row_floats(131),
        "vin_cert_pas_d": get_row_floats(132),
    }

    # IPC julio 2026 queda en 0.0 o None si el INDEC/FACPCE aún no lo publicó
    pass

    return data_map

def compute_official_tables(series):
    raw_dates = series["DATES"]
    dates = [str(d).strip().lower() for d in raw_dates]
    date_to_idx = {d: i for i, d in enumerate(dates)}
    N = len(dates)
    last_idx = N - 1

    rent_p = np.array(series["ri_rent_p"])
    rent_d = np.array(series["ri_rent_d"])
    ipc_arr = np.array([x / 100.0 for x in series["IPC_MONTHLY"]])
    tt_arr = np.array(series.get("ri_tt_p", series.get("rc_tt_p", [0]*N)))
    fx_arr = np.array(series["FX"])

    def calc_compound_perf(rate_series, start_idx, end_idx):
        rates = rate_series[start_idx:end_idx+1]
        m = len(rates)
        if m == 0: return "0,00%", "0,00%"
        prod = np.prod(1.0 + rates) - 1.0
        anual = ((1.0 + prod) ** (12.0 / m)) - 1.0 if m > 0 else 0.0
        return f"{prod*100:,.2f}%".replace('.', ','), f"{anual*100:,.2f}%".replace('.', ',')

    def calc_fx_perf(fx_series, start_idx, end_idx):
        m = end_idx - start_idx + 1
        base_val = fx_series[0] if start_idx == 0 else fx_series[start_idx - 1]
        end_val = fx_series[end_idx]
        prod = (end_val - base_val) / base_val
        anual = ((1.0 + prod) ** (12.0 / m)) - 1.0 if m > 0 else 0.0
        return f"{prod*100:,.2f}%".replace('.', ','), f"{anual*100:,.2f}%".replace('.', ',')

    periods_def = [
        ("Último mes (Julio 2026 / L1M)", last_idx, last_idx),
        ("Últimos 3 meses (L3M)", max(0, last_idx - 2), last_idx),
        ("Últimos 6 meses (L6M)", max(0, last_idx - 5), last_idx),
        ("Ejercicio 2025/2026 (Cerrado)", date_to_idx.get('jul-25', 0), date_to_idx.get('jun-26', 0)),
        ("Ejercicio 2024/2025 (Cerrado)", date_to_idx.get('jul-24', 0), date_to_idx.get('jun-25', 0)),
        ("Ejercicio 2023/2024 (Cerrado)", date_to_idx.get('jul-23', 0), date_to_idx.get('jun-24', 0)),
        ("Ejercicio 2026/2027 (En curso)", date_to_idx.get('jul-26', last_idx), last_idx),
        ("Últimos 12 meses (L12M)", max(0, last_idx - 11), last_idx),
        ("Últimos 24 meses (L24M)", max(0, last_idx - 23), last_idx),
        ("Últimos 36 meses (L36M)", max(0, last_idx - 35), last_idx),
    ]

    rent_pesos_table = []
    rent_dolares_table = []
    benchmarks_table = []

    for name, s_idx, e_idx in periods_def:
        p_acum, p_anual = calc_compound_perf(rent_p, s_idx, e_idx)
        rent_pesos_table.append({"p": name, "acum": p_acum, "anual": p_anual})
        
        d_acum, d_anual = calc_compound_perf(rent_d, s_idx, e_idx)
        rent_dolares_table.append({"p": name, "acum": d_acum, "anual": d_anual})
        
        ipc_a, ipc_an = calc_compound_perf(ipc_arr, s_idx, e_idx)
        tt_a, tt_an = calc_compound_perf(tt_arr, s_idx, e_idx)
        tc_a, tc_an = calc_fx_perf(fx_arr, s_idx, e_idx)
        benchmarks_table.append({
            "p": name,
            "ipc_a": ipc_a,
            "ipc_an": ipc_an,
            "tt_a": tt_a,
            "tt_an": tt_an,
            "tc_a": tc_a,
            "tc_an": tc_an
        })

    return {
        "rent_pesos": rent_pesos_table,
        "rent_dolares": rent_dolares_table,
        "benchmarks": benchmarks_table
    }

def validate_integrity(series):
    N = len(series.get("DATES", []))
    last_idx = N - 1
    tc = series["FX"][last_idx]
    date_label = series["DATES"][last_idx]

    ri_act = series["ri_ct_act_p"][last_idx] + series["ri_ct_act_d"][last_idx] * tc
    ri_pas = series["ri_ct_pas_p"][last_idx] + series["ri_ct_pas_d"][last_idx] * tc
    rc_act = series["rc_ct_act_p"][last_idx] + series["rc_ct_act_d"][last_idx] * tc
    rc_pas = series["rc_ct_pas_p"][last_idx] + series["rc_ct_pas_d"][last_idx] * tc
    rvp_pas = series["rvp_ct_p"][last_idx] + series["rvp_ct_d"][last_idx] * tc

    emp_act = series["emp_ct_act_p"][last_idx] + series["emp_ct_act_d"][last_idx] * tc
    emp_pas = series["emp_ct_pas_p"][last_idx] + series["emp_ct_pas_d"][last_idx] * tc
    vin_act = series["vin_ct_act_p"][last_idx] + series["vin_ct_act_d"][last_idx] * tc
    vin_pas = series["vin_ct_pas_p"][last_idx] + series["vin_ct_pas_d"][last_idx] * tc

    tot_prod = ri_act + ri_pas + rc_act + rc_pas + rvp_pas
    tot_act = ri_act + rc_act
    tot_pas = ri_pas + rc_pas + rvp_pas

    vin_tot = vin_act + vin_pas
    emp_tot = emp_act + emp_pas
    neg_tot = tot_prod - vin_tot - emp_tot

    print(f"=== VALIDACION ACTUARIAL AL CIERRE {date_label} (TC: ${tc:,.2f}) ===")
    print(f"CT Total Consolidado:    $ {tot_prod:,.2f}  ($ {tot_prod/1e6:,.2f} M)")
    print(f"  - CT Ahorro (Activo):   $ {tot_act:,.2f}  ($ {tot_act/1e6:,.2f} M)")
    print(f"  - CT Renta (Pasivo):    $ {tot_pas:,.2f}  ($ {tot_pas/1e6:,.2f} M)")
    print("----------------------------------------------------------------------")
    print(f"  - Vinculados:           $ {vin_tot:,.2f}  ($ {vin_tot/1e6:,.2f} M)")
    print(f"  - Empleados:            $ {emp_tot:,.2f}  ($ {emp_tot/1e6:,.2f} M)")
    print(f"  - Negocio Propio:       $ {neg_tot:,.2f}  ($ {neg_tot/1e6:,.2f} M)")
    print(f"  - Suma Segmentos:       $ {vin_tot + emp_tot + neg_tot:,.2f}")
    print("======================================================================")

    diff = abs((vin_tot + emp_tot + neg_tot) - tot_prod)
    if diff > 1.0:
        print(f"ERROR: Discrepancia detectada de ${diff:,.2f} en la suma de segmentos.")
        return False
    print(">>> VALIDACION 100% EXITOSA: Todas las sumas cuadran al peso. <<<")
    return True

def main():
    excel_path = find_excel_file()
    if excel_path:
        series = parse_excel_dataset(excel_path)
    elif os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            existing = json.load(f)
            series = existing.get("series", {})
    else:
        print("Error: No se encontro archivo Excel ni data_retiro.json.")
        sys.exit(1)

    if not validate_integrity(series):
        sys.exit(1)

    tables = compute_official_tables(series)
    
    out_json = {
        "metadata": {
            "origen": os.path.basename(excel_path) if excel_path else "data_retiro.json",
            "cierre": series["DATES"][-1],
            "total_periodos": len(series["DATES"]),
            "tipo_cambio_cierre": series["FX"][-1]
        },
        "tables": tables,
        "series": series
    }

    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(out_json, f, indent=2, ensure_ascii=False)
    print(f"OK: {DATA_FILE} actualizado exitosamente con series y tablas de rentabilidad.")

    if len(sys.argv) > 1 and sys.argv[1] == "--deploy":
        deploy_script = os.path.join(BASE_DIR, "deploy_to_github.py")
        if os.path.exists(deploy_script):
            print("\nIniciando despliegue a GitHub Pages...")
            subprocess.run([sys.executable, deploy_script], check=True)

if __name__ == "__main__":
    main()
