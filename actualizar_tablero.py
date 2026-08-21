import os, sys, json, re
import openpyxl
import numpy as np

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATASET_PATH = os.path.join(BASE_DIR, "Nuevo Dataset Retiro - Jul 26.xlsx")
OUTPUT_JSON_PATH = os.path.join(BASE_DIR, "data_retiro.json")

# Macro fallback series for earlier months (2023-03 to 2023-05) and complete 41 months
# FX BNA Divisa (Mar-23 to Jul-26, 41 periods)
MACRO_FX_FALLBACK = [
    209.01, 222.68, 239.85, 256.70, 275.25, 350.00, 349.95, 350.00, 360.50, 808.45, 
    826.40, 842.20, 858.00, 876.50, 895.50, 912.00, 932.00, 953.50, 970.50, 992.00, 
    1011.50, 1032.00, 1053.50, 1064.75, 1074.00, 1170.00, 1188.00, 1205.00, 1374.00, 1342.00, 
    1380.00, 1445.00, 1451.50, 1455.00, 1447.00, 1397.00, 1382.00, 1391.00, 1408.00, 1482.00, 
    1485.00
]

# IPC mensual INDEC/FACPCE (%) (Mar-23 to Jul-26, 41 periods)
MACRO_IPC_FALLBACK = [
    7.68, 8.41, 7.78, 5.95, 6.34, 12.44, 12.75, 8.30, 12.81, 25.47, 
    20.61, 13.24, 11.01, 8.83, 4.18, 4.58, 4.03, 4.17, 3.47, 2.69, 
    2.43, 2.70, 2.21, 2.40, 3.73, 2.78, 1.50, 1.62, 1.90, 1.88, 
    2.08, 2.34, 2.47, 2.85, 2.88, 2.90, 3.38, 2.58, 2.15, 1.89, 
    2.11
]

def load_and_parse_multidimensional_dataset(file_path):
    print(f"Cargando dataset multidimensional desde: {file_path}...")
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb["DataSet Andres"]

    # 1. Extraer fechas de la fila 1 (columnas 5 a 45)
    raw_dates = [ws.cell(row=1, column=c).value for c in range(5, 46)]
    dates = []
    for d in raw_dates:
        if d is not None:
            if hasattr(d, 'strftime'):
                dates.append(d.strftime('%Y-%m'))
            else:
                dates.append(str(d).strip())
    N = len(dates)
    print(f"Períodos detectados: {N} ({dates[0]} a {dates[-1]})")

    def clean_arr(row_idx):
        arr = []
        for c in range(5, 5 + N):
            v = ws.cell(row=row_idx, column=c).value
            if v is None:
                arr.append(0.0)
            else:
                try:
                    arr.append(float(v))
                except:
                    arr.append(0.0)
        return arr

    def clean_rate_arr(row_idx):
        arr = []
        for c in range(5, 5 + N):
            v = ws.cell(row=row_idx, column=c).value
            if v is None:
                arr.append(0.0)
            else:
                try:
                    arr.append(float(v))
                except:
                    arr.append(0.0)
        return arr

    # 2. Mapear todas las series del dataset multidimensional
    cube = {
        "individual": {
            "negocio": {
                "pesos": {
                    "primas": clean_arr(3), "rentas": clean_arr(11), "rescates": clean_arr(19),
                    "pase_pasividad": clean_arr(27), "ct_activos": clean_arr(35), "ct_pasivos": clean_arr(43),
                    "polizas_act": clean_arr(62), "polizas_pas": clean_arr(70), "cert_act": clean_arr(62), "cert_pas": clean_arr(70)
                },
                "dolares": {
                    "primas": clean_arr(6), "rentas": clean_arr(14), "rescates": clean_arr(22),
                    "pase_pasividad": clean_arr(30), "ct_activos": clean_arr(38), "ct_pasivos": clean_arr(46),
                    "polizas_act": clean_arr(65), "polizas_pas": clean_arr(73), "cert_act": clean_arr(65), "cert_pas": clean_arr(73)
                }
            },
            "empleados": {
                "pesos": {
                    "primas": clean_arr(4), "rentas": clean_arr(12), "rescates": clean_arr(20),
                    "pase_pasividad": clean_arr(28), "ct_activos": clean_arr(36), "ct_pasivos": clean_arr(44),
                    "polizas_act": clean_arr(63), "polizas_pas": clean_arr(71), "cert_act": clean_arr(63), "cert_pas": clean_arr(71)
                },
                "dolares": {
                    "primas": clean_arr(7), "rentas": clean_arr(15), "rescates": clean_arr(23),
                    "pase_pasividad": clean_arr(31), "ct_activos": clean_arr(39), "ct_pasivos": clean_arr(47),
                    "polizas_act": clean_arr(66), "polizas_pas": clean_arr(74), "cert_act": clean_arr(66), "cert_pas": clean_arr(74)
                }
            },
            "vinculados": {
                "pesos": {
                    "primas": clean_arr(5), "rentas": clean_arr(13), "rescates": clean_arr(21),
                    "pase_pasividad": clean_arr(29), "ct_activos": clean_arr(37), "ct_pasivos": clean_arr(45),
                    "polizas_act": clean_arr(64), "polizas_pas": clean_arr(72), "cert_act": clean_arr(64), "cert_pas": clean_arr(72)
                },
                "dolares": {
                    "primas": clean_arr(8), "rentas": clean_arr(16), "rescates": clean_arr(24),
                    "pase_pasividad": clean_arr(32), "ct_activos": clean_arr(40), "ct_pasivos": clean_arr(48),
                    "polizas_act": clean_arr(67), "polizas_pas": clean_arr(75), "cert_act": clean_arr(67), "cert_pas": clean_arr(75)
                }
            },
            "rentabilidad": {
                "pesos": clean_rate_arr(57),
                "tt_pesos": clean_rate_arr(58),
                "dolares": clean_rate_arr(59)
            }
        },
        "colectivo": {
            "negocio": {
                "pesos": {
                    "primas": clean_arr(78), "rentas": clean_arr(86), "rescates": clean_arr(94),
                    "pase_pasividad": clean_arr(102), "pase_pasividad_ind": clean_arr(110),
                    "ct_activos": clean_arr(118), "ct_pasivos": clean_arr(126),
                    "polizas": clean_arr(146), "cert_act": clean_arr(154), "cert_pas": clean_arr(162)
                },
                "dolares": {
                    "primas": clean_arr(81), "rentas": clean_arr(89), "rescates": clean_arr(97),
                    "pase_pasividad": clean_arr(105), "pase_pasividad_ind": clean_arr(113),
                    "ct_activos": clean_arr(121), "ct_pasivos": clean_arr(129),
                    "polizas": clean_arr(149), "cert_act": clean_arr(157), "cert_pas": clean_arr(165)
                }
            },
            "empleados": {
                "pesos": {
                    "primas": clean_arr(79), "rentas": clean_arr(87), "rescates": clean_arr(95),
                    "pase_pasividad": clean_arr(103), "pase_pasividad_ind": clean_arr(111),
                    "ct_activos": clean_arr(119), "ct_pasivos": clean_arr(127),
                    "polizas": clean_arr(147), "cert_act": clean_arr(155), "cert_pas": clean_arr(163)
                },
                "dolares": {
                    "primas": clean_arr(82), "rentas": clean_arr(90), "rescates": clean_arr(98),
                    "pase_pasividad": clean_arr(106), "pase_pasividad_ind": clean_arr(114),
                    "ct_activos": clean_arr(122), "ct_pasivos": clean_arr(130),
                    "polizas": clean_arr(150), "cert_act": clean_arr(158), "cert_pas": clean_arr(166)
                }
            },
            "vinculados": {
                "pesos": {
                    "primas": clean_arr(80), "rentas": clean_arr(88), "rescates": clean_arr(96),
                    "pase_pasividad": clean_arr(104), "pase_pasividad_ind": clean_arr(112),
                    "ct_activos": clean_arr(120), "ct_pasivos": clean_arr(128),
                    "polizas": clean_arr(148), "cert_act": clean_arr(156), "cert_pas": clean_arr(164)
                },
                "dolares": {
                    "primas": clean_arr(83), "rentas": clean_arr(91), "rescates": clean_arr(99),
                    "pase_pasividad": clean_arr(107), "pase_pasividad_ind": clean_arr(115),
                    "ct_activos": clean_arr(123), "ct_pasivos": clean_arr(131),
                    "polizas": clean_arr(151), "cert_act": clean_arr(159), "cert_pas": clean_arr(167)
                }
            },
            "rentabilidad": {
                "pesos": clean_rate_arr(140),
                "tt_pesos": clean_rate_arr(141),
                "dolares": clean_rate_arr(142)
            }
        },
        "rvp_art": {
            "negocio": {
                "pesos": {
                    "primas": clean_arr(170), "rentas": clean_arr(174), "rescates": clean_arr(178),
                    "pase_pasividad": clean_arr(182), "ct_activos": [0.0]*N, "ct_pasivos": clean_arr(186),
                    "polizas": [0.0]*N, "cert_act": [0.0]*N, "cert_pas": [0.0]*N
                },
                "dolares": {
                    "primas": clean_arr(171), "rentas": clean_arr(175), "rescates": clean_arr(179),
                    "pase_pasividad": clean_arr(183), "ct_activos": [0.0]*N, "ct_pasivos": clean_arr(187),
                    "polizas": [0.0]*N, "cert_act": [0.0]*N, "cert_pas": [0.0]*N
                }
            },
            "empleados": {
                "pesos": { "primas": [0.0]*N, "rentas": [0.0]*N, "rescates": [0.0]*N, "pase_pasividad": [0.0]*N, "ct_activos": [0.0]*N, "ct_pasivos": [0.0]*N, "polizas": [0.0]*N, "cert_act": [0.0]*N, "cert_pas": [0.0]*N },
                "dolares": { "primas": [0.0]*N, "rentas": [0.0]*N, "rescates": [0.0]*N, "pase_pasividad": [0.0]*N, "ct_activos": [0.0]*N, "ct_pasivos": [0.0]*N, "polizas": [0.0]*N, "cert_act": [0.0]*N, "cert_pas": [0.0]*N }
            },
            "vinculados": {
                "pesos": { "primas": [0.0]*N, "rentas": [0.0]*N, "rescates": [0.0]*N, "pase_pasividad": [0.0]*N, "ct_activos": [0.0]*N, "ct_pasivos": [0.0]*N, "polizas": [0.0]*N, "cert_act": [0.0]*N, "cert_pas": [0.0]*N },
                "dolares": { "primas": [0.0]*N, "rentas": [0.0]*N, "rescates": [0.0]*N, "pase_pasividad": [0.0]*N, "ct_activos": [0.0]*N, "ct_pasivos": [0.0]*N, "polizas": [0.0]*N, "cert_act": [0.0]*N, "cert_pas": [0.0]*N }
            },
            "rentabilidad": {
                "pesos": clean_rate_arr(192),
                "tt_pesos": clean_rate_arr(193),
                "dolares": clean_rate_arr(194)
            }
        }
    }

    # Macro data
    fx_series = MACRO_FX_FALLBACK[-N:]
    ipc_monthly = MACRO_IPC_FALLBACK[-N:]

    return {
        "DATES": dates,
        "FX": fx_series,
        "IPC_MONTHLY": ipc_monthly,
        "cube": cube
    }

def compute_official_tables(dates, cube, fx_series, ipc_monthly):
    N = len(dates)
    last_idx = N - 1
    date_to_idx = {d: i for i, d in enumerate(dates)}

    rent_p = np.array(cube["individual"]["rentabilidad"]["pesos"])
    rent_d = np.array(cube["individual"]["rentabilidad"]["dolares"])
    tt_arr = np.array(cube["individual"]["rentabilidad"]["tt_pesos"])
    ipc_arr = np.array([x / 100.0 for x in ipc_monthly])
    fx_arr = np.array(fx_series)

    def calc_compound(rates, s_idx, e_idx):
        r = rates[s_idx:e_idx+1]
        m = len(r)
        if m == 0: return "0,00%", "0,00%"
        prod = np.prod(1.0 + np.array(r)) - 1.0
        anual = ((1.0 + prod) ** (12.0 / m)) - 1.0 if m > 0 else 0.0
        return f"{prod*100:,.2f}%".replace('.', ','), f"{anual*100:,.2f}%".replace('.', ',')

    def calc_fx(fx, s_idx, e_idx):
        m = e_idx - s_idx + 1
        base = fx[0] if s_idx == 0 else fx[s_idx - 1]
        end = fx[e_idx]
        prod = (end - base) / base
        anual = ((1.0 + prod) ** (12.0 / m)) - 1.0 if m > 0 else 0.0
        return f"{prod*100:,.2f}%".replace('.', ','), f"{anual*100:,.2f}%".replace('.', ',')

    periods_def = [
        ("Último mes (Julio 2026 / L1M)", last_idx, last_idx),
        ("Últimos 3 meses (L3M)", max(0, last_idx - 2), last_idx),
        ("Últimos 6 meses (L6M)", max(0, last_idx - 5), last_idx),
        ("Ejercicio 2025/2026 (Cerrado)", date_to_idx.get('2025-07', 28), date_to_idx.get('2026-06', 39)),
        ("Ejercicio 2024/2025 (Cerrado)", date_to_idx.get('2024-07', 16), date_to_idx.get('2025-06', 27)),
        ("Ejercicio 2023/2024 (Cerrado)", date_to_idx.get('2023-07', 4), date_to_idx.get('2024-06', 15)),
        ("Ejercicio 2026/2027 (En curso)", date_to_idx.get('2026-07', last_idx), last_idx),
        ("Últimos 12 meses (L12M)", max(0, last_idx - 11), last_idx),
        ("Últimos 24 meses (L24M)", max(0, last_idx - 23), last_idx),
        ("Últimos 36 meses (L36M)", max(0, last_idx - 35), last_idx),
    ]

    rent_pesos_table = []
    rent_dolares_table = []
    benchmarks_table = []

    for name, s_idx, e_idx in periods_def:
        p_acum, p_anual = calc_compound(rent_p, s_idx, e_idx)
        rent_pesos_table.append({"p": name, "acum": p_acum, "anual": p_anual})
        
        d_acum, d_anual = calc_compound(rent_d, s_idx, e_idx)
        rent_dolares_table.append({"p": name, "acum": d_acum, "anual": d_anual})
        
        ipc_a, ipc_an = calc_compound(ipc_arr, s_idx, e_idx)
        tt_a, tt_an = calc_compound(tt_arr, s_idx, e_idx)
        tc_a, tc_an = calc_fx(fx_arr, s_idx, e_idx)
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

def main():
    parsed_data = load_and_parse_multidimensional_dataset(DATASET_PATH)
    dates = parsed_data["DATES"]
    cube = parsed_data["cube"]
    fx = parsed_data["FX"]
    ipc = parsed_data["IPC_MONTHLY"]
    last_idx = len(dates) - 1
    last_fx = fx[-1]

    # Validación Actuarial al Cierre 2026-07
    def get_ct_total(seg):
        tot = 0.0
        for prod in ["individual", "colectivo", "rvp_art"]:
            p_data = cube[prod].get(seg, {})
            # Pesos
            tot += p_data.get("pesos", {}).get("ct_activos", [0]*len(dates))[last_idx]
            tot += p_data.get("pesos", {}).get("ct_pasivos", [0]*len(dates))[last_idx]
            # Dólares pesificados
            tot += p_data.get("dolares", {}).get("ct_activos", [0]*len(dates))[last_idx] * last_fx
            tot += p_data.get("dolares", {}).get("ct_pasivos", [0]*len(dates))[last_idx] * last_fx
        return tot

    tot_neg = get_ct_total("negocio")
    tot_emp = get_ct_total("empleados")
    tot_vin = get_ct_total("vinculados")
    tot_gen = tot_neg + tot_emp + tot_vin

    print("\n=== VALIDACION ACTUARIAL AL CIERRE 2026-07 (TC: $1,485.00) ===")
    print(f"Negocio Propio:       $ {tot_neg:,.2f}  ($ {tot_neg/1e6:,.2f} M)")
    print(f"Plan Empleados:       $ {tot_emp:,.2f}  ($ {tot_emp/1e6:,.2f} M)")
    print(f"Plan Vinculados:      $ {tot_vin:,.2f}  ($ {tot_vin/1e6:,.2f} M)")
    print(f"CT Total Consolidado: $ {tot_gen:,.2f}  ($ {tot_gen/1e6:,.2f} M)")
    print("======================================================================")

    tables = compute_official_tables(dates, cube, fx, ipc)

    final_payload = {
        "metadata": {
            "origen": "Nuevo Dataset Retiro - Jul 26.xlsx",
            "cierre": dates[-1],
            "total_periodos": len(dates),
            "tipo_cambio_cierre": last_fx
        },
        "dates": dates,
        "macro": {
            "FX": fx,
            "IPC_MONTHLY": ipc
        },
        "cube": cube,
        "tables": tables
    }

    with open(OUTPUT_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(final_payload, f, ensure_ascii=False, indent=2)

    print(f"OK: {OUTPUT_JSON_PATH} actualizado exitosamente con el cubo multidimensional.")

    if "--deploy" in sys.argv:
        print("\nIniciando despliegue a GitHub Pages...")
        import deploy_to_github
        deploy_to_github.main()

if __name__ == "__main__":
    main()
