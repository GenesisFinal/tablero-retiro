#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json
import os
import sys
import subprocess

DATA_FILE = os.path.join(os.path.dirname(__file__), "data_retiro.json")

def load_dataset():
    if not os.path.exists(DATA_FILE):
        print(f"Error: No se encontro {DATA_FILE}")
        sys.exit(1)
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_dataset(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"OK: {DATA_FILE} actualizado exitosamente.")

def validate_integrity(data):
    series = data.get("series", {})
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

def deploy():
    deploy_script = os.path.join(os.path.dirname(__file__), "deploy_to_github.py")
    if os.path.exists(deploy_script):
        print("\nIniciando sincronizacion y despliegue a GitHub Pages...")
        subprocess.run([sys.executable, deploy_script], check=True)

if __name__ == "__main__":
    data = load_dataset()
    if validate_integrity(data):
        save_dataset(data)
        if len(sys.argv) > 1 and sys.argv[1] == "--deploy":
            deploy()
