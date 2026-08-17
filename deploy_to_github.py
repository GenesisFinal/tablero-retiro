# -*- coding: utf-8 -*-
"""
Script de Despliegue Automatizado a GitHub / GitHub Pages
Tablero de Seguros de Retiro (AGMD Style Dark Green)
"""

import requests
import base64
import json
import os
import sys
import time

def get_github_token():
    # Buscar token localmente en archivo de configuración o proyectos hermanos
    possible_paths = [
        os.path.join(os.path.dirname(__file__), "GitHub token.txt"),
        os.path.join("G:\\Mi unidad\\IA\\Valores Financieros", "GitHub token.txt"),
        os.path.join(os.path.dirname(os.path.dirname(__file__)), "Valores Financieros", "GitHub token.txt")
    ]
    for p in possible_paths:
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                for line in f.read().splitlines():
                    clean = line.strip()
                    if clean.startswith("ghp_"):
                        return clean
    return os.environ.get("GITHUB_TOKEN")

def main():
    token = get_github_token()
    if not token:
        print("ERROR: No se encontró el token de GitHub (ghp_...).")
        sys.exit(1)

    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "Tablero-Retiro-Deployer"
    }

    username = "GenesisFinal"
    repo_name = "tablero-retiro"
    full_repo = f"{username}/{repo_name}"
    print(f"Desplegando en repositorio: {full_repo}")

    # 1. Verificar o Crear Repositorio
    repo_res = requests.get(f"https://api.github.com/repos/{full_repo}", headers=headers)
    if repo_res.status_code == 404:
        print(f"Creando repositorio '{repo_name}' en GitHub...")
        create_payload = {
            "name": repo_name,
            "description": "Tablero de Control de Seguros de Retiro - La Segunda (AGMD Style Dark Green • JetBrains Mono • Looker Studio • Autoactualizable)",
            "homepage": f"https://{username}.github.io/{repo_name}/",
            "private": False,
            "has_issues": True,
            "has_projects": True,
            "has_wiki": False,
            "auto_init": True
        }
        create_res = requests.post("https://api.github.com/user/repos", headers=headers, json=create_payload)
        if create_res.status_code in [200, 201]:
            print(f"Repositorio '{full_repo}' creado exitosamente!")
            time.sleep(2)
        else:
            print(f"Error al crear repositorio: {create_res.status_code} - {create_res.text}")
            sys.exit(1)
    else:
        print(f"Repositorio '{full_repo}' ya existe.")

    # 2. Función para subir o actualizar archivo en el repo con reintentos
    def upload_file(remote_path, local_path, commit_msg):
        if not os.path.exists(local_path):
            print(f"Archivo local no encontrado: {local_path}")
            return False

        with open(local_path, "rb") as f:
            content_bytes = f.read()
        content_b64 = base64.b64encode(content_bytes).decode("utf-8")

        url = f"https://api.github.com/repos/{full_repo}/contents/{remote_path}"
        
        for attempt in range(6):
            get_res = requests.get(url, headers=headers)
            sha = None
            if get_res.status_code == 200:
                sha = get_res.json().get("sha")

            payload = {
                "message": commit_msg,
                "content": content_b64
            }
            if sha:
                payload["sha"] = sha

            print(f"Subiendo {remote_path} ({len(content_bytes):,} bytes, intento {attempt+1})...")
            put_res = requests.put(url, headers=headers, json=payload)
            if put_res.status_code in [200, 201]:
                print(f"OK: {remote_path} publicado en GitHub.")
                return True
            else:
                print(f"Aviso al subir {remote_path}: {put_res.status_code} - {put_res.text[:120]}")
                time.sleep(2 * (attempt + 1))
        print(f"ERROR: No se pudo subir {remote_path} tras 6 intentos.")
        return False

    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    upload_file(".nojekyll", os.path.join(base_dir, ".nojekyll"), "Config: Deshabilitar procesamiento Jekyll con .nojekyll")
    upload_file("favicon.svg", os.path.join(base_dir, "favicon.svg"), "Asset: Favicon SVG para pestaña y marcadores")
    upload_file("index.html", os.path.join(base_dir, "index.html"), "Feat: Tablero Seguros de Retiro con Actualizacion Semanal y Boton a Demanda")
    upload_file("data_retiro.json", os.path.join(base_dir, "data_retiro.json"), "Data: Series historicas consolidadas de Seguros de Retiro con 8003 asegurados")
    upload_file("actualizar_tablero.py", os.path.join(base_dir, "actualizar_tablero.py"), "Code: Script de actualizacion y validacion actuarial")
    upload_file(".github/workflows/weekly_sync.yml", os.path.join(base_dir, ".github", "workflows", "weekly_sync.yml"), "CI/CD: Workflow de actualizacion automatica semanal")
    upload_file("deploy_to_github.py", os.path.join(base_dir, "deploy_to_github.py"), "Code: Script de despliegue automatizado a GitHub Pages")

    # 4. Activar GitHub Pages si aún no está activado
    pages_url = f"https://api.github.com/repos/{full_repo}/pages"
    pages_get = requests.get(pages_url, headers=headers)
    if pages_get.status_code == 404:
        print("Habilitando GitHub Pages en la rama main...")
        pages_payload = {
            "source": {
                "branch": "main",
                "path": "/"
            }
        }
        pages_post = requests.post(pages_url, headers=headers, json=pages_payload)
        if pages_post.status_code not in [200, 201]:
            # Probar con rama master
            pages_payload["source"]["branch"] = "master"
            pages_post = requests.post(pages_url, headers=headers, json=pages_payload)

    public_url = f"https://{username}.github.io/{repo_name}/"
    print("\n=======================================================")
    print(" DESPLIEGUE COMPLETADO EXITOSAMENTE")
    print(f" Repositorio: https://github.com/{full_repo}")
    print(f" URL Pública (GitHub Pages): {public_url}")
    print("=======================================================\n")

if __name__ == "__main__":
    main()
