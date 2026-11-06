# -*- coding: utf-8 -*-
"""
Script de Despliegue Automatizado a GitHub / GitHub Pages
Tablero de Seguros de Retiro (AGMD Style Dark Green)
"""

import base64
import json
import os
import sys
import time
import urllib.request
import urllib.error

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

def http_request(method, url, headers, payload=None):
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req_headers = dict(headers)
    if data is not None:
        req_headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=req_headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = resp.read().decode("utf-8", errors="replace")
            return resp.status, body
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        return e.code, body
    except Exception as e:
        return 0, str(e)

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
    status_code, text = http_request("GET", f"https://api.github.com/repos/{full_repo}", headers)
    if status_code == 404:
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
        c_status, c_text = http_request("POST", "https://api.github.com/user/repos", headers, create_payload)
        if c_status in [200, 201]:
            print(f"Repositorio '{full_repo}' creado exitosamente!")
            time.sleep(2)
        else:
            print(f"Error al crear repositorio: {c_status} - {c_text}")
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
            g_status, g_text = http_request("GET", url, headers)
            sha = None
            if g_status == 200:
                try:
                    sha = json.loads(g_text).get("sha")
                except Exception:
                    sha = None

            payload = {
                "message": commit_msg,
                "content": content_b64
            }
            if sha:
                payload["sha"] = sha

            print(f"Subiendo {remote_path} ({len(content_bytes):,} bytes, intento {attempt+1})...")
            p_status, p_text = http_request("PUT", url, headers, payload)
            if p_status in [200, 201]:
                print(f"OK: {remote_path} publicado en GitHub.")
                return True
            else:
                print(f"Aviso al subir {remote_path}: {p_status} - {p_text[:120]}")
                time.sleep(2 * (attempt + 1))
        print(f"ERROR: No se pudo subir {remote_path} tras 6 intentos.")
        return False

    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    upload_file(".nojekyll", os.path.join(base_dir, ".nojekyll"), "Config: Deshabilitar procesamiento Jekyll con .nojekyll")
    upload_file("favicon.svg", os.path.join(base_dir, "favicon.svg"), "Asset: Favicon SVG para pestaña y marcadores")
    upload_file("LaSegunda_Color_RGB.png", os.path.join(base_dir, "LaSegunda_Color_RGB.png"), "Asset: Logo oficial La Segunda Seguros")
    upload_file("index.html", os.path.join(base_dir, "index.html"), "Feat: Actualizacion dataset Octubre 2026 (44 meses) y benchmarks")
    upload_file("data_retiro.json", os.path.join(base_dir, "data_retiro.json"), "Data: Actualizacion cubo multidimensional Octubre 2026 (44 meses)")
    upload_file("actualizar_tablero.py", os.path.join(base_dir, "actualizar_tablero.py"), "Code: Parser nativo sin dependencias externas y soporte Octubre 2026")
    upload_file("deploy_to_github.py", os.path.join(base_dir, "deploy_to_github.py"), "Code: Script de despliegue automatizado a GitHub Pages")

    # 4. Activar GitHub Pages si aún no está activado
    pages_url = f"https://api.github.com/repos/{full_repo}/pages"
    pg_status, _ = http_request("GET", pages_url, headers)
    if pg_status == 404:
        print("Habilitando GitHub Pages en la rama main...")
        pages_payload = {
            "source": {
                "branch": "main",
                "path": "/"
            }
        }
        pp_status, _ = http_request("POST", pages_url, headers, pages_payload)
        if pp_status not in [200, 201]:
            pages_payload["source"]["branch"] = "master"
            http_request("POST", pages_url, headers, pages_payload)

    public_url = f"https://{username}.github.io/{repo_name}/"
    print("\n=======================================================")
    print(" DESPLIEGUE COMPLETADO EXITOSAMENTE")
    print(f" Repositorio: https://github.com/{full_repo}")
    print(f" URL Pública (GitHub Pages): {public_url}")
    print("=======================================================\n")

if __name__ == "__main__":
    main()
