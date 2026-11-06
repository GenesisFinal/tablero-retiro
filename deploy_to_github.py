# -*- coding: utf-8 -*-
"""
Script de Despliegue Automatizado a GitHub / GitHub Pages
Tablero de Seguros de Retiro (AGMD Style Dark Green)
Usa Git Data API (Trees/Commits) para publicar todos los archivos en UN SOLO commit atómico,
evitando que múltiples commits consecutivos cancelen el despliegue de GitHub Pages.
"""

import base64
import json
import os
import sys
import time
import urllib.request
import urllib.error

def get_github_token():
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

    # 1.Obtener referencia actual de main
    ref_status, ref_text = http_request("GET", f"https://api.github.com/repos/{full_repo}/git/ref/heads/main", headers)
    if ref_status != 200:
        print(f"Error al obtener ref heads/main: {ref_status} - {ref_text}")
        sys.exit(1)
    latest_commit_sha = json.loads(ref_text)["object"]["sha"]

    commit_status, commit_text = http_request("GET", f"https://api.github.com/repos/{full_repo}/git/commits/{latest_commit_sha}", headers)
    base_tree_sha = json.loads(commit_text)["tree"]["sha"]

    base_dir = os.path.dirname(os.path.abspath(__file__))
    files_to_publish = [
        ".nojekyll",
        "favicon.svg",
        "LaSegunda_Color_RGB.png",
        "index.html",
        "data_retiro.json",
        "actualizar_tablero.py",
        "deploy_to_github.py",
    ]

    tree_items = []
    for fname in files_to_publish:
        local_path = os.path.join(base_dir, fname)
        if not os.path.exists(local_path):
            continue
        with open(local_path, "rb") as f:
            content_bytes = f.read()
        content_b64 = base64.b64encode(content_bytes).decode("utf-8")
        b_status, b_text = http_request(
            "POST",
            f"https://api.github.com/repos/{full_repo}/git/blobs",
            headers,
            {"content": content_b64, "encoding": "base64"}
        )
        if b_status not in [200, 201]:
            print(f"Error creando blob para {fname}: {b_status} - {b_text}")
            sys.exit(1)
        blob_sha = json.loads(b_text)["sha"]
        print(f"Preparado {fname} ({len(content_bytes):,} bytes) -> blob {blob_sha[:8]}")
        tree_items.append({
            "path": fname,
            "mode": "100644",
            "type": "blob",
            "sha": blob_sha
        })

    # Crear nuevo árbol
    t_status, t_text = http_request(
        "POST",
        f"https://api.github.com/repos/{full_repo}/git/trees",
        headers,
        {"base_tree": base_tree_sha, "tree": tree_items}
    )
    if t_status not in [200, 201]:
        print(f"Error creando tree: {t_status} - {t_text}")
        sys.exit(1)
    new_tree_sha = json.loads(t_text)["sha"]

    # Crear un único commit atómico
    commit_msg = "Feat: Actualizacion Tablero Octubre 2026 (44 meses) y despliegue atomico"
    nc_status, nc_text = http_request(
        "POST",
        f"https://api.github.com/repos/{full_repo}/git/commits",
        headers,
        {
            "message": commit_msg,
            "tree": new_tree_sha,
            "parents": [latest_commit_sha]
        }
    )
    if nc_status not in [200, 201]:
        print(f"Error creando commit: {nc_status} - {nc_text}")
        sys.exit(1)
    new_commit_sha = json.loads(nc_text)["sha"]

    # Actualizar ref heads/main
    u_status, u_text = http_request(
        "PATCH",
        f"https://api.github.com/repos/{full_repo}/git/refs/heads/main",
        headers,
        {"sha": new_commit_sha, "force": False}
    )
    if u_status not in [200, 201]:
        print(f"Error actualizando main: {u_status} - {u_text}")
        sys.exit(1)

    print(f"OK: Commit atómico publicado en main ({new_commit_sha[:8]}).")

    public_url = f"https://{username}.github.io/{repo_name}/"
    print("\n=======================================================")
    print(" DESPLIEGUE COMPLETADO EXITOSAMENTE")
    print(f" Repositorio: https://github.com/{full_repo}")
    print(f" URL Pública (GitHub Pages): {public_url}")
    print("=======================================================\n")

if __name__ == "__main__":
    main()
