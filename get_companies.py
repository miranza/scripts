import json
from pathlib import Path
from urllib.parse import quote

import requests

from ms_graph import get_bc_access_token
from settings import BC_ENVIRONMENT, BC_TENANT_ID

URL = (
    f"https://api.businesscentral.dynamics.com/v2.0/{BC_TENANT_ID}"
    f"/{BC_ENVIRONMENT}/api/v2.0/companies"
)

SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = SCRIPT_DIR / "data"
JSON_PATH = DATA_DIR / "companies.json"


def get_companies():
    """Devuelve la lista de compañías del entorno de BC."""
    resp = requests.get(
        URL,
        headers={
            "Authorization": get_bc_access_token(),
            "Content-Type": "application/json",
        },
        timeout=60,
    )
    resp.raise_for_status()
    return resp.json().get("value", [])


def as_empresa(company):
    """Mapea una compañía de BC a una entrada de EMPRESAS."""
    name = company["name"]
    display_name = company.get("displayName") or name
    return {
        "company_name": quote(name, safe=""),
        "api_name": display_name,
        "display_name": name,
        "sociedad": display_name,
        "company_id": company["id"],
    }


def format_empresa(empresa):
    """Formatea la entrada como literal Python indentado, con coma final."""
    lines = ["    {"]
    for key, value in empresa.items():
        lines.append(f'        "{key}": {json.dumps(value, ensure_ascii=False)},')
    lines.append("    },")
    return "\n".join(lines)


def main():
    companies = get_companies()
    empresas = [
        as_empresa(c)
        for c in sorted(companies, key=lambda c: c.get("displayName") or c["name"])
    ]

    for empresa in empresas:
        print(format_empresa(empresa))

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(empresas, f, ensure_ascii=False, indent=4)
        f.write("\n")

    print(f"\n{len(empresas)} compañías guardadas en {JSON_PATH}")


if __name__ == "__main__":
    main()
