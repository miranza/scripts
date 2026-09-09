# Script_movcont_csv.py
# Versión: 2025-08-30 (rev con formato Excel, aatcompany, validación, menú empresas)
# Guardado en el MISMO DIRECTORIO DEL SCRIPT.
# Salidas fijas: movcont_historical.csv / movcont_historical.xlsx

import logging
import shutil
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import requests

from logger import setup_log
from ms_graph import get_bc_access_token, sync_sharepoint
from settings import BC_ENVIRONMENT, BC_TENANT_ID, EMPRESAS

setup_log("movcont")
log = logging.getLogger(__name__)

BASE_URL = f"https://api.businesscentral.dynamics.com/v2.0/{BC_TENANT_ID}/{BC_ENVIRONMENT}/ODataV4"

SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = SCRIPT_DIR / "data"
FILENAME = "movcont_historical"
CSV_PATH = DATA_DIR / f"{FILENAME}.csv"
XLSX_PATH = DATA_DIR / f"{FILENAME}.xlsx"


def last_month_range():
    today = datetime.now()
    first_this = today.replace(day=1)
    last_prev = first_this - timedelta(days=1)
    first_prev = last_prev.replace(day=1)
    return first_prev.strftime("%Y-%m-%d"), last_prev.strftime("%Y-%m-%d")


def build_filter(start_date, end_date, acc_from=None, acc_to=None):
    parts = [f"Posting_Date ge {start_date}", f"Posting_Date le {end_date}"]
    if acc_from:
        parts.append(f"G_L_Account_No ge '{acc_from}'")
    if acc_to:
        parts.append(f"G_L_Account_No le '{acc_to}'")
    return " and ".join(parts)


# ---- Limpieza y coerción numérica robusta ----
def _text_to_float_es(value):
    """Convierte '1.234,56' o '1,234.56' o '1234,56'/'1234.56' en float."""
    if value is None:
        return None
    s = str(value).strip()
    if s == "" or s.lower() in {"nan", "none"}:
        return None
    has_dot = "." in s
    has_com = "," in s
    if has_dot and has_com:
        # Decimales = el separador que aparece más a la derecha
        if s.rfind(",") > s.rfind("."):
            # miles '.' / dec ',' -> quitar '.' y cambiar ','->'.'
            s = s.replace(".", "").replace(",", ".")
        else:
            # miles ',' / dec '.' -> quitar ',' y dejar '.'
            s = s.replace(",", "")
    elif has_com and not has_dot:
        s = s.replace(",", ".")
    else:
        # solo '.' o ninguno -> ya vale
        pass
    try:
        return float(s)
    except Exception:
        return None


def coerce_numeric_es(df: pd.DataFrame, cols):
    for c in cols:
        if c in df.columns:
            if not pd.api.types.is_numeric_dtype(df[c]):
                df[c] = df[c].apply(_text_to_float_es)
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def format_spanish_number(x):
    """Devuelve número como texto con miles '.' y decimales ',' (2 decimales)."""
    try:
        v = float(x)
    except Exception:
        return x
    s = f"{v:,.2f}"  # '1,234,567.89'
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


# ---- Menú de empresas en el orden solicitado ----
def companies_menu_order():
    labels = [
        "IMO",
        "Valcasado",
        "Andorra",
        "BACKUP_SGA",
        "Begitek Clinica Oftalmologica,",
        "Cordoba",
        "CAMACHO VINUESA",
        "CAPEVISION",
        "Palomares",
        "IMO Manresa",
        "COI",
        "CRM",
        "CRONUS ES",
        "CVL",
        "Fundacion Miranza",
        "IBO",
        "UOB",
        "IOA",
        "Begoña",
        "Galicia",
        "Gran Canaria",
        "MIO",
        "Malaga",
        "Miranza QX",
        "Santander",
        "CIMO",
        "Tenerife",
        "Oculsur",
        "TILICE",
        "Ophthalteam",
        "Algeciras",
        "OKULAR",
        "Seville -Tecnolaser",
        "Oporto",
        "Catalunya - Mataro",
    ]
    by_disp = {e["display_name"]: e for e in EMPRESAS}
    ordered, seen = [], set()
    for name in labels:
        e = by_disp.get(name)
        if not e:
            for cand in EMPRESAS:
                if (
                    cand["api_name"].upper() == name.upper()
                    or cand["company_name"].upper() == name.upper()
                ):
                    e = cand
                    break
        if e:
            key = (e["company_name"], e["api_name"])
            if key not in seen:
                ordered.append(e)
                seen.add(key)
    for e in EMPRESAS:
        key = (e["company_name"], e["api_name"])
        if key not in seen:
            ordered.append(e)
            seen.add(key)
    return ordered


def select_companies_interactive():
    ordered = companies_menu_order()
    print("\nSeleccione las empresas a procesar:")
    for i, e in enumerate(ordered, start=1):
        print(f"{i}. {e['display_name']}")
    print("0. Todas las empresas")
    raw = input("Ingrese el/los número(s) separados por coma (ej: 2,5,7) o 0 para todas: ").strip()
    if raw == "0" or raw == "":
        return ordered
    try:
        idxs = [int(x.strip()) for x in raw.split(",") if x.strip()]
        chosen = [ordered[i - 1] for i in idxs if 1 <= i <= len(ordered)]
        return chosen if chosen else ordered
    except Exception:
        return ordered


# ---- Enriquecidos y descarga ----
def fetch_company_movs(
    company, headers, start_date, end_date, acc_from=None, acc_to=None, page_size=1000
):
    comp_name = company.get("company_name") or company.get("api_name")
    if not comp_name:
        return pd.DataFrame()

    filter_q = build_filter(start_date, end_date, acc_from, acc_to)
    base = f"{BASE_URL}/Company('{comp_name}')/Movs_contabilidad_Excel"
    select = (
        "$select=Entry_No,Posting_Date,G_L_Account_No,G_L_Account_Name,Description,"
        "Amount,Debit_Amount,Credit_Amount,Source_Code,Document_No,Dimension_Set_ID, Bal_Account_No"
    )
    has_more, skip, rows = True, 0, []
    while has_more:
        url = f"{base}?$filter={filter_q}&{select}&$top={page_size}&$skip={skip}"
        r = requests.get(url, headers=headers, timeout=60)
        if r.status_code == 401:
            token = get_bc_access_token()
            headers["Authorization"] = f"Bearer {token}"
            r = requests.get(url, headers=headers, timeout=60)
        log.debug(f"FETCH COMPANY ROWS: {url}")
        r.raise_for_status()
        batch = r.json().get("value", [])
        rows.extend(batch)
        has_more = len(batch) == page_size
        skip += page_size

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    df["Empresa"] = company.get("display_name") or comp_name
    df["api_name"] = company.get("api_name")
    df["company_name"] = comp_name
    df["company_id"] = company.get("company_id")
    df["sociedad"] = company.get("sociedad")  # para aatcompany
    return df


def enrich_cost_center_for_specials(df, company, headers, page_size=1000):
    specials = {"IMO", "VALCASADO"}
    if (company.get("api_name") or "").upper() not in specials:
        df["CentroCoste"] = company.get("display_name") or company.get("company_name")
        return df

    comp_name = company.get("company_name") or company.get("api_name")
    base = f"{BASE_URL}/Company('{comp_name}')/DimensionSetEntries?$filter=Dimension_Code eq 'COST CENTER'"
    has_more, skip, rows = True, 0, []
    while has_more:
        url = f"{base}&$top={page_size}&$skip={skip}"
        r = requests.get(url, headers=headers, timeout=60)
        if r.status_code == 401:
            headers["Authorization"] = get_bc_access_token()
            r = requests.get(url, headers=headers, timeout=60)
        log.debug(f"ENRICH COST CENTER: {url}")
        r.raise_for_status()
        batch = r.json().get("value", [])
        rows.extend(batch)
        has_more = len(batch) == page_size
        skip += page_size

    if rows:
        dfd = pd.DataFrame(rows)[["Dimension_Set_ID", "Dimension_Value_Name"]].drop_duplicates()
        df = df.merge(dfd, on="Dimension_Set_ID", how="left")
        df["CentroCoste"] = df["Dimension_Value_Name"].fillna(company.get("display_name"))
        df.drop(columns=["Dimension_Value_Name"], inplace=True, errors="ignore")
    else:
        df["CentroCoste"] = company.get("display_name") or company.get("company_name")
    return df


def add_aatcompany(df):
    """Crea la columna aatcompany a partir de sociedad con reglas especiales."""
    df["aatcompany"] = df.get("sociedad")
    if "api_name" in df.columns and "CentroCoste" in df.columns:
        # IMO
        mask_imo = df["api_name"].str.upper().eq("IMO")
        if mask_imo.any():
            cc = df.loc[mask_imo, "CentroCoste"].astype(str).str.lower()
            df.loc[mask_imo & cc.str.contains("barcelona"), "aatcompany"] = (
                "INSTITUTO DE MICROCIRUGIA OCULAR  DOS S.LU. BCN"
            )
            df.loc[mask_imo & cc.str.contains("madrid"), "aatcompany"] = (
                "INSTITUTO DE MICROCIRUGIA OCULAR  DOS S.LU. MADRID"
            )
        # VALCASADO
        mask_val = df["api_name"].str.upper().eq("VALCASADO")
        if mask_val.any():
            cc = df.loc[mask_val, "CentroCoste"].astype(str).str.lower()
            df.loc[mask_val & cc.str.contains("albacete"), "aatcompany"] = "Valcasado, SAU Albacete"
            df.loc[mask_val & cc.str.contains("alicante"), "aatcompany"] = "Valcasado, SAU Alicante"
            df.loc[mask_val & cc.str.contains("getafe"), "aatcompany"] = "Valcasado, SAU Getafe"
    return df


def normalize_types(df):
    if "Posting_Date" in df.columns:
        df["Posting_Date"] = pd.to_datetime(df["Posting_Date"], errors="coerce")
    num_cols = ["Amount", "Debit_Amount", "Credit_Amount"]
    df = coerce_numeric_es(df, num_cols)
    return df


def ensure_columns_order(df):
    cols_pref = [
        "Empresa",
        "api_name",
        "company_name",
        "company_id",
        "sociedad",
        "aatcompany",
        "CentroCoste",
        "Posting_Date",
        "Entry_No",
        "Document_No",
        "G_L_Account_No",
        "G_L_Account_Name",
        "Description",
        "Amount",
        "Debit_Amount",
        "Credit_Amount",
        "Source_Code",
        "Dimension_Set_ID",
        "Bal_Account_No",
    ]
    for c in cols_pref:
        if c not in df.columns:
            df[c] = pd.NA
    return df[cols_pref]


def validate_historical(df):
    msgs = []
    if df.empty:
        return ["Histórico vacío."]
    processed = set(df.get("Empresa", pd.Series(dtype=str)).dropna().unique())
    expected = set([e["display_name"] for e in EMPRESAS])
    missing = expected - processed
    if missing:
        msgs.append(f"Empresas no presentes en el histórico: {', '.join(sorted(missing))}")
    if "Posting_Date" in df.columns:
        try:
            dmin = pd.to_datetime(df["Posting_Date"], errors="coerce").min()
            dmax = pd.to_datetime(df["Posting_Date"], errors="coerce").max()
            msgs.append(f"Rango de fechas detectado: {str(dmin)[:10]} a {str(dmax)[:10]}")
        except Exception:
            msgs.append("No se pudo calcular el rango de fechas.")
    for f in ["G_L_Account_No", "Amount", "Empresa", "Entry_No"]:
        if f in df.columns:
            n = df[f].isna().sum()
            if n > 0:
                msgs.append(f"Valores nulos en {f}: {n}")
    if all(c in df.columns for c in ["Empresa", "Entry_No", "Posting_Date"]):
        dup = df.duplicated(subset=["Empresa", "Entry_No", "Posting_Date"]).sum()
        if dup > 0:
            msgs.append(f"Posibles duplicados (Empresa, Entry_No, Posting_Date): {dup}")
    return msgs or ["Validación OK: sin incidencias relevantes."]


def backup_file(path: Path):
    if path.exists():
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        bkp = path.with_name(f"{path.stem}__backup_{ts}{path.suffix}")
        shutil.copy2(path, bkp)
        return bkp
    return None


# ---- Escrituras ----
def write_csv_spanish(df: pd.DataFrame, path: Path):
    # Convertimos a texto con formato español SÓLO para la salida CSV
    cols = [c for c in ["Amount", "Debit_Amount", "Credit_Amount"] if c in df.columns]
    df_out = df.copy()
    for c in cols:
        df_out[c] = df_out[c].apply(format_spanish_number)
    df_out.to_csv(path, index=False, encoding="utf-8", sep=";", decimal=",")


def write_excel_spanish(df: pd.DataFrame, path: Path):
    # Asegurar que las columnas numéricas sean realmente numéricas
    cols = [c for c in ["Amount", "Debit_Amount", "Credit_Amount"] if c in df.columns]
    df_x = df.copy()
    df_x = coerce_numeric_es(df_x, cols)

    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        df_x.to_excel(xw, sheet_name="MovCont", index=False)
        summary = pd.DataFrame(
            {
                "KPI": ["Filas", "Fecha mínima", "Fecha máxima", "Empresas distintas"],
                "Valor": [
                    len(df_x),
                    (
                        str(pd.to_datetime(df_x["Posting_Date"], errors="coerce").min())[:10]
                        if "Posting_Date" in df_x
                        else ""
                    ),
                    (
                        str(pd.to_datetime(df_x["Posting_Date"], errors="coerce").max())[:10]
                        if "Posting_Date" in df_x
                        else ""
                    ),
                    df_x["Empresa"].nunique() if "Empresa" in df_x else 0,
                ],
            }
        )
        summary.to_excel(xw, sheet_name="Resumen", index=False)

        present = [c for c in cols if c in df_x.columns]
        ws = xw.book["MovCont"]
        headers = [cell.value for cell in ws[1]]
        for c in present:
            try:
                idx = headers.index(c) + 1  # 1-based
                for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
                    for cell in row:
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = "#.##0,00"
            except ValueError:
                continue


def save_outputs(df, start_date, end_date, selected_companies=None):
    """
    Escribe movcont_historical.{csv,xlsx}.
    - Backups automáticos si existen.
    - Reemplaza filas del histórico dentro del rango [start_date,end_date]
      y (si se seleccionaron) sólo de esas empresas; luego añade lo nuevo.
    """
    # Cargar histórico existente (parseando miles y decimales españoles)
    hist = pd.DataFrame()
    if CSV_PATH.exists():
        hist = pd.read_csv(CSV_PATH, sep=";", decimal=",", thousands=".")
        if "Posting_Date" in hist.columns:
            hist["Posting_Date"] = pd.to_datetime(hist["Posting_Date"], errors="coerce")

    # Filtrar histórico para quitar el rango que vamos a reescribir
    if not hist.empty and "Posting_Date" in hist.columns:
        s = pd.to_datetime(start_date)
        e = pd.to_datetime(end_date)
        mask_range = (hist["Posting_Date"] >= s) & (hist["Posting_Date"] <= e)
        if selected_companies:
            sel_names = set([c.get("display_name") for c in selected_companies])
            mask_comp = hist["Empresa"].isin(sel_names)
            hist = hist[~(mask_range & mask_comp)]
        else:
            hist = hist[~mask_range]

    # Unir histórico restante + nuevo
    df_out = pd.concat([hist, df], ignore_index=True) if not hist.empty else df.copy()

    # Backups
    b1 = backup_file(CSV_PATH)
    b2 = backup_file(XLSX_PATH)
    if b1:
        log.info(f"Backup CSV creado: {b1}")
    if b2:
        log.info(f"Backup XLSX creado: {b2}")

    # Guardar CSV (formato español como texto) y Excel (numérico con formato)
    write_csv_spanish(df_out, CSV_PATH)
    write_excel_spanish(df_out, XLSX_PATH)

    return CSV_PATH, XLSX_PATH


def build_dataset(start_date, end_date, acc_from=None, acc_to=None, selected_companies=None):
    token = get_bc_access_token()
    headers = {"Authorization": token, "Content-Type": "application/json"}

    companies = selected_companies if selected_companies else EMPRESAS

    all_chunks = []
    for idx, comp in enumerate(companies, start=1):
        company_name = comp.get("display_name") or comp.get("company_name")
        rango = f"{start_date} - {end_date}"
        accounts = (
            f" | cuentas: [{acc_from or 'min'}..{acc_to or 'max'}]" if (acc_from or acc_to) else ""
        )
        log.info(f"[{idx}/{len(companies)}] {company_name} | {rango}{accounts}")
        try:
            df = fetch_company_movs(comp, headers, start_date, end_date, acc_from, acc_to)
            if not df.empty:
                df = enrich_cost_center_for_specials(df, comp, headers)
                df = add_aatcompany(df)
                df = normalize_types(df)
                df = ensure_columns_order(df)
                all_chunks.append(df)
        except Exception as e:
            log.error(f"Error en {company_name}: {e}")

    if not all_chunks:
        return pd.DataFrame()
    return pd.concat(all_chunks, ignore_index=True)


# ========= MODOS DE USO =========
def menu():
    print("\n" + "=" * 64)
    print(" MOVIMIENTOS CONTABLES (Business Central) ".center(64, "="))
    print("=" * 64)
    print("1) Construir histórico por rango de fechas")
    print("2) Actualizar mes pasado (reempalma en histórico)")
    print("3) Validar archivo histórico existente")
    print("4) Exportar snapshot del histórico a Excel")
    print("5) Salir")

    opt = input("\nOpción (1-5): ").strip()
    if opt == "1":
        s = input("Fecha inicio (YYYY-MM-DD): ").strip()
        e = input("Fecha fin    (YYYY-MM-DD): ").strip()
        acc_from = input("Cuenta desde (enter para omitir): ").strip() or None
        acc_to = input("Cuenta hasta (enter para omitir): ").strip() or None
        selected_companies = select_companies_interactive()

        df = build_dataset(s, e, acc_from, acc_to, selected_companies)
        if df.empty:
            print("\nNo se obtuvieron datos en el rango.")
            return
        csv_path, xlsx_path = save_outputs(df, s, e, selected_companies)
        print(f"\n✓ Histórico CSV:  {csv_path}")
        print(f"✓ Histórico Excel:{xlsx_path}")

    elif opt == "2":
        s, e = last_month_range()
        acc_from = input("Cuenta desde (enter para omitir): ").strip() or None
        acc_to = input("Cuenta hasta (enter para omitir): ").strip() or None
        selected_companies = select_companies_interactive()

        df = build_dataset(s, e, acc_from, acc_to, selected_companies)
        if df.empty:
            print("\nNo se obtuvieron datos del mes pasado.")
            return
        csv_path, xlsx_path = save_outputs(df, s, e, selected_companies)
        print(f"\n✓ Histórico CSV:  {csv_path}")
        print(f"✓ Histórico Excel:{xlsx_path}")

    elif opt == "3":
        if not CSV_PATH.exists():
            print("\nNo existe histórico para validar.")
            return
        df = pd.read_csv(CSV_PATH, sep=";", decimal=",", thousands=".")
        msgs = validate_historical(df)
        print("\nVALIDACIÓN DEL HISTÓRICO")
        for m in msgs:
            print(f"- {m}")

    elif opt == "4":
        if not CSV_PATH.exists():
            print("\nNo existe histórico para exportar a snapshot.")
            return
        df = pd.read_csv(CSV_PATH, sep=";", decimal=",", thousands=".")
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        snap_xlsx = DATA_DIR / f"movcont_historico_snapshot_{ts}.xlsx"
        write_excel_spanish(df, snap_xlsx)
        print(f"\n✓ Snapshot Excel: {snap_xlsx}")

    elif opt == "5":
        print("Saliendo...")
        sys.exit(0)
    else:
        print("Opción no válida.")


def cli_mode(argv):
    """
    CLI opcional:
      python Script_movcont_csv.py run --start 2024-01-01 --end 2024-01-31 --acc-from 600000 --acc-to 699999
      python Script_movcont_csv.py update_last_month --acc-from 600000
    En modo CLI se usan todas las empresas por defecto.
    """
    if len(argv) < 2:
        return False

    cmd = argv[1].lower()
    if cmd == "run":

        def arg(name, default=None):
            if f"--{name}" in argv:
                return argv[argv.index(f"--{name}") + 1]
            return default

        s = arg("start")
        e = arg("end")
        if not (s and e):
            print("Faltan --start y/o --end")
            return True

        acc_from = arg("acc-from")
        acc_to = arg("acc-to")

        df = build_dataset(s, e, acc_from, acc_to, selected_companies=None)
        if df.empty:
            print("No se obtuvieron datos.")
            return True
        csv_path, xlsx_path = save_outputs(df, s, e, selected_companies=None)
        print(f"\n✓ Histórico CSV:  {csv_path}\n✓ Histórico Excel:{xlsx_path}")
        return True

    elif cmd == "update_last_month":

        def arg(name, default=None):
            if f"--{name}" in argv:
                return argv[argv.index(f"--{name}") + 1]
            return default

        acc_from = arg("acc-from")
        acc_to = arg("acc-to")
        s, e = last_month_range()
        df = build_dataset(s, e, acc_from, acc_to, selected_companies=None)
        if df.empty:
            print("No se obtuvieron datos del mes pasado.")
            return True
        csv_path, xlsx_path = save_outputs(df, s, e, selected_companies=None)
        print(f"\n✓ Histórico CSV:  {csv_path}\n✓ Histórico Excel:{xlsx_path}")
        return True

    return False


# # ========= MAIN =========
# if __name__ == "__main__":
#     print(f"Script MovCont iniciado - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
#     try:
#         _ = get_bc_access_token()
#     except Exception as e:
#         print(f"Error autenticación: {e}")
#         sys.exit(1)

#     if cli_mode(sys.argv):
#         sys.exit(0)

#     while True:
#         try:
#             menu()
#             input("\nEnter para continuar...")
#         except KeyboardInterrupt:
#             print("\nInterrumpido por USER.")
#             sys.exit(0)
#         except Exception as e:
#             print(f"\nError no controlado: {e}")


if __name__ == "__main__":
    log.info("--- MOVCONT START ---")

    start_date, end_date = last_month_range()
    acc_from = 600000
    acc_to = 699999

    data = build_dataset(start_date, end_date, acc_from, acc_to)
    csv_path, xlsx_path = save_outputs(data, start_date, end_date)

    sync_sharepoint(CSV_PATH)
    sync_sharepoint(XLSX_PATH)

    log.info("--- MOVCONT END ---")
