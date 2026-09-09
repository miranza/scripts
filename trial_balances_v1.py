import calendar
import logging
import os
import shutil
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import requests

from logger import setup_log
from ms_graph import get_bc_access_token, sync_sharepoint
from settings import BC_ENVIRONMENT, BC_TENANT_ID, EMPRESAS

setup_log("trial_balances")

log = logging.getLogger(__name__)

BASE_URL = f"https://api.businesscentral.dynamics.com/v2.0/{BC_TENANT_ID}/{BC_ENVIRONMENT}/api/att/miranzaBC/v2.0"

SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = SCRIPT_DIR / "data"
FILENAME = "trial_balance_historico"
CSV_PATH = DATA_DIR / f"{FILENAME}.csv"
XLSX_PATH = DATA_DIR / f"{FILENAME}.xlsx"

BATCH_SIZE = 1000


def get_last_day(year, month):
    """Obtiene el último día del mes"""
    return calendar.monthrange(year, month)[1]


def get_trial_balance_data(empresa, start_date, end_date):
    """Obtiene datos Trial Balance combinando OData y (si aplica) REST trialBalances para IMO / VALCASADO."""

    log.info(f"[{empresa['display_name']}] Get trial balance data")

    token = get_bc_access_token()
    if not token:
        return pd.DataFrame()

    headers = {"Authorization": token, "Content-Type": "application/json"}

    all_dataframes = []

    # --- PRIMERA EXTRACCIÓN: OData TrialBalanceSheet ---
    try:
        company_segment = empresa["company_name"]
        odata_url = (
            f"https://api.businesscentral.dynamics.com/v2.0/{BC_TENANT_ID}/{BC_ENVIRONMENT}/ODataV4/"
            f"Company('{company_segment}')/TrialBalanceSheet"
            f"?$select=number,accountid,accountType,display,totalDebit,totalCredit,balanceAtDateDebit,balanceAtDateCredit,dateFilter"
            f"&$filter=dateFilter ge {start_date} and dateFilter le {end_date}"
        )

        response = requests.get(odata_url, headers=headers, timeout=120)
        response.raise_for_status()
        odata = response.json().get("value", [])

        if odata:
            df_odata = pd.DataFrame(odata)
            df_odata["Empresa"] = empresa["display_name"]
            df_odata["Sociedad"] = empresa["sociedad"]
            df_odata["company_id"] = empresa["company_id"]

            df_odata = df_odata[df_odata["accountType"] == "Posting"]

            # Limpieza de números según Power Query M
            cols_to_process = [
                "totalDebit",
                "totalCredit",
                "balanceAtDateDebit",
                "balanceAtDateCredit",
            ]
            for col in cols_to_process:
                if col in df_odata.columns:
                    df_odata[col] = df_odata[col].astype(str).str.replace(",", "", regex=False)
                    df_odata[col] = df_odata[col].astype(str).str.replace(".", ",", regex=False)
                    df_odata[col] = pd.to_numeric(
                        df_odata[col].str.replace(",", ".", regex=False), errors="coerce"
                    )
                else:
                    log.warning(
                        f"Advertencia: La columna '{col}' no existe en los datos de {empresa['display_name']}."
                    )

            try:
                df_odata["number"] = df_odata["number"].astype("Int64")
            except:
                pass

            suma = df_odata[
                ["totalDebit", "totalCredit", "balanceAtDateDebit", "balanceAtDateCredit"]
            ].sum(axis=1, skipna=False)
            df_odata = df_odata[(suma != 0) | (suma.isna())].copy()

            # Estandarizar dateFilter a YYYY-MM-DD y eliminar filas en blanco
            if "dateFilter" in df_odata.columns:
                df_odata["dateFilter"] = pd.to_datetime(
                    df_odata["dateFilter"], errors="coerce"
                ).dt.strftime("%Y-%m-%d")
                df_odata = df_odata[df_odata["dateFilter"].notna()]

            df_odata["Month"] = pd.to_datetime(df_odata["dateFilter"], errors="coerce").dt.month
            df_odata["Year"] = pd.to_datetime(df_odata["dateFilter"], errors="coerce").dt.year

            df_odata["attCompany"] = empresa["sociedad"]
            df_odata["attCostCenter"] = empresa["sociedad"]
            df_odata["Sociedad_attCostCenter"] = empresa["sociedad"]

            all_dataframes.append(df_odata)
    except Exception as e:
        log.error(f"Error en extracción OData para {empresa['display_name']}: {e}")

    # --- SEGUNDA EXTRACCIÓN: REST solo para IMO , Miranza Santander S.L.U. y VALCASADO ---
    if empresa["api_name"] in {"IMO", "VALCASADO"}:
        try:
            allowed_cost_centers = {"BARCELONA", "MADRID", "ALICANTE", "ALBACETE", "GETAFE", "I+D"}
            batch = []
            skip = 0
            has_more = True

            while has_more:
                rest_url = (
                    f"{BASE_URL}/companies({empresa['company_id']})/trialBalances"
                    f"?$filter=dateFilter ge {start_date} and dateFilter le {end_date}"
                    f"&$top={BATCH_SIZE}&$skip={skip}"
                )
                response = requests.get(rest_url, headers=headers, timeout=60)
                response.raise_for_status()
                result = response.json().get("value", [])
                batch.extend(result)
                skip += BATCH_SIZE
                has_more = len(result) == BATCH_SIZE

            if batch:
                df_rest = pd.DataFrame(batch)
                df_rest = df_rest[df_rest["accountType"] == "Posting"]

            # Primero convierto todas las columnas a numérico
            for col in cols_to_process:
                if col in df_rest.columns:
                    df_rest[col] = df_rest[col].astype(str).str.replace(",", "", regex=False)
                    df_rest[col] = df_rest[col].astype(str).str.replace(".", ",", regex=False)
                    df_rest[col] = pd.to_numeric(
                        df_rest[col].str.replace(",", ".", regex=False), errors="coerce"
                    )
                else:
                    log.warning(
                        f"Advertencia: La columna '{col}' no existe en los datos de {empresa['display_name']}."
                    )

            try:
                df_rest["number"] = df_rest["number"].astype("Int64")
            except:
                pass

            # Ahora hago la suma y el filtrado
            suma = df_rest[
                ["totalDebit", "totalCredit", "balanceAtDateDebit", "balanceAtDateCredit"]
            ].sum(axis=1, skipna=False)
            df_rest = df_rest[(suma != 0) | (suma.isna())].copy()

            # Estandarizar dateFilter a YYYY-MM-DD y eliminar filas en blanco
            if "dateFilter" in df_rest.columns:
                df_rest["dateFilter"] = pd.to_datetime(
                    df_rest["dateFilter"], errors="coerce"
                ).dt.strftime("%Y-%m-%d")
                df_rest = df_rest[df_rest["dateFilter"].notna()]

            df_rest["Month"] = pd.to_datetime(df_rest["dateFilter"], errors="coerce").dt.month
            df_rest["Year"] = pd.to_datetime(df_rest["dateFilter"], errors="coerce").dt.year

            df_rest["Empresa"] = empresa["display_name"]
            df_rest["Sociedad"] = empresa["sociedad"]
            df_rest["company_id"] = empresa["company_id"]

            df_rest = df_rest[
                df_rest["attCostCenter"].str.upper().isin(allowed_cost_centers)
            ].copy()

            def map_attCompany(row):
                soc = row["Sociedad"]
                cc = row["attCostCenter"].upper()
                if soc == "INSTITUTO DE MICROCIRUGIA OCULAR  DOS S.LU.":
                    if cc == "BARCELONA":
                        return "INSTITUTO DE MICROCIRUGIA OCULAR  DOS S.LU. BCN"
                    elif cc == "MADRID":
                        return "INSTITUTO DE MICROCIRUGIA OCULAR  DOS S.LU. MADRID"
                elif soc == "Valcasado, SAU":
                    if cc == "ALICANTE":
                        return "Valcasado, SAU Alicante"
                    elif cc == "ALBACETE":
                        return "Valcasado, SAU Albacete"
                    elif cc == "GETAFE":
                        return "Valcasado, SAU Getafe"
                    elif cc == "I+D":
                        return "Valcasado, SAU I+D"
                return soc

            df_rest["attCompany"] = df_rest.apply(map_attCompany, axis=1)
            df_rest["Sociedad_attCostCenter"] = df_rest["attCompany"]

            all_dataframes.append(df_rest)
        except Exception as e:
            log.error(f"[{empresa['display_name']}]: Error en extracción REST: {e}")

    if all_dataframes:
        df_final = pd.concat(all_dataframes, ignore_index=True)
        output_columns = [
            "Year",
            "Month",
            "dateFilter",
            "attCompany",
            "attCostCenter",
            "Empresa",
            "Sociedad",
            "Sociedad_attCostCenter",
            "company_id",
            "number",
            "accountId",
            "accountType",
            "display",
            "totalDebit",
            "totalCredit",
            "balanceAtDateDebit",
            "balanceAtDateCredit",
        ]
        output_columns_present = [col for col in output_columns if col in df_final.columns]
        return df_final[output_columns_present].copy()
    else:
        log.warning(f"[{empresa['display_name']}]: No se obtuvieron datos")
        return pd.DataFrame()


def save_to_downloads(df):
    """Guarda el dataframe en CSV y XLSX, asegurando que ambos estén sincronizados siempre."""
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        df.to_csv(CSV_PATH, index=False, encoding="utf-8-sig", sep=";", decimal=",")
        log.info(f"Archivo guardado correctamente en: {CSV_PATH}")
        df_csv = pd.read_csv(CSV_PATH, sep=";", decimal=",")
        df_csv.to_excel(XLSX_PATH, index=False)
        log.info(f"Archivo Excel guardado correctamente en: {XLSX_PATH}")
        return True
    except Exception as e:
        log.error(f"Error al guardar archivo localmente: {e}")
        return False


# Sincronizar XLSX con CSV al inicio del script si solo existe el CSV o si el CSV es más reciente


def sync_xlsx_with_csv():
    """Si existe el CSV, asegura que el XLSX tenga los mismos datos."""
    if os.path.exists(CSV_PATH):
        csv_mtime = os.path.getmtime(CSV_PATH)
        xlsx_mtime = os.path.getmtime(XLSX_PATH) if os.path.exists(XLSX_PATH) else 0
        if not os.path.exists(XLSX_PATH) or csv_mtime > xlsx_mtime:
            try:
                df = pd.read_csv(CSV_PATH, sep=";", decimal=",")
                df.to_excel(XLSX_PATH, index=False)
                log.debug(f"Sincronizado: {XLSX_PATH} actualizado desde {CSV_PATH}")
            except Exception as e:
                log.error(f"Error al sincronizar XLSX con CSV: {e}")


def backup_historical_files():
    """Crea un backup automático del archivo histórico CSV y XLSX antes de sobrescribirlos."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if os.path.exists(CSV_PATH):
        backup_csv = CSV_PATH.with_name(f"{CSV_PATH.stem}_backup_{timestamp}.csv")
        try:
            shutil.copy2(CSV_PATH, backup_csv)
            log.debug(f"Backup creado: {backup_csv}")
        except Exception as e:
            log.error(f"Error al crear backup CSV: {e}")

    if os.path.exists(XLSX_PATH):
        backup_xlsx = XLSX_PATH.with_name(f"{XLSX_PATH.stem}_backup_{timestamp}.xlsx")
        try:
            shutil.copy2(XLSX_PATH, backup_xlsx)
            log.debug(f"Backup creado: {backup_xlsx}")
        except Exception as e:
            log.error(f"Error al crear backup XLSX: {e}")


def merge_and_replace_historical(new_data, start_date_str, end_date_str, empresas_a_procesar=None):
    """Elimina del histórico los datos del rango y empresas solicitados, agrega los nuevos y guarda ambos archivos. Backup automático antes de sobrescribir."""
    historical, _ = read_historical_file()
    # Estandarizar dateFilter en todos los DataFrames involucrados
    if "dateFilter" in historical.columns:
        historical["dateFilter"] = pd.to_datetime(
            historical["dateFilter"], errors="coerce"
        ).dt.strftime("%Y-%m-%d")
        historical = historical[historical["dateFilter"].notna()]
    if "dateFilter" in new_data.columns:
        new_data["dateFilter"] = pd.to_datetime(
            new_data["dateFilter"], errors="coerce"
        ).dt.strftime("%Y-%m-%d")
        new_data = new_data[new_data["dateFilter"].notna()]
    log.debug("[DEPURACIÓN] Histórico antes de filtrar")
    # log.debug(historical.info())
    if not historical.empty:
        log.debug(historical.head())
    # Determinar empresas a procesar
    if empresas_a_procesar is None:
        empresas_a_procesar = EMPRESAS
    empresas_nombres = [e["display_name"] for e in empresas_a_procesar]
    log.debug(f"[DEPURACIÓN] Empresas a filtrar: {empresas_nombres}")
    # Convertir fechas a datetime
    start_date = pd.to_datetime(start_date_str)
    end_date = pd.to_datetime(end_date_str)
    log.debug(f"[DEPURACIÓN] Rango de fechas a filtrar: {start_date} a {end_date}")
    if not historical.empty:
        if "dateFilter" in historical.columns and "Empresa" in historical.columns:
            log.debug(
                f"[DEPURACIÓN] Valores únicos de Empresa en histórico: {historical['Empresa'].unique()}"
            )
            log.debug(
                f"[DEPURACIÓN] Rango de fechas en histórico: {historical['dateFilter'].min()} a {historical['dateFilter'].max()}"
            )
            # Ya está en formato YYYY-MM-DD
            historical["dateFilter"] = pd.to_datetime(historical["dateFilter"], errors="coerce")
            mask = ~(
                (historical["Empresa"].isin(empresas_nombres))
                & (historical["dateFilter"] >= start_date)
                & (historical["dateFilter"] <= end_date)
            )
            filtered = historical[mask].copy()
            # Volver a estandarizar tras el filtrado
            filtered["dateFilter"] = pd.to_datetime(
                filtered["dateFilter"], errors="coerce"
            ).dt.strftime("%Y-%m-%d")
            filtered = filtered[filtered["dateFilter"].notna()]
            log.debug(f"[DEPURACIÓN] Filas antes de filtrar: {len(historical)}")
            log.debug(
                f"[DEPURACIÓN] Filas después de filtrar: {len(filtered)} (eliminadas: {len(historical) - len(filtered)})"
            )
        else:
            log.debug(
                "[DEPURACIÓN] No se encontraron columnas 'Empresa' y 'dateFilter' en el histórico. No se filtrará nada."
            )
            filtered = historical.copy()
    else:
        log.debug("[DEPURACIÓN] Histórico vacío antes de filtrar.")
        filtered = pd.DataFrame()
    # Concatenar el histórico filtrado con los datos nuevos
    log.debug(f"[DEPURACIÓN] Nuevos datos a agregar: {len(new_data)} filas")
    combined = pd.concat([filtered, new_data], ignore_index=True)
    # Estandarizar dateFilter en el combinado final
    if "dateFilter" in combined.columns:
        combined["dateFilter"] = pd.to_datetime(
            combined["dateFilter"], errors="coerce"
        ).dt.strftime("%Y-%m-%d")
        combined = combined[combined["dateFilter"].notna()]
    log.debug(f"[DEPURACIÓN] Total de filas antes de guardar: {len(combined)}")
    log.debug(combined.head())
    # Hacer backup antes de sobrescribir
    log.debug("Creando backup de archivos históricos...")
    backup_historical_files()
    # Guardar ambos archivos
    save_to_downloads(combined)
    return combined


# Llamar a la sincronización al inicio del script principal


def validate_data(df, start_date_str, end_date_str):
    """Realiza validaciones de consistencia en los datos del Trial Balance"""
    errors = []

    if df.empty:
        errors.append("El DataFrame está vacío. No hay datos para validar.")
        return errors

    # 1. Verificar empresas no procesadas
    processed_companies = df["Empresa"].unique()
    all_display_names = [e["display_name"] for e in EMPRESAS]
    missing_companies = set(all_display_names) - set(processed_companies)

    if missing_companies:
        errors.append(f"Empresas faltantes en los datos procesados: {', '.join(missing_companies)}")

    # 2. Verificar si hay datos (aunque la API no devuelve fecha por registro, el filtro se aplicó)
    # Se asume que si hay datos, son del rango solicitado.
    # Podríamos añadir una columna 'dateFilterApplied' si fuera necesario.
    errors.append(f"Rango de fechas solicitado para el filtro: {start_date_str} a {end_date_str}")

    # 3. Verificar valores nulos en campos clave esperados del Trial Balance
    # Las columnas clave podrían ser 'number' (G/L Account No.), y los campos de balance.
    key_fields = [
        "number",
        "totalDebit",
        "totalCredit",
        "balanceAtDateDebit",
        "balanceAtDateCredit",
        "Empresa",
    ]
    for field in key_fields:
        if field in df.columns:
            if df[field].isnull().any():
                null_count = df[field].isnull().sum()
                errors.append(f"Valores nulos en {field}: {null_count} de {len(df)}")
        else:
            errors.append(f"Campo clave esperado '{field}' no encontrado en el DataFrame.")

    # 4. Verificar tipos de datos (ejemplo)
    expected_types = {
        "totalDebit": "float64",
        "totalCredit": "float64",
        "balanceAtDateDebit": "float64",
        "balanceAtDateCredit": "float64",
    }
    if "number" in df.columns and pd.api.types.is_integer_dtype(df["number"]):
        expected_types["number"] = "Int64"  # o el tipo que se haya podido castear

    for col, dtype in expected_types.items():
        if col in df.columns:
            if str(df[col].dtype) != dtype:
                errors.append(
                    f"Tipo de dato incorrecto para '{col}'. Esperado: {dtype}, Actual: {df[col].dtype}"
                )
        # No añadir error si la columna no existe, ya se reportó en el punto 3

    return errors


def build_historical(start_date_str, end_date_str, empresas_a_procesar=None):
    """Construye dataset histórico completo de Trial Balance, extrayendo mes a mes. Solo retorna los datos nuevos."""
    all_data = []
    try:
        start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()
        end_date = datetime.strptime(end_date_str, "%Y-%m-%d").date()
    except ValueError:
        log.error("Error: Formato de fecha incorrecto. Use YYYY-MM-DD.")
        return None
    if start_date > end_date:
        log.error("Error: La fecha de inicio no puede ser posterior a la fecha de fin.")
        return None
    current_month_start = start_date.replace(day=1)
    empresas = empresas_a_procesar if empresas_a_procesar is not None else EMPRESAS
    while current_month_start <= end_date:
        year = current_month_start.year
        month = current_month_start.month
        last_day_of_month = current_month_start.replace(day=calendar.monthrange(year, month)[1])
        current_month_end = min(last_day_of_month, end_date)
        log.debug(
            f"Procesando mes: {current_month_start.strftime('%Y-%m')} ({current_month_start.strftime('%Y-%m-%d')} - {current_month_end.strftime('%Y-%m-%d')})"
        )
        for empresa_config in empresas:
            df_empresa_month = get_trial_balance_data(
                empresa_config,
                current_month_start.strftime("%Y-%m-%d"),
                current_month_end.strftime("%Y-%m-%d"),
            )
            if not df_empresa_month.empty:
                all_data.append(df_empresa_month)
            log.debug(
                f"Progreso para {empresa_config['display_name']}: {empresas.index(empresa_config) + 1}/{len(empresas)} empresas"
            )
        if current_month_end == end_date:
            break
        if current_month_start.month == 12:
            current_month_start = current_month_start.replace(
                year=current_month_start.year + 1, month=1, day=1
            )
        else:
            current_month_start = current_month_start.replace(
                month=current_month_start.month + 1, day=1
            )
    if not all_data:
        log.error(
            "No se encontraron datos de Trial Balance en el rango especificado para ninguna empresa o mes!"
        )
        return None
    full_data = pd.concat(all_data, ignore_index=True)
    log.debug("Validando datos finales...")
    validation_errors = validate_data(full_data, start_date_str, end_date_str)
    if validation_errors:
        log.warning("Advertencias de validación:")
        for error in validation_errors:
            log.warning(f"- {error}")
    else:
        log.debug("Datos validados sin errores aparentes.")
    return full_data


def read_historical_file():
    """Lee el archivo histórico, ya sea en CSV o XLSX, priorizando CSV si existe."""
    csv_exists = os.path.exists(CSV_PATH)
    xlsx_exists = os.path.exists(XLSX_PATH)
    if csv_exists:
        try:
            historical = pd.read_csv(CSV_PATH, sep=";", decimal=",")
            log.debug(f"Archivo histórico encontrado: {CSV_PATH} (CSV)")
            return historical, "csv"
        except Exception as e:
            log.error(f"Error al leer el archivo histórico CSV {CSV_PATH}: {e}.")
    if xlsx_exists:
        try:
            historical = pd.read_excel(XLSX_PATH)
            log.debug(f"Archivo histórico encontrado: {XLSX_PATH} (XLSX)")
            return historical, "xlsx"
        except Exception as e:
            log.error(f"Error al leer el archivo histórico XLSX {XLSX_PATH}: {e}.")
    log.debug(f"No se encontró archivo histórico {CSV_PATH} ni {XLSX_PATH}. Se creará uno nuevo.")
    return pd.DataFrame(), None


def update_current_month():
    """Actualiza solo con datos del mes pasado para Trial Balance"""
    today = datetime.now()
    first_day_this_month = today.replace(day=1)
    last_day_last_month = first_day_this_month - timedelta(days=1)
    first_day_last_month = last_day_last_month.replace(day=1)

    start_date_str = first_day_last_month.strftime("%Y-%m-%d")
    end_date_str = last_day_last_month.strftime("%Y-%m-%d")

    log.info(
        f"[{start_date_str} - {end_date_str}] Extrayendo datos de Trial Balance del mes pasado"
    )

    # Obtener nuevos datos
    new_data = build_historical(start_date_str, end_date_str)

    if new_data is not None and not new_data.empty:
        log.info(
            f"[{start_date_str} - {end_date_str}] Actualización mensual de Trial Balance completada"
        )
        # Eliminar y reemplazar solo el rango/empresas solicitados
        merge_and_replace_historical(new_data, start_date_str, end_date_str)
        return new_data  # Retornamos los datos generados para este periodo.
    else:
        log.info(
            f"[{start_date_str} - {end_date_str}] No se encontraron datos nuevos de Trial Balance."
        )
        return None


def main_menu():
    """Menú interactivo"""
    print("\n" + "=" * 60)
    print("SISTEMA DE EXTRACCIÓN DE TRIAL BALANCE - BUSINESS CENTRAL")
    print("=" * 60)
    print("\n1. Extraer Trial Balance (rango personalizado)")
    print("2. Extraer Trial Balance del mes pasado")
    print("3. Validar archivo histórico existente (trial_balance_historico.csv/xlsx)")
    print("4. Salir")

    choice = input("\nSeleccione opción (1-4): ")

    if choice == "1":
        print("\nEXTRACCIÓN DE TRIAL BALANCE PERSONALIZADA")
        print("Ejemplo de fecha: 2023-01-01")
        start_date = input(
            "Fecha inicio (YYYY-MM-DD): "
        )  # El usuario ingresa, por ejemplo, 2024-04-01
        end_date = input("Fecha fin (YYYY-MM-DD): ")  # El usuario ingresa, por ejemplo, 2024-06-30

        # Mostrar listado numerado de empresas
        print("\nSeleccione las empresas a procesar:")
        for idx, emp in enumerate(EMPRESAS, 1):
            print(f"{idx}. {emp['display_name']}")
        print("0. Todas las empresas")
        seleccion = input(
            "Ingrese el/los número(s) separados por coma (ej: 2,5,7) o 0 para todas: "
        ).strip()

        if seleccion == "0":
            empresas_seleccionadas = EMPRESAS
        else:
            try:
                indices = [int(x.strip()) for x in seleccion.split(",") if x.strip().isdigit()]
                empresas_seleccionadas = [
                    EMPRESAS[i - 1] for i in indices if 1 <= i <= len(EMPRESAS)
                ]
                if not empresas_seleccionadas:
                    print("Selección vacía o inválida. Se procesarán todas las empresas.")
                    empresas_seleccionadas = EMPRESAS
            except Exception as e:
                print(f"Error en la selección: {e}. Se procesarán todas las empresas.")
                empresas_seleccionadas = EMPRESAS

        new_data = build_historical(start_date, end_date, empresas_seleccionadas)
        if new_data is not None and not new_data.empty:
            merge_and_replace_historical(new_data, start_date, end_date, empresas_seleccionadas)

    elif choice == "2":
        update_current_month()

    elif choice == "3":
        # Validar ambos históricos si existen
        historical_csv = None
        historical_xlsx = None
        if os.path.exists(CSV_PATH):
            try:
                historical_csv = pd.read_csv(CSV_PATH, sep=";", decimal=",")
                print(f"\nVALIDANDO DATOS DEL ARCHIVO: {CSV_PATH}")
                errors = validate_data(historical_csv, "Desconocido", "Desconocido")
                if errors:
                    print("\nErrores/Advertencias encontrados en CSV:")
                    for error in errors:
                        print(f"- {error}")
                else:
                    print("\n✓ Datos validados sin errores aparentes en CSV.")
                print(
                    f"\nResumen del archivo CSV ({len(historical_csv)} filas):\n{historical_csv.info()}"
                )
                print(historical_csv.head().to_string())
            except Exception as e:
                print(f"Error al leer o validar el archivo CSV: {e}")

        if os.path.exists(XLSX_PATH):
            try:
                historical_xlsx = pd.read_excel(XLSX_PATH)
                print(f"\nVALIDANDO DATOS DEL ARCHIVO: {XLSX_PATH}")
                errors = validate_data(historical_xlsx, "Desconocido", "Desconocido")
                if errors:
                    print("\nErrores/Advertencias encontrados en XLSX:")
                    for error in errors:
                        print(f"- {error}")
                else:
                    print("\n✓ Datos validados sin errores aparentes en XLSX.")
                print(
                    f"\nResumen del archivo XLSX ({len(historical_xlsx)} filas):\n{historical_xlsx.info()}"
                )
                print(historical_xlsx.head().to_string())
            except Exception as e:
                print(f"Error al leer o validar el archivo XLSX: {e}")
        if not os.path.exists(CSV_PATH) and not os.path.exists(XLSX_PATH):
            print(f"\nNo existe archivo histórico para validar.")

    elif choice == "4":
        print("\nSaliendo del sistema...")
        sys.exit()

    else:
        print("\n¡Opción no válida!")


# if __name__ == "__main__":
#     sync_xlsx_with_csv()
#     print(
#         f"Script de extracción de Trial Balance iniciado - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
#     )

#     # Prueba de conexión inicial
#     print("\nProbando conexión con Business Central...")
#     bc_token = get_bc_access_token()

#     if not bc_token:
#         print(
#             "\nError: No se pudo obtener el token de acceso para Business Central. Verifica las credenciales y la configuración de la App Registration en Azure AD."
#         )
#         sys.exit(1)
#     else:
#         print("✓ Conexión exitosa y token obtenido.")

#     # Ejecución automática si se pasa --auto-update
#     if len(sys.argv) > 1 and sys.argv[1] == "--auto-update":
#         print("\nEjecutando actualización automática del mes pasado...")
#         update_current_month()
#         print("\nProceso de actualización automática finalizado.")
#     else:
#         while True:
#             main_menu()
#             input("\nPresione Enter para continuar...")

if __name__ == "__main__":
    log.info("--- TRIAL BALANCES START ---")

    update_current_month()

    sync_sharepoint(CSV_PATH)
    sync_sharepoint(XLSX_PATH)

    log.info("--- TRIAL BALANCES END ---")
