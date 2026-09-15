import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent / ".env")

# Business Central
BC_CLIENT_ID = os.getenv("BC_CLIENT_ID")
BC_CLIENT_SECRET = os.getenv("BC_CLIENT_SECRET")
BC_TENANT_ID = os.getenv("BC_TENANT_ID")
BC_AUTHORITY = f"https://login.microsoftonline.com/{BC_TENANT_ID}"
BC_SCOPE = ["https://api.businesscentral.dynamics.com/.default"]
BC_ENVIRONMENT = "Production"

# Microsoft Graph
APP_TENANT_ID = os.getenv("APP_TENANT_ID")
APP_CLIENT_ID = os.getenv("APP_CLIENT_ID")
APP_CLIENT_SECRET = os.getenv("APP_CLIENT_SECRET")
APP_AUTHORITY = f"https://login.microsoftonline.com/{APP_TENANT_ID}"
APP_SCOPE = ["https://graph.microsoft.com/.default"]

# Sharepoint
SHAREPOINT_BASE_URL = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOST = os.getenv("SHAREPOINT_HOST")
SHAREPOINT_SITE_NAME = os.getenv("SHAREPOINT_SITE_NAME")
SHAREPOINT_DRIVE_NAME = os.getenv("SHAREPOINT_DRIVE_NAME")
SHAREPOINT_FOLDER_PATH = os.getenv("SHAREPOINT_FOLDER_PATH")


EMPRESAS = [
    {
        "company_name": "IMO",
        "api_name": "IMO",
        "display_name": "IMO",
        "sociedad": "INSTITUTO DE MICROCIRUGIA OCULAR  DOS S.LU.",
        "company_id": "f544e512-20ea-eb11-a1de-000d3a4b628d",
    },
    {
        "company_name": "VALCASADO",
        "api_name": "VALCASADO",
        "display_name": "Valcasado",
        "sociedad": "Valcasado, SAU",
        "company_id": "19d6fe81-9532-ec11-a459-6045bd8b79a4",
    },
    {
        "company_name": "ANDORRA",
        "api_name": "ANDORRA",
        "display_name": "Andorra",
        "sociedad": "Miranza Pirineos S.A.U.",
        "company_id": "c889dcf6-4e6d-ec11-bf26-000d3abfaaef",
    },
    {
        "company_name": "BACKUP_SGA",
        "api_name": "BACKUP_SGA",
        "display_name": "BACKUP_SGA",
        "sociedad": "BACKUP_SGA",
        "company_id": "f4ec3bc8-c794-ec11-80f2-002248833581",
    },
    {
        "company_name": "Begitek%20Clinica%20Oftalmologica%2C",
        "api_name": "Begitek Clinica Oftalmologica,",
        "display_name": "Begitek Clinica Oftalmologica,",
        "sociedad": "Begitek Clinica Oftalmologica,",
        "company_id": "1f8b6cc7-2e21-ec11-8f45-6045bd8b7aa6",
    },
    {
        "company_name": "BRILLANTE%20-%20C%C3%93RDOBA",
        "api_name": "BRILLANTE - CÓRDOBA",
        "display_name": "Cordoba",
        "sociedad": "Clínica El Brillante Oftamología S.L.U.",
        "company_id": "be6b25b5-626b-ef11-a673-000d3ab799a0",
    },
    {
        "company_name": "CAMACHO%20VINUESA",
        "api_name": "CAMACHO VINUESA",
        "display_name": "CAMACHO VINUESA",
        "sociedad": "Centro Camacho Vinuesa, SL",
        "company_id": "0a6ec322-c681-ec11-b85a-000d3a27b8eb",
    },
    {
        "company_name": "CAPEVISION",
        "api_name": "CAPEVISION",
        "display_name": "CAPEVISION",
        "sociedad": "Capevision, SL",
        "company_id": "7aa7e268-d081-ec11-b85a-000d3a27b8eb",
    },
    {
        "company_name": "CENTRO%20OFTALMOLOGICO%20ALTEA%2C",
        "api_name": "CENTRO OFTALMOLOGICO ALTEA,",
        "display_name": "Palomares",
        "sociedad": "Centro Oftalmologico Altea, S.L.U.",
        "company_id": "d7aa16c1-ea21-ec11-8f45-6045bd8b7085",
    },
    {
        "company_name": "COB",
        "api_name": "COB",
        "display_name": "IMO Manresa",
        "sociedad": "Institut Oftalmologic Laser Visio, SLU",
        "company_id": "edcc537a-a12e-ec11-8f45-6045bd8b7218",
    },
    {
        "company_name": "COI%20BILBAO",
        "api_name": "COI BILBAO",
        "display_name": "COI",
        "sociedad": "Centro Ofalmologico Integral Bilbao Berri, S.L.U.",
        "company_id": "20056cf4-832a-ec11-8f45-6045bd8b79a4",
    },
    {
        "company_name": "CRM",
        "api_name": "CRM",
        "display_name": "CRM",
        "sociedad": "Centro de Reconocimiento Medico Moratalaz, SLU",
        "company_id": "6bd8fe50-d221-ec11-8f45-6045bd8b7085",
    },
    {
        "company_name": "CRONUS%20ES",
        "api_name": "CRONUS ES",
        "display_name": "CRONUS ES",
        "sociedad": "CRONUS ES",
        "company_id": "b9bff1e2-11a8-eb11-bb64-000d3a2e849d",
    },
    {
        "company_name": "CVL",
        "api_name": "CVL",
        "display_name": "CVL",
        "sociedad": "Clinica Virgen de Lujan, SLU",
        "company_id": "d9d06667-a32d-ec11-8f45-6045bd8b7085",
    },
    {
        "company_name": "FUNDACION%20MIRANZA",
        "api_name": "FUNDACION MIRANZA",
        "display_name": "Fundacion Miranza",
        "sociedad": "Fundacion Miranza",
        "company_id": "5c45611e-82e0-ee11-904c-000d3ade8d55",
    },
    {
        "company_name": "IBO",
        "api_name": "IBO",
        "display_name": "IBO",
        "sociedad": "Instituto Balear Oftalmologico, SLU",
        "company_id": "b2a6f628-2730-ec11-8f45-000d3a39ebc2",
    },
    {
        "company_name": "UOB",
        "api_name": "UOB",
        "display_name": "UOB",
        "sociedad": "Unidad Oftalmologica Balear, SLU",
        "company_id": "",
    },
    {
        "company_name": "IOA",
        "api_name": "IOA",
        "display_name": "IOA",
        "sociedad": "Inverlasik Mad II, SLU",
        "company_id": "6125866e-e530-ec11-8f45-6045bd8b7085",
    },
    # {"company_name": "IRADER","api_name": "IRADIER", "display_name": "IRADIER", "sociedad": "Clinica Oftalmologica Iradier, SLU", "company_id": "1046880e-012d-ec11-8f45-6045bd8b7218"},
    {
        "company_name": "Miranza%20Begoña%2C%20SLU",
        "api_name": "Miranza Begoña, SLU",
        "display_name": "Begoña",
        "sociedad": "Miranza Begoña, SL",
        "company_id": "efdcdb37-a917-ec11-86bc-00224881d396",
    },
    {
        "company_name": "MIRANZA%20GALICIA%2C%20S.L.",
        "api_name": "MIRANZA GALICIA, S.L.",
        "display_name": "Galicia",
        "sociedad": "Miranza Galicia S.L.",
        "company_id": "5ac3ed5b-5c61-ee11-8df1-000d3a69f85e",
    },
    {
        "company_name": "MIRANZA%20GRAN%20CANARIA%2C%20S.L.",
        "api_name": "MIRANZA GRAN CANARIA, S.L.",
        "display_name": "Gran Canaria",
        "sociedad": "Miranza Gran Canaria S.L.",
        "company_id": "a64cc2f0-5c61-ee11-8df1-000d3a69f85e",
    },
    {
        "company_name": "MIRANZA%20INV%20OFT",
        "api_name": "MIRANZA INV OFT",
        "display_name": "MIO",
        "sociedad": "Miranza Inversiones Oftalmologicas, SL",
        "company_id": "3a6f998f-4c33-ec11-a459-6045bd8b7d50",
    },
    {
        "company_name": "MIRANZA%20MÁLAGA%20CORRECTO",
        "api_name": "MIRANZA MÁLAGA CORRECTO",
        "display_name": "Malaga",
        "sociedad": "Miranza Málaga",
        "company_id": "8a364e75-b215-ee11-8f6e-0022489b795b",
    },
    {
        "company_name": "MIRANZA%20QX",
        "api_name": "MIRANZA QX",
        "display_name": "Miranza QX",
        "sociedad": "Miranza QX, SLU",
        "company_id": "a8324736-8f31-ec11-8f45-6045bd8b7ace",
    },
    {
        "company_name": "Miranza%20Santander%20S.L.U.",
        "api_name": "Miranza Santander S.L.U.",
        "display_name": "Santander",
        "sociedad": "Miranza Santander S.L.U.",
        "company_id": "1ff0a388-10d7-ef11-b8ec-000d3ab0bb3b",
    },
    {
        "company_name": "MIRANZA%20SEVILLA",
        "api_name": "MIRANZA SEVILLA",
        "display_name": "CIMO",
        "sociedad": "Miranza Sevilla S.L.U",
        "company_id": "87c88a0b-5af3-ef11-9345-7c1e527652b5",
    },
    {
        "company_name": "MIRANZA%20TENERIFE",
        "api_name": "MIRANZA TENERIFE",
        "display_name": "Tenerife",
        "sociedad": "Miranza Tenerife, SLU",
        "company_id": "029ab625-9231-ec11-8f45-6045bd8b7ace",
    },
    {
        "company_name": "OFTALMED",
        "api_name": "OFTALMED",
        "display_name": "Oculsur",
        "sociedad": "Oftalmed CCT, SLU",
        "company_id": "74e946f2-ba31-ec11-8f45-6045bd8b79a4",
    },
    {
        "company_name": "TILICE%20EUROPE%20S.L.",
        "api_name": "TILICE",
        "display_name": "TILICE",
        "sociedad": "Tilice Europe S.L.",
        "company_id": "238f87ce-ed90-ed11-bff5-002248a164c4",
    },
    {
        "company_name": "OT20",
        "api_name": "Ophthalteam",
        "display_name": "Ophthalteam",
        "sociedad": "Ophthalteam 2020, SLU",
        "company_id": "8c70e35a-3d32-ec11-a459-6045bd8b7085",
    },
    {
        "company_name": "VIRGEN%20DE%20LA%20PALMA",
        "api_name": "VIRGEN DE LA PALMA",
        "display_name": "Algeciras",
        "sociedad": "Virgen de La Palma, S.L.U.",
        "company_id": "74c07d70-df5f-ec11-9f08-000d3abfaf5f",
    },
    {
        "company_name": "OKULAR",
        "api_name": "OKULAR",
        "display_name": "OKULAR",
        "sociedad": "Clinica Oftalmologica Gasteiz, S.L.U.",
        "company_id": "666ede81-0322-ec11-8f45-6045bd8b7085",
    },
    {
        "company_name": "TECNOLASER%20SANTA%20JUSTA%20S.L.U.",
        "api_name": "TECNOLASER SANTA JUSTA S.L.U.",
        "display_name": "Seville -Tecnolaser",
        "sociedad": "TECNOLASER SANTA JUSTA S.L.U.",
        "company_id": "666ede81-0322-ec11-8f45-6045bd8b7085",
    },
    {
        "company_name": "CLINICA%20DR.%20MONTEIRO",
        "api_name": "Clínica Oftalmológica Prof. Doutor Manuel Monteiro, LDA",
        "display_name": "CLINICA DR. MONTEIRO",
        "sociedad": "Clínica Oftalmológica Prof. Doutor Manuel Monteiro, LDA",
        "company_id": "03794791-9c9e-f011-b41a-6045bdde20d6"
    },
    {
        "company_name": "ROIJEN%20NACAR%20OFTALMOLOGIA%2C%20S.L",
        "api_name": "Roijen Nacar Oftalmología, S.L ",
        "display_name": "ROIJEN NACAR OFTALMOLOGIA, S.L",
        "sociedad": "Roijen Nacar Oftalmología, S.L ",
        "company_id": "43af1950-84b5-f011-bbd1-7c1e5235d486"
    },
    {
        "company_name": "MIRANZA%20PORTUGAL%2C%20LDA",
        "api_name": "Miranza Portugal, LDA",
        "display_name": "MIRANZA PORTUGAL, LDA",
        "sociedad": "Miranza Portugal, LDA",
        "company_id": "91ac4605-19a5-f011-bbd0-7ced8d4957bf"
    },
    {
        "company_name": "INSTITUT%20OFTAMOLOGIC%20DEL%20PRAT",
        "api_name": "Institut Oftalmologic del Prat, S.L",
        "display_name": "INSTITUT OFTAMOLOGIC DEL PRAT",
        "sociedad": "Institut Oftalmologic del Prat, S.L",
        "company_id": "471002a2-84b5-f011-bbd1-7c1e5235d486"
    },
    {
        "company_name": "OFTALMOCENTER",
        "api_name": "OFTALMOCENTER - CLINICA MEDICA, LDA",
        "display_name": "OFTALMOCENTER",
        "sociedad": "OFTALMOCENTER - CLINICA MEDICA, LDA",
        "company_id": "0a42708c-87f0-f011-8405-7ced8d77d6a6"
    },
    {
        "company_name": "CLINICA%20OFTALMOL%C3%93GICA%20PRIVADA",
        "api_name": "CLINICA PRIVADA OFTALMOLÓGICA",
        "display_name": "CLINICA OFTALMOLÓGICA PRIVADA",
        "sociedad": "CLINICA PRIVADA OFTALMOLÓGICA",
        "company_id": "e525a312-1790-f111-8074-7c1e5276275c"
    },
]
