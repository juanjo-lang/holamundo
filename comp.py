import win32com.client
import subprocess
import time
import pandas as pd
import numpy as np
import datetime
import os


# --- FUNCIONES AUXILIARES ---


def wait_for_app(app_name, timeout=120):
    shell = win32com.client.Dispatch("WScript.Shell")
    waited = 0
    while not shell.AppActivate(app_name):
        time.sleep(5)
        waited += 5
        if waited >= timeout:
            raise TimeoutError(
                f"No se pudo activar {app_name} después de {timeout} segundos."
            )
    return shell


def get_sap_month(date_obj):
    m = date_obj.month
    sap_m = m - 3 if m >= 4 else m + 9
    return sap_m


# --- ABRIR SAP LOGON Y CONECTAR SESIÓN ---

subprocess.Popen([r"C:\Program Files (x86)\SAP\FrontEnd\SapGUI\saplogon.exe"])
shell = wait_for_app("SAP Logon")  # Espera a que SAP Logon esté activo

SapGui = win32com.client.GetObject("SAPGUI")
if not SapGui:
    raise Exception(
        "No se pudo obtener objeto SAPGUI. Verifica que SAP está abierto y con scripting activo."
    )

Appl = SapGui.GetScriptingEngine
Connection = Appl.OpenConnection(
    "2.06 - SAP PRD - Aliconsumo - Vitapro Chile - HCM Col", True
)
session = Connection.Children(0)

# --- PREPARAR FECHAS ---

hoy = datetime.datetime.now()
fecha_hoy = hoy.strftime("%d.%m.%Y")
mes_sap = get_sap_month(hoy)

# --- INICIO ---


def main():
    # Define configurable paths
    BASE_PATH = os.getenv('SAP_COMP_PATH', r"D:\Revisar juan\Bolivia\PythonOtros\CompSap")
    PREREGISTRO_FILE = os.path.join(BASE_PATH, "PREREGISTRO.xlsx")
    FACTURA_FILE = os.path.join(BASE_PATH, "FACTURA.xlsx")
    PEDIDO_FILE = os.path.join(BASE_PATH, "PEDIDO.xlsx")
    BANCO_FILE = os.path.join(BASE_PATH, "BANCO.xlsx")
    
    # Validate file existence
    if not os.path.exists(PREREGISTRO_FILE):
        raise FileNotFoundError(f"Required file not found: {PREREGISTRO_FILE}")
    
    try:
        df_facturas = pd.read_excel(PREREGISTRO_FILE)
        df_facturas = df_facturas.iloc[:-2]
        
        # More efficient DataFrame filtering using boolean indexing
        factura_mask = df_facturas["Factura"].notna() & (df_facturas["Factura"] != "")
        df_factura = df_facturas[factura_mask].copy()  # Use copy() to avoid SettingWithCopyWarning
        df_pedido = df_facturas[~factura_mask].copy()  # Use negation of mask for efficiency
        
        # Write DataFrames to Excel files
        df_factura.to_excel(FACTURA_FILE, index=False)
        df_pedido.to_excel(PEDIDO_FILE, index=False)
    except Exception as e:
        raise Exception(f"Error processing PREREGISTRO file: {str(e)}")

    try:
        df_banco = pd.read_excel(BANCO_FILE, header=4)
        df_factura = pd.read_excel(FACTURA_FILE,
            dtype={
                "Cliente": str,
                # "Número Pedido": str,
                # "Factura": str,
            },
        )
        df_pedido = pd.read_excel(PEDIDO_FILE,
            dtype={
                "Cliente": str,
                "Territorio": str,
                "Número Pedido": str,
                # "Factura": str,
            },
        )
    except Exception as e:
        raise Exception(f"Error reading Excel files: {str(e)}")

    # --- OBTENER FECHA Y TERRITORIO ---
    try:
        fecha = str(df_banco.iloc[0]["FECHA"])
        redondeo = float(df_banco.iloc[0]["REDONDEO"]) if pd.notna(df_banco.iloc[0]["REDONDEO"]) else 0.0
        # Fix: Add validation for territorio conversion
        territorio_raw = df_factura.iloc[0]["Territorio"]
        if pd.isna(territorio_raw):
            raise ValueError("Territorio value is missing")
        territorio = str(int(float(territorio_raw)))
    except (ValueError, IndexError) as e:
        raise Exception(f"Error extracting data from Excel files: {str(e)}")

    # --- LEER USUARIO Y CONTRASEÑA ---
    excel_app = None
    try:
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False  # Hide Excel application
        workbook = excel_app.Workbooks.Open(BANCO_FILE)
        sheet = workbook.Sheets("Hoja")
        user = sheet.Range("C2").Value
        key = sheet.Range("C3").Value
        
        # Validate credentials
        if not user or not key:
            raise ValueError("Username or password is empty")
        
        workbook.Close(SaveChanges=False)
    except Exception as e:
        raise Exception(f"Error reading credentials: {str(e)}")
    finally:
        if excel_app:
            excel_app.Quit()

    # --- LOGIN AUTOMÁTICO EN SAP ---
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = key
    session.findById("wnd[0]").sendVKey(0)

    # --- CABECERA DE LA OPERACIÓN ---
    session.findById("wnd[0]/tbar[0]/okcd").text = "FB05"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtBKPF-BLDAT").text = fecha
    session.findById("wnd[0]/usr/ctxtBKPF-BLART").text = "DB"
    session.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = "246"
    session.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text = fecha_hoy
    session.findById("wnd[0]/usr/txtBKPF-MONAT").text = mes_sap
    session.findById("wnd[0]/usr/ctxtBKPF-WAERS").text = "BOB"
    session.findById("wnd[0]/usr/txtBKPF-XBLNR").text = territorio
    session.findById("wnd[0]/usr/txtBKPF-BKTXT").text = f"COBRANZAS DUAL-{territorio}"

    # --- PARTIDAS DE BANCO ---
    last_nro_operacion = ""  # Initialize to handle case where no records exist
    for _, fila in df_banco.iterrows():
        banco = str(fila["CUENTA"])
        importes = str(fila["IMPORTE"])
        nro_operacion = str(fila["NRO.OPERACION"])
        fecha_banco = str(fila["FECHA"])  # Don't overwrite the main fecha variable
        last_nro_operacion = nro_operacion  # Keep track of last operation number
        
        session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "40"
        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = banco
        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").setFocus()
        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").caretPosition = len(banco)
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = importes
        session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = nro_operacion

    # Only set focus if we have at least one record
    if last_nro_operacion:
        session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").setFocus()
        session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").caretPosition = len(last_nro_operacion)
    # session.findById("wnd[0]/tbar[1]/btn[14]").press()

    # --- PARTIDAS DE PEDIDO ---
    if not df_pedido.empty:
        for _, fila in df_pedido.iterrows():
            cliente = str(fila["Cliente"])
            importe = str(fila["Importe Abonado"])
            territorio = str(fila["Territorio"])
            pedido = str(fila["Número Pedido"])

            # session.findById("wnd[0]").sendVKey(0)
            # session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "19"
            session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = cliente
            session.findById("wnd[0]/usr/ctxtRF05A-NEWUM").text = "X"
            session.findById("wnd[0]/usr/ctxtRF05A-NEWUM").setFocus()
            session.findById("wnd[0]/usr/ctxtRF05A-NEWUM").caretPosition = 1
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = importe
            session.findById("wnd[0]/usr/ctxtBSEG-ZFBDT").text = fecha_hoy
            session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = territorio
            session.findById("wnd[0]/usr/ctxtBSEG-VBEL2").text = pedido
            session.findById("wnd[0]/usr/ctxtBSEG-POSN2").text = "000010"
            session.findById("wnd[0]/usr/ctxtBSEG-ETEN2").text = "0001"

    # --- PARTIDAS DE FACTURA ---
    if not df_factura.empty:
        for i, (_, fila) in enumerate(df_factura.iterrows()):
            cliente = str(fila["Cliente"])
            factura = str(fila["Factura"])

            if i == 0:
                if not df_pedido.empty:
                    session.findById("wnd[0]/tbar[1]/btn[16]").press()
                    session.findById("wnd[0]").sendVKey(0)
                else:
                    session.findById("wnd[0]/tbar[1]/btn[16]").press()

            session.findById(
                "wnd[0]/usr/sub:SAPMF05A:0710/radRF05A-XPOS1[3,0]"
            ).select()
            session.findById("wnd[0]/usr/ctxtRF05A-AGBUK").text = "246"
            session.findById("wnd[0]/usr/ctxtRF05A-AGKON").text = cliente
            session.findById("wnd[0]/usr/ctxtRF05A-AGKOA").text = "D"
            session.findById("wnd[0]/usr/ctxtRF05A-AGUMS").text = "AXE"
            session.findById(
                "wnd[0]/usr/sub:SAPMF05A:0710/radRF05A-XPOS1[3,0]"
            ).setFocus()
            session.findById("wnd[0]/tbar[1]/btn[16]").press()
            session.findById(
                "wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[0,0]"
            ).text = factura
            session.findById(
                "wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[0,0]"
            ).caretPosition = len(factura)
            session.findById("wnd[0]/tbar[1]/btn[7]").press()

    session.findById("wnd[0]/tbar[0]/btn[12]").press()
    session.findById("wnd[0]/tbar[1]/btn[14]").press()

    # ---REDONDEO DE LA OPERACIÓN---
    # Simplified and more efficient redondeo logic
    if abs(redondeo) > 0.01:  # Use a small threshold to avoid floating point precision issues
        valor = abs(redondeo)
        valor_str = f"{valor:.2f}"  # More efficient string formatting
        
        # Use positive value for credit (50) and negative for debit (40)
        newbs_code = "50" if redondeo > 0 else "40"
        session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = newbs_code
        
        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = "659310999"
        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").setFocus()
        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").caretPosition = 9
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = valor_str
        session.findById("wnd[0]/usr/txtBSEG-WRBTR").caretPosition = len(valor_str)
        session.findById("wnd[0]/tbar[1]/btn[14]").press()

    session.findById("wnd[0]/mbar/menu[0]/menu[3]").select()
    session.findById("wnd[0]").sendVKey(0)


if __name__ == "__main__":
    main()
