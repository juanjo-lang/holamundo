import win32com.client
import subprocess
import time
import pandas as pd
import numpy as np
import datetime
import os
import configparser
from pathlib import Path


# --- CONFIGURATION MANAGEMENT ---
def load_config():
    """Load configuration from config.ini file or environment variables"""
    config = configparser.ConfigParser()
    config_file = Path("config.ini")
    
    if config_file.exists():
        config.read(config_file)
    else:
        # Create default config structure
        config['PATHS'] = {
            'sap_logon': r"C:\Program Files (x86)\SAP\FrontEnd\SapGUI\saplogon.exe",
            'preregistro': r"D:\Revisar juan\Bolivia\PythonOtros\CompSap\PREREGISTRO.xlsx",
            'banco': r"D:\Revisar juan\Bolivia\PythonOtros\CompSap\BANCO.xlsx",
            'factura': r"D:\Revisar juan\Bolivia\PythonOtros\CompSap\FACTURA.xlsx",
            'pedido': r"D:\Revisar juan\Bolivia\PythonOtros\CompSap\PEDIDO.xlsx"
        }
        config['SAP'] = {
            'connection': "2.06 - SAP PRD - Aliconsumo - Vitapro Chile - HCM Col",
            'company_code': "246",
            'currency': "BOB"
        }
        config['CREDENTIALS'] = {
            'user_cell': "C2",
            'password_cell': "C3",
            'sheet_name': "Hoja"
        }
        
        # Save default config
        with open(config_file, 'w') as f:
            config.write(f)
    
    return config


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


def safe_read_excel(file_path, **kwargs):
    """Safely read Excel file with proper error handling"""
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Archivo no encontrado: {file_path}")
        return pd.read_excel(file_path, **kwargs)
    except Exception as e:
        raise Exception(f"Error al leer {file_path}: {str(e)}")


def get_credentials_from_excel(file_path, config):
    """Safely extract credentials from Excel file with proper resource management"""
    excel_app = None
    workbook = None
    try:
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False
        workbook = excel_app.Workbooks.Open(file_path)
        sheet = workbook.Sheets(config['CREDENTIALS']['sheet_name'])
        
        user = sheet.Range(config['CREDENTIALS']['user_cell']).Value
        key = sheet.Range(config['CREDENTIALS']['password_cell']).Value
        
        if not user or not key:
            raise ValueError("Credenciales no encontradas en el archivo Excel")
            
        return user, key
    except Exception as e:
        raise Exception(f"Error al obtener credenciales: {str(e)}")
    finally:
        if workbook:
            try:
                workbook.Close(SaveChanges=False)
            except:
                pass
        if excel_app:
            try:
                excel_app.Quit()
            except:
                pass


# --- ABRIR SAP LOGON Y CONECTAR SESIÓN ---
def initialize_sap(config):
    """Initialize SAP connection with proper error handling"""
    try:
        subprocess.Popen([config['PATHS']['sap_logon']])
        shell = wait_for_app("SAP Logon")
        
        SapGui = win32com.client.GetObject("SAPGUI")
        if not SapGui:
            raise Exception(
                "No se pudo obtener objeto SAPGUI. Verifica que SAP está abierto y con scripting activo."
            )
        
        Appl = SapGui.GetScriptingEngine
        Connection = Appl.OpenConnection(config['SAP']['connection'], True)
        session = Connection.Children(0)
        
        return session
    except Exception as e:
        raise Exception(f"Error al inicializar SAP: {str(e)}")


# --- PREPARAR FECHAS ---
def prepare_dates():
    """Prepare current date and SAP month"""
    hoy = datetime.datetime.now()
    fecha_hoy = hoy.strftime("%d.%m.%Y")
    mes_sap = get_sap_month(hoy)
    return fecha_hoy, mes_sap


# --- INICIO ---
def main():
    # Load configuration
    config = load_config()
    
    # Initialize SAP
    session = initialize_sap(config)
    
    # Prepare dates
    fecha_hoy, mes_sap = prepare_dates()
    
    # Read and process data files
    df_facturas = safe_read_excel(config['PATHS']['preregistro'])
    
    # Fix Bug 2: Correct DataFrame filtering logic
    # Remove last 2 rows only if they exist and are empty
    if len(df_facturas) > 2:
        # Check if last 2 rows are empty or contain only NaN values
        last_two_rows = df_facturas.tail(2)
        if last_two_rows.isna().all().all() or (last_two_rows == "").all().all():
            df_facturas = df_facturas.iloc[:-2]
    
    # Improved filtering logic
    df_factura = df_facturas[
        df_facturas["Factura"].notna() & 
        (df_facturas["Factura"].astype(str).str.strip() != "")
    ]
    
    df_pedido = df_facturas[
        df_facturas["Factura"].isna() | 
        (df_facturas["Factura"].astype(str).str.strip() == "")
    ]
    
    # Save filtered data
    df_factura.to_excel(config['PATHS']['factura'], index=False)
    df_pedido.to_excel(config['PATHS']['pedido'], index=False)
    
    # Read bank data
    df_banco = safe_read_excel(config['PATHS']['banco'], header=4)
    
    # Re-read processed data with proper data types
    df_factura = safe_read_excel(
        config['PATHS']['factura'],
        dtype={"Cliente": str}
    )
    
    df_pedido = safe_read_excel(
        config['PATHS']['pedido'],
        dtype={
            "Cliente": str,
            "Territorio": str,
            "Número Pedido": str,
        }
    )
    
    # Extract key data
    if df_banco.empty:
        raise ValueError("Archivo de banco está vacío")
    
    fecha = str(df_banco.iloc[0]["FECHA"])
    redondeo = float(df_banco.iloc[0]["REDONDEO"])
    
    if df_factura.empty:
        raise ValueError("No hay datos de factura para procesar")
    
    territorio = str(int(df_factura.iloc[0]["Territorio"]))
    
    # Get credentials securely
    user, key = get_credentials_from_excel(config['PATHS']['banco'], config)
    
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
    session.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = config['SAP']['company_code']
    session.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text = fecha_hoy
    session.findById("wnd[0]/usr/txtBKPF-MONAT").text = mes_sap
    session.findById("wnd[0]/usr/ctxtBKPF-WAERS").text = config['SAP']['currency']
    session.findById("wnd[0]/usr/txtBKPF-XBLNR").text = territorio
    session.findById("wnd[0]/usr/txtBKPF-BKTXT").text = f"COBRANZAS DUAL-{territorio}"
    
    # --- PARTIDAS DE BANCO ---
    for _, fila in df_banco.iterrows():
        banco = str(fila["CUENTA"])
        importes = str(fila["IMPORTE"])
        nro_operacion = str(fila["NRO.OPERACION"])
        fecha = str(fila["FECHA"])
        session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "40"
        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = banco
        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").setFocus()
        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").caretPosition = len(banco)
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = importes
        session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = nro_operacion
    
    session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").setFocus()
    session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").caretPosition = len(nro_operacion)
    
    # --- PARTIDAS DE PEDIDO ---
    if not df_pedido.empty:
        for _, fila in df_pedido.iterrows():
            cliente = str(fila["Cliente"])
            importe = str(fila["Importe Abonado"])
            territorio = str(fila["Territorio"])
            pedido = str(fila["Número Pedido"])
            
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
            session.findById("wnd[0]/usr/ctxtRF05A-AGBUK").text = config['SAP']['company_code']
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
    if not (redondeo == 0 or redondeo == "" or redondeo is None or np.isnan(redondeo)):
        valor = abs(redondeo)
        valor_str = "{:.2f}".format(valor)
        if redondeo > 0:
            session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "50"
        else:
            session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "40"
        
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
