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
    df_facturas = pd.read_excel(
        r"D:\Revisar juan\Bolivia\PythonOtros\CompSap\PREREGISTRO.xlsx"
    )
    df_facturas = df_facturas.iloc[:-2]
    df_factura = df_facturas[
        df_facturas["Factura"].notnull() & (df_facturas["Factura"] != "")
    ]
    df_factura.to_excel(
        r"D:\Revisar juan\Bolivia\PythonOtros\CompSap\FACTURA.xlsx", index=False
    )
    df_pedido = df_facturas[
        df_facturas["Factura"].isnull() | (df_facturas["Factura"] == "")
    ]
    df_pedido.to_excel(
        r"D:\Revisar juan\Bolivia\PythonOtros\CompSap\PEDIDO.xlsx", index=False
    )

    df_banco = pd.read_excel(
        r"D:\Revisar juan\Bolivia\PythonOtros\CompSap\BANCO.xlsx", header=4
    )
    df_factura = pd.read_excel(
        r"D:\Revisar juan\Bolivia\PythonOtros\CompSap\FACTURA.xlsx",
        dtype={
            "Cliente": str,
            # "Número Pedido": str,
            # "Factura": str,
        },
    )
    df_pedido = pd.read_excel(
        r"D:\Revisar juan\Bolivia\PythonOtros\CompSap\PEDIDO.xlsx",
        dtype={
            "Cliente": str,
            "Territorio": str,
            "Número Pedido": str,
            # "Factura": str,
        },
    )

    # --- OBTENER FECHA Y TERRITORIO ---
    fecha = str(df_banco.iloc[0]["FECHA"])
    redondeo = float(df_banco.iloc[0]["REDONDEO"])
    territorio = str(int(df_factura.iloc[0]["Territorio"]))

    # --- LEER USUARIO Y CONTRASEÑA ---
    excel_app = win32com.client.Dispatch("Excel.Application")
    workbook = excel_app.Workbooks.Open(
        r"D:\Revisar juan\Bolivia\PythonOtros\CompSap\BANCO.xlsx"
    )
    sheet = workbook.Sheets("Hoja")
    user = sheet.Range("C2").Value
    key = sheet.Range("C3").Value
    workbook.Close(SaveChanges=False)

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
