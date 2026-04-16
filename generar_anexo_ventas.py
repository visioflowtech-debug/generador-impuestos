import json
import os
import csv
import datetime
import openpyxl

# --- Configuración ---
JSON_DIR = "C:/phytonmailjson/ventas"
OUTPUT_DIR = "C:/phytonmailjson"

def get_output_filenames():
    """Genera nombres de archivos CSV y XLSX con la fecha actual."""
    today = datetime.date.today().strftime("%Y-%m-%d")
    csv_file = os.path.join(OUTPUT_DIR, f"Anexo_Ventas_{today}.csv")
    xlsx_file = os.path.join(OUTPUT_DIR, f"Anexo_Ventas_{today}.xlsx")
    return csv_file, xlsx_file

def format_date(date_str):
    """Convierte fecha YYYY-MM-DD a DD/MM/AAAA."""
    if not date_str: return ""
    try:
        return datetime.datetime.strptime(date_str, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        try:
            return datetime.datetime.fromisoformat(date_str).strftime("%d/%m/%Y")
        except ValueError:
            return date_str

def format_float(value):
    """Formatea flotante a string con 2 decimales."""
    try:
        return "{:.2f}".format(float(value))
    except (ValueError, TypeError):
        return "0.00"

def main():
    if not os.path.exists(JSON_DIR):
        print(f"Error: No existe el directorio {JSON_DIR}")
        return

    csv_file, xlsx_file = get_output_filenames()
    
    # Columnas del Anexo Ventas
    headers = [
        "A FECHA DE EMISIÓN", "B CLASE DE DOCUMENTO", "C TIPO DE DOCUMENTO", "D NÚMERO DE RESOLUCIÓN",
        "E SERIE DE DOCUMENTO", "F NÚMERO DE CONTROL INTERNO (DEL)", "G NÚMERO DE CONTROL INTERNO (AL)",
        "H NÚMERO DE DOCUMENTO (DEL)", "I NÚMERO DE DOCUMENTO (AL)", "J N° DE MAQUINA REGISTRADORA",
        "K VENTAS EXENTAS", "L VENTAS INTERNAS EXENTAS NO SUJETAS A PROPORCIONALIDAD", "M VENTAS NO SUJETAS",
        "N VENTAS GRAVADAS LOCALES", "O EXPORTACIONES DENTRO DEL ÁREA CENTROAMERICANA", 
        "P EXPORTACIONES FUERA DEL ÁREA CENTROAMERICANA", "Q EXPORTACIONES DE SERVICIOS",
        "R VENTAS A ZONAS FRANCAS Y DPA (TASA CERO)", "S VENTAS A CUENTA DE TERCEROS NO DOMICILADOS",
        "T TOTAL VENTAS", "U TIPO DE OPERACIÓN (Renta)", "V TIPO DE INGRESO (Renta)", "W NÚMERO DE ANEXO"
    ]
    
    csv_rows = []
    count_processed = 0

    try:
        files = [f for f in os.listdir(JSON_DIR) if f.lower().endswith(".json")]
    except OSError as e:
        print(f"Error al acceder a {JSON_DIR}: {e}")
        return

    print(f"Procesando {len(files)} archivos de Ventas en {JSON_DIR}...")
    
    # Iterar archivos
    for filename in files:
        filepath = os.path.join(JSON_DIR, filename)
        
        try:
            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except UnicodeDecodeError:
                with open(filepath, "r", encoding="latin-1") as f:
                    data = json.load(f)
            
            # Filtrar solo DTE 01 (Factura Consumidor Final)
            identificacion = data.get("identificacion", {})
            tipo_dte = identificacion.get("tipoDte")
            
            if tipo_dte != "01":
                continue
            
            # Fecha (YYYY-MM-DD)
            fecha_emi = identificacion.get("fecEmi")
            if not fecha_emi:
                continue

            codigo_gen = identificacion.get("codigoGeneracion", "").replace("-", "")
            
            resumen = data.get("resumen", {})
            gravada = float(resumen.get("totalGravada", 0))
            exenta = float(resumen.get("totalExenta", 0))
            nosujeta = float(resumen.get("totalNoSuj", 0))
            pagar = float(resumen.get("totalPagar", 0))
            
            # Crear fila para este DTE individual
            # A: Fecha
            col_a = format_date(fecha_emi)
            # N: Ventas Gravadas Locales
            col_n = format_float(gravada)
            # T: Total Ventas
            col_t = format_float(pagar)
            
            # Renta 2025 Defaults
            col_u = "1"
            col_v = "3"
            
            # Fila completa
            row = [
                col_a, "4", "01", "N/A", "N/A", "N/A", "N/A", 
                codigo_gen, codigo_gen, "", 
                format_float(exenta), "0.00", format_float(nosujeta),
                col_n, "0.00", "0.00", "0.00", "0.00", "0.00", 
                col_t, col_u, col_v, "2"
            ]
            csv_rows.append(row)
            count_processed += 1
            
        except json.JSONDecodeError:
            print(f"Error: {filename} no es un JSON válido.")
        except Exception as e:
            print(f"Error procesando {filename}: {e}")

    # Ordenar por fecha (asumiendo formato DD/MM/AAAA en col_a)
    if csv_rows:
        csv_rows.sort(key=lambda x: datetime.datetime.strptime(x[0], "%d/%m/%Y"))


    # Output Loop
    if csv_rows:
        # CSV
        try:
            with open(csv_file, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerows(csv_rows)
            print(f"Generado CSV: {csv_file}")
        except Exception as e:
            print(f"Error escribiendo CSV: {e}")

        # XLSX
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Anexo Ventas"
            ws.append(headers)
            for r in csv_rows:
                ws.append(r)
            wb.save(xlsx_file)
            print(f"Generado Excel: {xlsx_file}")
        except Exception as e:
             print(f"Error escribiendo Excel: {e}")
             
        print(f"Total DTEs de Ventas (01) procesados: {count_processed}")
    else:
        print("No se encontraron DTEs de Ventas (01) para reportar.")
    
    print("-" * 30)

if __name__ == "__main__":
    main()
