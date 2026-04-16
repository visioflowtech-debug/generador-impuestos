import json
import os
import csv
import datetime
import openpyxl

# --- Configuración ---
JSON_DIR = "C:/output_dir/json"
OUTPUT_DIR = "C:/output_dir"

def get_output_filename():
    """Genera el nombre del archivo CSV con la fecha actual."""
    today = datetime.date.today().strftime("%Y-%m-%d")
    return os.path.join(OUTPUT_DIR, f"Anexo_Compras_{today}.csv")

def format_date(date_str):
    """Convierte fecha YYYY-MM-DD a DD/MM/AAAA."""
    if not date_str: return ""
    try:
        # Intentar formato estándar YYYY-MM-DD
        return datetime.datetime.strptime(date_str, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        try:
            # Intentar formato ISO con hora
            return datetime.datetime.fromisoformat(date_str).strftime("%d/%m/%Y")
        except ValueError:
            return date_str

def format_float(value):
    """Formatea flotante a string con 2 decimales."""
    try:
        return "{:.2f}".format(float(value))
    except (ValueError, TypeError):
        return "0.00"

def get_classification(json_data):
    """
    Determina Q, R, S, T basado en reglas de negocio.
    Retorna una tupla (Q, R, S, T).
    """
    # Extraer datos relevantes para las reglas
    emisor = json_data.get("emisor", {})
    emisor_nombre = emisor.get("nombre", "").upper() if emisor else ""
    
    # Buscar descripción en cuerpoDocumento
    descripcion_items = ""
    cuerpo = json_data.get("cuerpoDocumento", [])
    if isinstance(cuerpo, list):
        for item in cuerpo:
            desc = item.get("descripcion", "")
            if desc:
                descripcion_items += desc.upper() + " "
            
    # Regla 1: ALQUILER (Gasto)
    if "ALQUILER" in descripcion_items or "CHONSA" in emisor_nombre:
        return "1", "2", "2", "2"
        
    # Regla 2: COMISIONES (Gasto Financiero)
    # Corrección lógica: Separar condiciones
    if "COMISION" in descripcion_items or "SERFINSA" in emisor_nombre or "SERVICIOS FINANCIEROS" in emisor_nombre:
        return "1", "2", "2", "3"
        
    # Regla 3: COSTOS (Materiales/Default)
    return "1", "1", "2", "5"


def main():
    if not os.path.exists(JSON_DIR):
        print(f"Error: No existe el directorio {JSON_DIR}")
        return

    output_file = get_output_filename()
    
    # Columnas del Anexo
    headers = [
        "A FECHA DE EMISIÓN", "B CLASE DE DOCUMENTO", "C TIPO DE DOCUMENTO", "D NÚMERO DE DOCUMENTO",
        "E NIT O NRC DEL PROVEEDOR", "F NOMBRE DEL PROVEEDOR", "G COMPRAS INTERNAS EXENTAS",
        "H INTERNACIONES EXENTAS", "I IMPORTACIONES EXENTAS", "J COMPRAS INTERNAS GRAVADAS",
        "K INTERNACIONES GRAVADAS", "L IMPORTACIONES GRAVADAS BIENES", "M IMPORTACIONES GRAVADAS SERVICIOS",
        "N CRÉDITO FISCAL", "O TOTAL DE COMPRAS", "P DUI DEL PROVEEDOR", "Q TIPO DE OPERACIÓN",
        "R CLASIFICACIÓN", "S SECTOR", "T TIPO DE COSTO / GASTO", "U NÚMERO DE ANEXO"
    ]
    
    rows = []
    count = 0
    
    # Listar archivos JSON
    try:
        files = [f for f in os.listdir(JSON_DIR) if f.lower().endswith(".json")]
    except OSError as e:
        print(f"Error al acceder a {JSON_DIR}: {e}")
        return

    print(f"Procesando {len(files)} archivos en {JSON_DIR}...")

    for filename in files:
        filepath = os.path.join(JSON_DIR, filename)
        
        try:
            # Intentar leer con utf-8, si falla probar latin-1 o cp1252
            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except UnicodeDecodeError:
                with open(filepath, "r", encoding="latin-1") as f:
                    data = json.load(f)
                
            # Filtrar solo DTE 03 (Crédito Fiscal)
            identificacion = data.get("identificacion", {})
            tipo_dte = identificacion.get("tipoDte")
            
            if tipo_dte != "03":
                continue
                
            emisor = data.get("emisor", {})
            resumen = data.get("resumen", {})
            
            # --- Mapeo de Columnas ---
            
            # A: Fecha Emisión
            col_a = format_date(identificacion.get("fecEmi", ""))
            
            # B: Clase Documento (4 = DTE)
            col_b = "4" 
            
            # C: Tipo Documento (03 = Crédito Fiscal)
            col_c = "03" 
            
            # D: Número Documento (Código Generación)
            col_d = identificacion.get("codigoGeneracion", "").replace("-", "")
            
            # E: NIT o NRC
            nit = emisor.get("nit", "").replace("-", "")
            nrc = emisor.get("nrc", "").replace("-", "")
            # Prioridad NRC, si no tiene usa NIT (aunque regla dice NRC obligatorio si no hay DUI)
            col_e = nrc if nrc else nit
            
            # F: Nombre Proveedor
            col_f = emisor.get("nombre", "")
            
            # Montos
            col_g = "0.00" # Internas Exentas
            col_h = "0.00" # Internaciones Exentas
            col_i = "0.00" # Importaciones Exentas
            
            # J: Compras Internas Gravadas
            gravada = float(resumen.get("totalGravada", 0))
            col_j = "{:.2f}".format(gravada)
            
            col_k = "0.00" # Internaciones Gravadas
            col_l = "0.00" # Importaciones Gravadas Bienes
            col_m = "0.00" # Importaciones Gravadas Servicios
            
            # N: Crédito Fiscal (IVA) - Recalculado al 13% exacto de J
            iva_calculado = round(gravada * 0.13, 2)
            col_n = "{:.2f}".format(iva_calculado)
            
            # G, H, I: Exentas
            exenta = float(resumen.get("totalExenta", 0))
            col_g = "{:.2f}".format(exenta)
            col_h = "0.00"
            col_i = "0.00"

            # O: Total Compras - Suma recalculada para consistencia
            # Total = Gravada + IVA + Exenta
            total_calculado = gravada + iva_calculado + exenta
            col_o = "{:.2f}".format(total_calculado)
            
            # P: DUI (Opcional/Vacío para PJ)
            col_p = "" 
            
            # Q, R, S, T: Clasificación
            col_q, col_r, col_s, col_t = get_classification(data)
            
            # U: Número Anexo (3)
            col_u = "3"
            
            row = [
                col_a, col_b, col_c, col_d, col_e, col_f, col_g, col_h, col_i, 
                col_j, col_k, col_l, col_m, col_n, col_o, col_p, 
                col_q, col_r, col_s, col_t, col_u
            ]
            rows.append(row)
            count += 1
            
        except json.JSONDecodeError:
            print(f"Error: {filename} no es un JSON válido.")
        except Exception as e:
            print(f"Error procesando {filename}: {e}")

    # Escribir CSV
    if rows:
        try:
            # encoding='utf-8' sin BOM para compatibilidad con validadores estrictos
            with open(output_file, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f, delimiter=";") 
                # writer.writerow(headers) # Requerimiento: Sin encabezados
                writer.writerows(rows)
            print("-" * 30)
            print(f"Generado archivo CSV exitosamente.")
            print(f"Ubicación: {output_file}")
            print(f"Total documentos procesados: {count}")
            print("-" * 30)

        except Exception as e:
            print(f"Error al escribir el archivo CSV: {e}")
            
        # Generar Excel (Con encabezados)
        try:
            xlsx_file = output_file.replace(".csv", ".xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Anexo Compras"
            ws.append(headers) # Requerimiento: Con encabezados
            for row in rows:
                ws.append(row)
            wb.save(xlsx_file)
            print(f"Generado Excel exitosamente.")
            print(f"Ubicación: {xlsx_file}")
            print("-" * 30)
        except Exception as e:
            print(f"Error generando Excel: {e}")
    else:
        print("No se encontraron documentos DTE-03 válidos para procesar.")

if __name__ == "__main__":
    main()
