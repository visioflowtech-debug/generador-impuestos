import shutil
import generar_anexo  # Importar el script generador
import imaplib
import email
import os
import json
import sys
import datetime

# Configuración Global
username = "optica.nueva.imagensv@gmail.com"
password = "cdmy iyyg txqk ertd"
imap_url = "imap.gmail.com"
output_dir = "C:/phytonmailjson"

# Rango de fechas (mes anterior completo)
def get_previous_month_range():
    today = datetime.date.today()
    # ... (resto de la función igual)
    first_current = today.replace(day=1)
    last_previous = first_current - datetime.timedelta(days=1)
    first_previous = last_previous.replace(day=1)
    
    months = {
        1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
        7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
    }
    
    fmt = "{day:02d}-{month}-{year}"
    
    since_str = fmt.format(
        day=first_previous.day,
        month=months[first_previous.month],
        year=first_previous.year
    )
    
    before_str = fmt.format(
        day=first_current.day,
        month=months[first_current.month],
        year=first_current.year
    )
    
    return since_str, before_str

date_since, date_before = get_previous_month_range()
print(f"Buscando correos desde {date_since} antes de {date_before}")

def clean(text):
    """Limpia el texto para crear nombres de archivo válidos."""
    return "".join(c if c.isalnum() else "_" for c in text)

def decode_payload(part):
    """Decodifica el contenido de un adjunto manejando errores."""
    payload = part.get_payload(decode=True)
    for encoding in ("utf-8", "latin-1"):
        try:
            return payload.decode(encoding)
        except UnicodeDecodeError:
            continue
    return payload.decode(errors="ignore")

def clear_directory(directory):
    """Elimina todos los archivos de un directorio."""
    if os.path.exists(directory):
        for filename in os.listdir(directory):
            file_path = os.path.join(directory, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f"Error al borrar {file_path}. Razón: {e}")

def process_mailbox():
    """Se conecta a la casilla de correo y procesa los emails."""
    if not username or not password:
        print("Error: Credenciales no configuradas.")
        sys.exit(1)

    try:
        # Carpetas de salida específicas
        json_output_dir = os.path.join(output_dir, "json")
        pdf_output_dir = os.path.join(output_dir, "pdf")

        # Limpiar directorios antes de descargar
        print("Limpiando carpetas de salida...")
        clear_directory(json_output_dir)
        clear_directory(pdf_output_dir)

        os.makedirs(json_output_dir, exist_ok=True)
        os.makedirs(pdf_output_dir, exist_ok=True)

        mail = imaplib.IMAP4_SSL(imap_url)
        # ... (resto del código de IMAP) ...
        # (El bloque try/except original continúa aquí, asegurando que mail.login y mail.select estén dentro)
        mail.login(username, password)
        mail.select("inbox")

        # Buscar correos solo por rango de fechas
        search_criteria = f'(SINCE {date_since} BEFORE {date_before})'
        status, messages = mail.search(None, search_criteria)

        if status != "OK":
            print("Error al buscar correos:", messages)
            return

        email_ids = messages[0].split()
        print(f"Se encontraron {len(email_ids)} correos en el rango de fechas.")
        
        # ... (loop de descarga) ...
        for email_id in email_ids:
            # Fetch solo cabeceras primero para evitar descarga innecesaria si no hay adjuntos
            # Pero en este script se usa RFC822 completo, lo mantendremos simple
            try:
                status, msg_data = mail.fetch(email_id, "(RFC822)")
            except Exception as e:
                print(f"Error obteniendo correo {email_id}: {e}")
                continue

            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])

                    # Verificar adjuntos
                    attachments = [part for part in msg.walk() if part.get_filename()]
                    if not attachments:
                        continue

                    # Buscar si algún JSON cumple condición
                    should_download = False
                    for part in attachments:
                        filename = part.get_filename()
                        if filename and filename.lower().endswith(".json"):
                            try:
                                json_content = decode_payload(part)
                                json_data = json.loads(json_content)
                                if json_data.get("identificacion", {}).get("tipoDte") == "03":
                                    should_download = True
                                    break
                            except Exception as e:
                                print(f"Error procesando JSON en correo {email_id.decode()}: {e}")

                    # Descargar adjuntos si cumple
                    if should_download:
                        print(f"Correo {email_id.decode()} cumple la condición. Descargando adjuntos...")
                        for part in attachments:
                            filename = part.get_filename()
                            if filename:
                                # Separar nombre y extensión para limpiar solo el nombre
                                name, ext = os.path.splitext(filename)
                                clean_name = clean(name) + ext
                                
                                target_dir = output_dir # Default
                                if ext.lower() == ".json":
                                    target_dir = json_output_dir
                                elif ext.lower() == ".pdf":
                                    target_dir = pdf_output_dir
                                else:
                                    pass

                                filepath = os.path.join(target_dir, clean_name)
                                with open(filepath, "wb") as f:
                                    f.write(part.get_payload(decode=True))
                                print(f"-> Descargado '{clean_name}' en '{target_dir}'")

    except imaplib.IMAP4.error as e:
        print(f"Error de IMAP: {e}")
    finally:
        if 'mail' in locals() and mail.state == 'SELECTED':
            mail.logout()
            
        print("-" * 30)
        print(f"Descarga finalizada.")
        print(f"Archivos JSON guardados en: {os.path.abspath(os.path.join(output_dir, 'json'))}")
        print(f"Archivos PDF guardados en: {os.path.abspath(os.path.join(output_dir, 'pdf'))}")
        print("-" * 30)
        
        # Ejecutar generación de anexo automáticamente
        print("Ejecutando generación de Anexo CSV...")
        try:
            generar_anexo.main()
        except Exception as e:
            print(f"Error generando el anexo: {e}")
        print("-" * 30)

if __name__ == "__main__":
    process_mailbox()

