import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, PhotoImage
import ttkbootstrap as ttk
import smtplib

# --- Agregar carpeta raíz al PYTHONPATH ---
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from services.excel_reader import leer_datos_excel
from services.email_sender import preparar_mensaje
from services.excel_writer import marcar_como_enviado

def obtener_ruta_recurso(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class EnvioMasivoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🚀 Envio Masivo de Correos - Astrana Health")
        self.root.geometry("750x700")
        self.root.resizable(False, False)

        icon_path = obtener_ruta_recurso(os.path.join("images", "logo_barratareas.png"))
        if os.path.exists(icon_path):
            icono = PhotoImage(file=icon_path)
            self.root.iconphoto(False, icono)

        self.ruta_excel = tk.StringVar()
        self.correo_emisor = tk.StringVar()
        self.password = tk.StringVar()
        self.nombre_emisor = tk.StringVar()
        self.detener_envio = False
        self.contador_enviados = 0
        self.contador_omitidos = 0

        self.construir_interfaz()

    def construir_interfaz(self):
        estilo = ttk.Style("flatly")
        marco = ttk.Frame(self.root, padding=20)
        marco.pack(fill="both", expand=True)

        ttk.Button(marco, text="❓ Cómo funciona", bootstyle="info", command=self.abrir_instrucciones).pack(anchor="ne")

        ttk.Label(marco, text="Archivo Excel:", font=("Arial", 12, "bold")).pack(anchor="w")
        frame_archivo = ttk.Frame(marco)
        frame_archivo.pack(fill="x", pady=5)
        ttk.Entry(frame_archivo, textvariable=self.ruta_excel, width=60, state="readonly").pack(side="left", padx=(0,10))
        ttk.Button(frame_archivo, text="Seleccionar archivo", bootstyle="primary", command=self.seleccionar_archivo).pack(side="left")

        ttk.Label(marco, text="Correo Gmail:", font=("Arial", 12, "bold")).pack(anchor="w", pady=(15,0))
        ttk.Entry(marco, textvariable=self.correo_emisor).pack(fill="x")

        ttk.Label(marco, text="Contraseña de aplicación:", font=("Arial", 12, "bold")).pack(anchor="w", pady=(15,0))
        ttk.Entry(marco, textvariable=self.password, show="*").pack(fill="x")

        ttk.Label(marco, text="Tu nombre (firma):", font=("Arial", 12, "bold")).pack(anchor="w", pady=(15,0))
        ttk.Entry(marco, textvariable=self.nombre_emisor).pack(fill="x")

        estilo.configure("Enviar.TButton", font=("Arial", 14, "bold"))
        boton_frame = ttk.Frame(marco)
        boton_frame.pack(pady=10)

        self.boton_enviar = ttk.Button(boton_frame, text="🚀 Enviar Correos", bootstyle="success", command=self.enviar_correos, width=25, style="Enviar.TButton")
        self.boton_enviar.pack(side="left", padx=10)

        self.boton_detener = ttk.Button(boton_frame, text="⛔️ Detener", bootstyle="danger", command=self.detener_envio_correos, state="disabled")
        self.boton_detener.pack(side="left")

        # Barra de progreso
        self.progress = ttk.Progressbar(marco, orient="horizontal", length=400, mode="determinate", bootstyle="info-striped")
        self.progress.pack(pady=10)

        self.log = tk.Text(marco, height=10, state="disabled")
        self.log.pack(fill="both", expand=True, pady=10)

    def seleccionar_archivo(self):
        archivo = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=[("Archivos Excel", "*.xlsx")])
        if archivo:
            self.ruta_excel.set(archivo)

    def enviar_correos(self):
        ruta_excel = self.ruta_excel.get()
        correo_emisor = self.correo_emisor.get()
        password = self.password.get()
        nombre_emisor = self.nombre_emisor.get()

        if not ruta_excel or not correo_emisor or not password or not nombre_emisor:
            messagebox.showerror("Error", "Por favor completa todos los campos.")
            return

        try:
            plantilla_path = obtener_ruta_recurso(os.path.join("templates", "cuerpo_correo.html"))
            with open(plantilla_path, "r", encoding="utf-8") as f:
                plantilla_html = f.read()

            datos = leer_datos_excel(ruta_excel, "Envio masivo")
            if not datos:
                messagebox.showinfo("Información", "Todos los correos ya fueron enviados.")
                return

            total_correos = len(datos)
            progreso_por_cada = 100 / total_correos
            self.progress['value'] = 0

            smtp_server = 'smtp.gmail.com'
            smtp_port = 587

            self.detener_envio = False
            self.contador_enviados = 0
            self.contador_omitidos = 0
            self.boton_detener.config(state="normal")

            with smtplib.SMTP(smtp_server, smtp_port) as servidor:
                servidor.starttls()
                servidor.login(correo_emisor, password)

                self.log_mensaje("📢 Conexión segura establecida. Listo para enviar correos.")

                for destinatario in datos:
                    if self.detener_envio:
                        self.log_mensaje("⛔ Envío detenido por el usuario.")
                        break

                    if not destinatario['correo'] or not destinatario['nombre']:
                        continue

                    if not destinatario['enviado']:
                        asunto = "Recordatorio de Envío de Documento"

                        fecha_expiracion = destinatario['fecha_expiracion']
                        fecha_texto = fecha_expiracion.strftime('%Y-%m-%d') if fecha_expiracion else "Fecha no disponible"

                        cuerpo_final = plantilla_html.format(
                            nombre=destinatario['nombre'],
                            fecha_expiracion=fecha_texto,
                            nombre_emisor=nombre_emisor,
                            correo_emisor=correo_emisor
                        )

                        mensaje = preparar_mensaje(correo_emisor, destinatario['correo'], asunto, cuerpo_final)

                        servidor.send_message(mensaje)
                        marcar_como_enviado(ruta_excel, destinatario['correo'], "Envio masivo")

                        self.contador_enviados += 1
                        self.log_mensaje(f"✅ Correo enviado a {destinatario['nombre']} ({destinatario['correo']})")
                    else:
                        self.contador_omitidos += 1
                        self.log_mensaje(f"⚡ Se omitió a {destinatario['nombre']} porque ya estaba enviado.")

                    # Actualizar progreso
                    self.progress['value'] += progreso_por_cada
                    self.root.update_idletasks()

            self.boton_detener.config(state="disabled")

            if self.contador_enviados > 0:
                messagebox.showinfo("Éxito", f"✅ Se enviaron {self.contador_enviados} correos correctamente. ({self.contador_omitidos} omitidos)")
            else:
                messagebox.showinfo("Información", "⚡ No había correos pendientes para enviar.")

            self.progress['value'] = 100

        except Exception as e:
            self.boton_detener.config(state="disabled")
            messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")

    def detener_envio_correos(self):
        self.detener_envio = True

    def abrir_instrucciones(self):
        contenido = """Bienvenido a Envío Masivo Astrana Health.

REQUISITOS:
- Descargar el .exe en python.
- El Excel debe tener columnas: Nombre, Email, Fecha de expiración, Enviado.
- La hoja debe llamarse exactamente: "Envio masivo".

FUNCIONAMIENTO:
1. Selecciona tu Excel.
2. Ingresa tu correo y contraseña de aplicación.
3. Coloca tu nombre (firma).
4. Haz clic en Enviar Correos.

Notas:
- Cada correo se envía en 1-2 segundos.
- El Excel se actualiza automáticamente.
- Puedes detener el proceso si lo deseas.

¡Gracias por usar esta herramienta!
"""
        ventana = tk.Toplevel(self.root)
        ventana.title("📄 Instrucciones de Uso")
        ventana.geometry("650x500")
        texto = tk.Text(ventana, wrap="word")
        texto.insert("1.0", contenido)
        texto.pack(expand=True, fill="both")
        texto.config(state="disabled")

    def log_mensaje(self, mensaje):
        self.log.config(state="normal")
        self.log.insert("end", mensaje + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

if __name__ == "__main__":
    root = ttk.Window(themename="flatly")
    app = EnvioMasivoApp(root)
    root.mainloop()
