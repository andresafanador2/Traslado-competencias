import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta
from docx import Document
import os
import unicodedata

# ------------------------- FESTIVOS -------------------------
festivos = [
    "01/01/2025","06/01/2025","24/03/2025","17/04/2025","18/04/2025",
    "01/05/2025","02/06/2025","23/06/2025","30/06/2025","20/07/2025",
    "07/08/2025","18/08/2025","13/10/2025","03/11/2025","17/11/2025",
    "08/12/2025","25/12/2025","01/01/2026","12/01/2026","23/03/2026",
    "02/04/2026","03/04/2026","01/05/2026","18/05/2026","08/06/2026",
    "15/06/2026","29/06/2026","20/07/2026","07/08/2026","17/08/2026",
    "12/10/2026","02/11/2026","16/11/2026","08/12/2026","25/12/2026",
    "01/01/2027","11/01/2027","22/03/2027","25/03/2027","26/03/2027",
    "01/05/2027","10/05/2027","31/05/2027","07/06/2027","05/07/2027",
    "20/07/2027","07/08/2027","16/08/2027","18/10/2027","01/11/2027",
    "15/11/2027","08/12/2027","25/12/2027"
]
festivos = [datetime.strptime(fecha, "%d/%m/%Y").date() for fecha in festivos]

# ------------------------- FUNCIONES -------------------------
def sumar_dias_habiles(fecha_inicio, dias_habiles):
    fecha_actual = fecha_inicio
    dias_sumados = 0
    while dias_sumados < dias_habiles:
        fecha_actual += timedelta(days=1)
        if fecha_actual.weekday() < 5 and fecha_actual not in festivos:
            dias_sumados += 1
    return fecha_actual

# Causales por tipo de periodo
causas = {
    "Exención de contribución": ["Ejecutar inspección en terreno para realizar validaciones respecto a la actividad económica principal reportada en el predio, con el fin de determinar la viabilidad de otorgar el beneficio de exención solicitado"],
    "Revisión por reclamos": ["Ejecutar inspección en terreno para realizar las validaciones y confirmar el funcionamiento del medidor"],
    "Revisión de Mtto": [
        "Ejecutar inspección en terreno para realizar validaciones respecto a las condiciones de la infraestructura reportadas y determinar la viabilidad de los trabajos requeridos",
        "Ejecutar validaciones respecto a la ausencia del servicio de energía",
        "Ejecutar inspección en terreno para realizar validaciones respecto a las condiciones reportadas en la calidad del servicio en el sector"
    ],
    "Revisión de ATC": [
        "Ejecutar inspección en terreno para realizar las validaciones y confirmar el funcionamiento del medidor",
        "Ejecutar inspección en terreno para realizar validaciones con el fin de determinar la viabilidad para la conexión del servicio de energía"
    ],
    "Poda": ["Ejecutar inspección en terreno para realizar validaciones respecto a las condiciones de las redes del sector y determinar el estado de los árboles cercanos a la infraestructura eléctrica"],
    "Daño en equipo eléctrico": ["Ejecutar inspección en terreno para realizar validaciones respecto a las condiciones reportadas por daño en equipo eléctrico"]
}

municipios_norte_santander = ["Abrego", "Los Patios", "Bochalema", "Cucuta", "Villa del Rosario"]

def quitar_tildes(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def actualizar_causas(event):
    seleccion = combo_periodos.get()
    combo_causales['values'] = causas.get(seleccion, [])
    combo_causales.set('')

# Reemplazo run por run para mantener formato
def reemplazar_texto_run(doc, marcador, valor):
    for p in doc.paragraphs:
        for run in p.runs:
            if marcador in run.text:
                run.text = run.text.replace(marcador, valor)

# Función para generar Word
def generar_word(datos, fecha_inicio, dias_habiles, inicio_probatorio, final_probatorio, causal):
    escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
    plantilla_path = os.path.join(escritorio, "Probatorio_plantilla.docx")

    if not os.path.exists(plantilla_path):
        messagebox.showerror("Error", f"No se encontró la plantilla en {plantilla_path}")
        return

    doc = Document(plantilla_path)

    # Determinar departamento
    municipio_normalizado = quitar_tildes(datos["Municipio"].strip().lower())
    departamento = "Norte de Santander" if any(quitar_tildes(m.lower())==municipio_normalizado for m in municipios_norte_santander) else "Cesar"

    reemplazos = {
        "(USUARIO)": datos["Nombre"],
        "(NO_CLIENTE)": datos["No. Cliente"],
        "(DIRECCION)": datos["Dirección"],
        "(CORREO)": datos["Correo electrónico"],
        "(TELEFONO)": datos["Teléfono"],
        "(MUNICIPIO)": datos["Municipio"],
        "(DEPARTAMENTO)": departamento,
        "(RADICADO)": datos["Radicado"],
        "(FECHA)": datos["Fecha de radicación"],
        "(EXPEDIENTE)": datos["Expediente"],
        "(FECHA_INICIO)": fecha_inicio.strftime("%d/%m/%Y"),
        "(DIAS_HABILES)": str(dias_habiles),
        "(INICIO_PROBATORIO)": inicio_probatorio.strftime("%d/%m/%Y"),
        "(FINAL_PROBATORIO)": final_probatorio.strftime("%d/%m/%Y"),
        "(CAUSA)": causal,
        "(PROCESO)": datos["Proceso"]
    }

    for marcador, valor in reemplazos.items():
        reemplazar_texto_run(doc, marcador, valor)

    save_path = os.path.join(escritorio, f"Probatorio_{datos['Nombre']}.docx")
    doc.save(save_path)
    messagebox.showinfo("Word Generado", f"Documento generado en: {save_path}")

# Mostrar formulario de probatorio
def mostrar_formulario_probatorio():
    periodo = combo_periodos.get()
    causal = combo_causales.get()
    fecha_str = entrada_fecha.get()
    dias_str = entrada_dias.get()

    if not periodo or not causal:
        messagebox.showwarning("Atención","Seleccione un periodo y una causal.")
        return
    try:
        fecha_inicio = datetime.strptime(fecha_str, "%d/%m/%Y").date()
        dias_habiles = int(dias_str)
        inicio_probatorio = sumar_dias_habiles(fecha_inicio, 1)
        final_probatorio = sumar_dias_habiles(fecha_inicio, dias_habiles)

        messagebox.showinfo("Periodo calculado",
            f"Periodo: {periodo}\nCausal: {causal}\n"
            f"Fecha inicial: {fecha_inicio.strftime('%d/%m/%Y')}\n"
            f"Fecha final: {final_probatorio.strftime('%d/%m/%Y')}")

        # Limpiar ventana y mostrar formulario
        for widget in root.winfo_children():
            widget.destroy()

        ttk.Label(root, text="Crear nuevo periodo probatorio", font=("Arial",14)).pack(pady=10)

        campos = ["Nombre", "No. Cliente", "Dirección","Correo electrónico","Teléfono","Municipio",
                  "Radicado","Fecha de radicación","Expediente","Proceso"]
        entradas = {}

        for campo in campos:
            ttk.Label(root, text=campo + ":").pack(pady=2)
            ent = ttk.Entry(root, width=40)
            ent.pack(pady=2)
            entradas[campo] = ent

        def guardar_y_generar():
            datos = {campo: ent.get() for campo, ent in entradas.items()}
            generar_word(datos, fecha_inicio, dias_habiles, inicio_probatorio, final_probatorio, causal)

        ttk.Button(root, text="Guardar y Generar Word", command=guardar_y_generar).pack(pady=10)
        ttk.Button(root, text="Crear otro probatorio", command=reiniciar).pack(pady=5)

    except ValueError:
        messagebox.showerror("Error","Revise la fecha o los días ingresados.")

# Función para reiniciar interfaz
def reiniciar():
    for widget in root.winfo_children():
        widget.destroy()
    main_ui()

# ------------------------- INTERFAZ -------------------------
def main_ui():
    global root, combo_periodos, combo_causales, entrada_fecha, entrada_dias
    ttk.Label(root, text="Seleccione el tipo de periodo probatorio:").pack(pady=5)
    combo_periodos = ttk.Combobox(root, values=list(causas.keys()), width=60)
    combo_periodos.pack(pady=5)
    combo_periodos.bind("<<ComboboxSelected>>", actualizar_causas)

    ttk.Label(root, text="Seleccione la causal del periodo:").pack(pady=5)
    combo_causales = ttk.Combobox(root, width=60)
    combo_causales.pack(pady=5)

    ttk.Label(root, text="Digite la fecha inicial (dd/mm/yyyy):").pack(pady=5)
    entrada_fecha = ttk.Entry(root, width=40)
    entrada_fecha.pack(pady=5)

    ttk.Label(root, text="Digite el número de días hábiles:").pack(pady=5)
    entrada_dias = ttk.Entry(root, width=40)
    entrada_dias.pack(pady=5)

    ttk.Button(root, text="Calcular y Continuar", command=mostrar_formulario_probatorio).pack(pady=15)

# ------------------------- EJECUCIÓN -------------------------
root = tk.Tk()
root.title("Apertura de Periodos Probatorios")
root.geometry("550x700")
main_ui()
root.mainloop()
