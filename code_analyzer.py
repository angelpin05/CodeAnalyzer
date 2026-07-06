import os
import re
import sys
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter
from datetime import datetime

# ─── CONFIGURACIÓN DE ESTILOS ───────────────────────────────────
DARK_BLUE  = "1A3A5C"
MID_BLUE   = "2E6DA4"
LIGHT_BLUE = "D6E4F0"
WHITE      = "FFFFFF"
DARK_GRAY  = "343A40"
LIGHT_GRAY = "F4F4F4"
GREEN      = "27AE60"
ORANGE     = "E67E22"

def header_fill(color):
    return PatternFill("solid", fgColor=color)

def header_font(color=WHITE, bold=True, size=11):
    return Font(name="Calibri", color=color, bold=bold, size=size)

def data_font(bold=False, size=10, color=DARK_GRAY):
    return Font(name="Calibri", bold=bold, size=size, color=color)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

# ─── ANÁLISIS DE CÓDIGO ─────────────────────────────────────────

def analizar_java(contenido):
    """Analiza un archivo Java y retorna clases, métodos y conteos."""
    lineas = contenido.split("\n")
    lineas_totales = len(lineas)
    lineas_codigo = sum(
        1 for l in lineas
        if l.strip() and not l.strip().startswith("//")
        and not l.strip().startswith("*")
        and not l.strip().startswith("/*")
    )

    # Buscar clases
    patron_clase = re.compile(
        r'(public|private|protected)?\s*(abstract|final)?\s*'
        r'(class|interface|enum)\s+(\w+)'
    )
    clases = patron_clase.findall(contenido)
    nombres_clases = [c[3] for c in clases]

    # Buscar métodos
    patron_metodo = re.compile(
        r'(public|private|protected|static|\s)+[\w\<\>\[\]]+\s+'
        r'(\w+)\s*\([^)]*\)\s*(throws\s+[\w,\s]+)?\s*\{'
    )
    metodos = patron_metodo.findall(contenido)
    nombres_metodos = [m[1] for m in metodos
                       if m[1] not in ["if", "while", "for", "switch",
                                       "catch", "try", "else"]]

    # Buscar paquete
    patron_paquete = re.compile(r'^package\s+([\w.]+);', re.MULTILINE)
    paquete = patron_paquete.search(contenido)
    paquete = paquete.group(1) if paquete else "sin paquete"

    return {
        "paquete": paquete,
        "clases": nombres_clases,
        "metodos": nombres_metodos,
        "lineas_totales": lineas_totales,
        "lineas_codigo": lineas_codigo,
    }


def analizar_csharp(contenido):
    """Analiza un archivo C# y retorna clases, métodos y conteos."""
    lineas = contenido.split("\n")
    lineas_totales = len(lineas)
    lineas_codigo = sum(
        1 for l in lineas
        if l.strip() and not l.strip().startswith("//")
        and not l.strip().startswith("*")
        and not l.strip().startswith("/*")
    )

    # Buscar clases
    patron_clase = re.compile(
        r'(public|private|protected|internal)?\s*(abstract|sealed|static)?\s*'
        r'(class|interface|enum|struct)\s+(\w+)'
    )
    clases = patron_clase.findall(contenido)
    nombres_clases = [c[3] for c in clases]

    # Buscar métodos
    patron_metodo = re.compile(
        r'(public|private|protected|internal|static|virtual|override|async|\s)+'
        r'[\w\<\>\[\]]+\s+(\w+)\s*\([^)]*\)\s*(\{|=>)'
    )
    metodos = patron_metodo.findall(contenido)
    nombres_metodos = [m[1] for m in metodos
                       if m[1] not in ["if", "while", "for", "switch",
                                       "catch", "try", "else", "Main"]]

    # Buscar namespace
    patron_ns = re.compile(r'^namespace\s+([\w.]+)', re.MULTILINE)
    namespace = patron_ns.search(contenido)
    namespace = namespace.group(1) if namespace else "sin namespace"

    return {
        "paquete": namespace,
        "clases": nombres_clases,
        "metodos": nombres_metodos,
        "lineas_totales": lineas_totales,
        "lineas_codigo": lineas_codigo,
    }


def escanear_proyecto(ruta, lenguaje):
    """Recorre la carpeta del proyecto y analiza todos los archivos."""
    extension = ".java" if lenguaje == "java" else ".cs"
    resultados = []

    for root, dirs, files in os.walk(ruta):
        # Ignorar carpetas comunes que no son código fuente
        dirs[:] = [d for d in dirs if d not in
                   [".git", "target", "bin", "obj", ".idea",
                    "node_modules", ".mvn", "out", "build"]]

        for archivo in files:
            if archivo.endswith(extension):
                ruta_completa = os.path.join(root, archivo)
                try:
                    with open(ruta_completa, "r", encoding="utf-8",
                              errors="ignore") as f:
                        contenido = f.read()

                    if lenguaje == "java":
                        analisis = analizar_java(contenido)
                    else:
                        analisis = analizar_csharp(contenido)

                    ruta_relativa = os.path.relpath(ruta_completa, ruta)
                    analisis["archivo"] = archivo
                    analisis["ruta"] = ruta_relativa
                    resultados.append(analisis)
                    print(f"  ✓ Analizado: {ruta_relativa}")
                except Exception as e:
                    print(f"  ✗ Error en {archivo}: {e}")

    return resultados


# ─── GENERACIÓN DEL EXCEL ───────────────────────────────────────

def aplicar_estilo_header(ws, fila, columnas, color=DARK_BLUE):
    for col in columnas:
        cell = ws.cell(row=fila, column=col)
        cell.fill = header_fill(color)
        cell.font = header_font()
        cell.alignment = center()
        cell.border = thin_border()


def aplicar_estilo_fila(ws, fila, num_cols, alternado=False):
    color = LIGHT_BLUE if alternado else WHITE
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=fila, column=col)
        cell.fill = PatternFill("solid", fgColor=color)
        cell.font = data_font()
        cell.alignment = left()
        cell.border = thin_border()


def generar_excel(resultados, ruta_salida, nombre_proyecto, lenguaje):
    wb = Workbook()

    # ── Hoja 1: Resumen ─────────────────────────────────────────
    ws_resumen = wb.active
    ws_resumen.title = "Resumen"
    ws_resumen.column_dimensions["A"].width = 35
    ws_resumen.column_dimensions["B"].width = 20

    # Título
    ws_resumen.merge_cells("A1:B1")
    t = ws_resumen["A1"]
    t.value = f"CodeAnalyzer — Reporte de Proyecto"
    t.fill = header_fill(DARK_BLUE)
    t.font = Font(name="Calibri", color=WHITE, bold=True, size=14)
    t.alignment = center()
    ws_resumen.row_dimensions[1].height = 30

    ws_resumen.merge_cells("A2:B2")
    t2 = ws_resumen["A2"]
    t2.value = f"Proyecto: {nombre_proyecto}  |  Lenguaje: {lenguaje.upper()}  |  Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    t2.fill = header_fill(MID_BLUE)
    t2.font = Font(name="Calibri", color=WHITE, size=10)
    t2.alignment = center()
    ws_resumen.row_dimensions[2].height = 20

    # Totales
    total_archivos = len(resultados)
    total_clases   = sum(len(r["clases"]) for r in resultados)
    total_metodos  = sum(len(r["metodos"]) for r in resultados)
    total_lineas   = sum(r["lineas_totales"] for r in resultados)
    total_codigo   = sum(r["lineas_codigo"] for r in resultados)

    ws_resumen.append([])
    metricas = [
        ("Total de archivos analizados", total_archivos),
        ("Total de clases encontradas",  total_clases),
        ("Total de métodos encontrados", total_metodos),
        ("Total de líneas (con vacías y comentarios)", total_lineas),
        ("Total de líneas de código puro", total_codigo),
        ("Promedio de líneas por archivo", round(total_lineas / total_archivos, 1) if total_archivos else 0),
        ("Promedio de métodos por clase",  round(total_metodos / total_clases, 1) if total_clases else 0),
    ]

    for i, (etiqueta, valor) in enumerate(metricas):
        fila = 4 + i
        ws_resumen.cell(row=fila, column=1, value=etiqueta)
        ws_resumen.cell(row=fila, column=2, value=valor)
        alternado = i % 2 == 0
        for col in [1, 2]:
            cell = ws_resumen.cell(row=fila, column=col)
            cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE if alternado else WHITE)
            cell.font = data_font(bold=(col == 1))
            cell.alignment = left() if col == 1 else center()
            cell.border = thin_border()

    # ── Hoja 2: Detalle por Archivo ──────────────────────────────
    ws_detalle = wb.create_sheet("Detalle por Archivo")
    ws_detalle.column_dimensions["A"].width = 30
    ws_detalle.column_dimensions["B"].width = 40
    ws_detalle.column_dimensions["C"].width = 20
    ws_detalle.column_dimensions["D"].width = 15
    ws_detalle.column_dimensions["E"].width = 15
    ws_detalle.column_dimensions["F"].width = 20
    ws_detalle.column_dimensions["G"].width = 20

    # Título
    ws_detalle.merge_cells("A1:G1")
    t = ws_detalle["A1"]
    t.value = "Detalle por Archivo"
    t.fill = header_fill(DARK_BLUE)
    t.font = Font(name="Calibri", color=WHITE, bold=True, size=13)
    t.alignment = center()
    ws_detalle.row_dimensions[1].height = 25

    # Headers
    headers = ["Archivo", "Ruta", "Paquete / Namespace",
               "Clases", "Métodos", "Líneas Totales", "Líneas de Código"]
    for col, h in enumerate(headers, 1):
        cell = ws_detalle.cell(row=2, column=col, value=h)
        cell.fill = header_fill(MID_BLUE)
        cell.font = header_font()
        cell.alignment = center()
        cell.border = thin_border()
    ws_detalle.row_dimensions[2].height = 20

    for i, r in enumerate(resultados):
        fila = 3 + i
        valores = [
            r["archivo"],
            r["ruta"],
            r["paquete"],
            len(r["clases"]),
            len(r["metodos"]),
            r["lineas_totales"],
            r["lineas_codigo"],
        ]
        for col, val in enumerate(valores, 1):
            cell = ws_detalle.cell(row=fila, column=col, value=val)
            cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE if i % 2 == 0 else WHITE)
            cell.font = data_font()
            cell.alignment = left() if col <= 3 else center()
            cell.border = thin_border()

    # ── Hoja 3: Métodos ──────────────────────────────────────────
    ws_metodos = wb.create_sheet("Métodos")
    ws_metodos.column_dimensions["A"].width = 30
    ws_metodos.column_dimensions["B"].width = 30
    ws_metodos.column_dimensions["C"].width = 40

    ws_metodos.merge_cells("A1:C1")
    t = ws_metodos["A1"]
    t.value = "Listado Completo de Métodos"
    t.fill = header_fill(DARK_BLUE)
    t.font = Font(name="Calibri", color=WHITE, bold=True, size=13)
    t.alignment = center()
    ws_metodos.row_dimensions[1].height = 25

    for col, h in enumerate(["Método", "Clase(s) en el Archivo", "Archivo"], 1):
        cell = ws_metodos.cell(row=2, column=col, value=h)
        cell.fill = header_fill(MID_BLUE)
        cell.font = header_font()
        cell.alignment = center()
        cell.border = thin_border()

    fila_m = 3
    for r in resultados:
        clases_str = ", ".join(r["clases"]) if r["clases"] else "sin clase"
        for i, metodo in enumerate(r["metodos"]):
            valores = [metodo, clases_str, r["archivo"]]
            for col, val in enumerate(valores, 1):
                cell = ws_metodos.cell(row=fila_m, column=col, value=val)
                cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE if fila_m % 2 == 0 else WHITE)
                cell.font = data_font()
                cell.alignment = left()
                cell.border = thin_border()
            fila_m += 1

    # ── Hoja 4: Gráfico ──────────────────────────────────────────
    ws_grafico = wb.create_sheet("Gráfico")

    ws_grafico.merge_cells("A1:B1")
    t = ws_grafico["A1"]
    t.value = "Top 10 Archivos con más Líneas de Código"
    t.fill = header_fill(DARK_BLUE)
    t.font = Font(name="Calibri", color=WHITE, bold=True, size=13)
    t.alignment = center()

    for col, h in enumerate(["Archivo", "Líneas de Código"], 1):
        cell = ws_grafico.cell(row=2, column=col, value=h)
        cell.fill = header_fill(MID_BLUE)
        cell.font = header_font()
        cell.alignment = center()
        cell.border = thin_border()

    top10 = sorted(resultados, key=lambda x: x["lineas_codigo"], reverse=True)[:10]
    for i, r in enumerate(top10):
        fila = 3 + i
        ws_grafico.cell(row=fila, column=1, value=r["archivo"])
        ws_grafico.cell(row=fila, column=2, value=r["lineas_codigo"])
        for col in [1, 2]:
            cell = ws_grafico.cell(row=fila, column=col)
            cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE if i % 2 == 0 else WHITE)
            cell.font = data_font()
            cell.alignment = left() if col == 1 else center()
            cell.border = thin_border()

    # Gráfico de barras
    chart = BarChart()
    chart.type = "col"
    chart.title = "Top 10 Archivos por Líneas de Código"
    chart.y_axis.title = "Líneas de Código"
    chart.x_axis.title = "Archivo"
    chart.style = 10
    chart.width = 20
    chart.height = 12

    data_ref = Reference(ws_grafico, min_col=2, min_row=2,
                         max_row=2 + len(top10))
    cats_ref = Reference(ws_grafico, min_col=1, min_row=3,
                         max_row=2 + len(top10))
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    ws_grafico.add_chart(chart, "D2")

    # Guardar
    wb.save(ruta_salida)
    print(f"\n✅ Reporte generado: {ruta_salida}")


# ─── PUNTO DE ENTRADA ───────────────────────────────────────────

def main():
    print("=" * 55)
    print("       CodeAnalyzer — Analizador de Código Fuente")
    print("=" * 55)

    # Solicitar ruta del proyecto
    ruta = input("\n📁 Ruta del proyecto (Enter para carpeta actual): ").strip()
    if not ruta:
        ruta = os.getcwd()

    if not os.path.exists(ruta):
        print(f"❌ La ruta '{ruta}' no existe.")
        sys.exit(1)

    # Solicitar lenguaje
    print("\n🔤 Lenguaje:")
    print("   1. Java")
    print("   2. C#")
    opcion = input("Selecciona (1 o 2): ").strip()
    lenguaje = "java" if opcion == "1" else "csharp"

    nombre_proyecto = os.path.basename(os.path.abspath(ruta))
    print(f"\n🔍 Analizando proyecto: {nombre_proyecto}")
    print(f"   Lenguaje: {'Java' if lenguaje == 'java' else 'C#'}")
    print(f"   Ruta: {ruta}\n")

    # Escanear archivos
    resultados = escanear_proyecto(ruta, lenguaje)

    if not resultados:
        print(f"\n❌ No se encontraron archivos "
              f"{'Java (.java)' if lenguaje == 'java' else 'C# (.cs)'} en la ruta indicada.")
        sys.exit(1)

    print(f"\n📊 Archivos analizados: {len(resultados)}")

    # Generar Excel
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_excel = f"Reporte_{nombre_proyecto}_{timestamp}.xlsx"
    ruta_salida = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               nombre_excel)

    print("\n📝 Generando reporte Excel...")
    generar_excel(resultados, ruta_salida, nombre_proyecto, lenguaje)

    print(f"\n📈 Resumen:")
    print(f"   Archivos : {len(resultados)}")
    print(f"   Clases   : {sum(len(r['clases']) for r in resultados)}")
    print(f"   Métodos  : {sum(len(r['metodos']) for r in resultados)}")
    print(f"   Líneas   : {sum(r['lineas_totales'] for r in resultados)}")
    print("\n✅ ¡Listo! Abre el archivo Excel para ver el reporte completo.")
    input("\nPresiona Enter para salir...")


if __name__ == "__main__":
    main()