# CodeAnalyzer 🔍

Herramienta de análisis de código fuente en Python que escanea proyectos **Java** o **C#** y genera automáticamente un reporte profesional en Excel con estadísticas detalladas del proyecto.

## ¿Qué hace?

Lee todos los archivos `.java` o `.cs` de un proyecto, analiza su contenido y genera un archivo `.xlsx` con 4 hojas:

- **Resumen**: totales generales del proyecto (archivos, clases, métodos y líneas)
- **Detalle por Archivo**: métricas individuales por cada archivo analizado
- **Métodos**: listado completo de todos los métodos encontrados
- **Gráfico**: top 10 archivos con más líneas de código en gráfico de barras

## Ejemplo de Reporte

| Métrica | Resultado |
|---------|-----------|
| Archivos analizados | 34 |
| Clases encontradas | 33 |
| Métodos detectados | 289 |
| Líneas totales | 3,208 |
| Líneas de código puro | 2,795 |

> Reporte generado analizando el proyecto [SistemaCochera](https://github.com/angelpin05/SistemaCochera)

## Requisitos

- Python 3.12 o superior
- openpyxl

## Instalación

```bash
pip install openpyxl
```

## Uso

```bash
python code_analyzer.py
```

El script te pedirá:
1. La ruta del proyecto a analizar
2. El lenguaje (Java o C#)

El reporte Excel se genera automáticamente en la misma carpeta del script.

## Tecnologías

- Python 3.12
- openpyxl (generación de Excel)
- re (expresiones regulares para análisis de código)
- os (recorrido de directorios)

## Autor

**Angel Developer**
- GitHub: [@angelpin05](https://github.com/angelpin05)

## Licencia

Este proyecto está bajo la licencia MIT. Ver el archivo [LICENSE](LICENSE) para más detalles.
Al usar este proyecto debes mantener el crédito al autor original.