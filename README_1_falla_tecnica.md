# Falla Técnica — Automatización de Reporte Diario Unificado

Script en Python que automatiza el reporte diario de "Falla Técnica" del área de cobranzas, un proceso que antes se realizaba manualmente en ~40 minutos.

## 📌 Problema que resuelve

El equipo operativo necesitaba, todos los días, cruzar tres bases de datos distintas (cero pagos, falla técnica y un pago) contra el reporte de transferencias del día, para determinar si cada caso había sido derivado o no a otra área. Este cruce se hacía manualmente en Excel, con alto riesgo de error por inconsistencias de formato entre archivos (tildes, mayúsculas/minúsculas, distintos formatos de fecha).

## ⚙️ Cómo funciona

1. **Detección automática de archivos:** localiza el archivo más reciente de cada base según fecha de modificación, sin necesidad de indicar la ruta manualmente cada día.
2. **Unificación:** consolida las tres bases en un único DataFrame.
3. **Normalización:** limpia texto (acentos, mayúsculas, espacios) y estandariza fechas en dos formatos distintos (DD/MM/YYYY y YYYY/MM/DD) para tolerar inconsistencias entre fuentes.
4. **Cruce con trazabilidad:** hace merge contra la base de Transferencias del día usando el ID de cliente como clave, con `indicator=True` para detectar explícitamente qué registros no encontraron match.
5. **Homologación de canal:** aplica un mapeo externo de motivos para clasificar cada caso.
6. **Exportación con formato:** genera un Excel final con encabezados de dos colores según sección y columnas autoajustadas, listo para entrega sin intervención manual.

## 📊 Resultado

| Métrica | Antes | Después | Mejora |
| Tiempo de generación del reporte | 40 min | 3 min | **-93%** |

## 🛠️ Tecnologías

- **Lenguaje:** Python 3.x
- **Librerías:** `pandas`, `openpyxl`, `unicodedata`, `re`

## 🔧 Técnicas destacables

- `merge` con `indicator=True` para trazabilidad del cruce (detección explícita de registros sin match)
- Detección dinámica de archivos por patrón de nombre + fecha de modificación (`os.path.getmtime`)
- Matching de columnas por nombres candidatos, tolerante a variaciones de encabezado entre fuentes
- Manejo dual de formatos de fecha (DD/MM/YYYY vs YYYY/MM/DD)
- Normalización de texto con `unicodedata` y expresiones regulares

## ⚠️ Nota sobre los datos

Este script fue sanitizado para el portfolio: las rutas de red originales y nombres de personas fueron reemplazados por placeholders genéricos. La lógica y estructura del código son idénticas a la versión en producción.
