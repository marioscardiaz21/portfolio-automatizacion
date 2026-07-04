# KPIs IVR — Pipeline Automatizado de Reportería (Python + SQL Server)

Pipeline de extremo a extremo que automatiza el cálculo y reporte de KPIs de una campaña IVR (respuesta de voz interactiva) de cobranzas, desde archivos Excel crudos hasta un reporte ejecutivo final.

## 📌 Problema que resuelve

Cada día se generaban múltiples archivos Excel con resultados de la campaña IVR, que debían consolidarse manualmente, cargarse a base de datos, y luego calcular KPIs (clientes contactados, informados, interesados, recuperados, etc.) para presentar a nivel gerencial. Este proceso completo tomaba cerca de una hora diaria.

## ⚙️ Cómo funciona

1. **Detección y consolidación:** localiza automáticamente todos los archivos Excel diarios de la campaña en una carpeta, según patrón de nombre, y extrae la fecha de carga directamente del nombre del archivo (regex).
2. **Limpieza y tipificación:** normaliza montos (formato decimal) y fechas antes de cargar los datos.
3. **Carga a SQL Server:** trunca la tabla de staging y carga los nuevos registros (con manejo de excepción: si el `TRUNCATE` falla por dependencias, hace `DELETE` como fallback).
4. **Cálculo de KPIs:** ejecuta un stored procedure en SQL Server que calcula los KPIs de negocio sobre los datos recién cargados.
5. **Transformación:** consulta el resultado y lo transforma de formato largo a formato ancho (`pivot_table`) para lectura ejecutiva.
6. **Exportación:** genera un Excel final con formato de fecha aplicado, listo para distribución.

## 📊 Resultado

| Métrica | Antes | Después | Mejora |
|---|---|---|---|
| Tiempo de generación del reporte de KPIs | 60 min | 10 min | **-83%** |

## 🛠️ Tecnologías

- **Lenguaje:** Python 3.x
- **Librerías:** `pandas`, `sqlalchemy`, `pyodbc`, `xlsxwriter`, `re`, `urllib.parse`
- **Base de datos:** SQL Server (carga de datos + stored procedure de cálculo de KPIs)

## 🔧 Técnicas destacables

- Detección dinámica de archivos por patrón de nombre en vez de rutas fijas
- Extracción de metadata (fecha) desde el nombre del archivo vía expresiones regulares
- Pipeline ETL completo: extracción (Excel) → transformación (pandas) → carga (SQL) → procesamiento (stored procedure) → consulta de resultados → presentación (pivot + Excel)
- `pivot_table` con `aggfunc="sum"` para transformar KPIs de formato largo a ancho
- Manejo de excepciones diferenciado (TRUNCATE con fallback a DELETE)
- Conexión segura a base de datos mediante variables de entorno (ver sección de configuración)

## ⚙️ Configuración

Las credenciales de conexión a base de datos **no están hardcodeadas** en este repositorio. El script espera las siguientes variables de entorno:

```bash
export DB_USER="usuario_db"
export DB_PASSWORD="password_db"
export DB_SERVER="servidor_db"
```

## ⚠️ Nota sobre los datos

Este script fue sanitizado para el portfolio: credenciales, rutas de red e IPs internas fueron reemplazadas por variables de entorno y placeholders genéricos. La lógica de negocio y estructura del pipeline son idénticas a la versión en producción.
