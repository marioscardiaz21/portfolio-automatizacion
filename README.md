# Portafolio de Automatización de Procesos y Datos

Selección de proyectos de automatización desarrollados en un entorno operativo real (sector telecomunicaciones), sanitizados de toda información confidencial (credenciales, IPs internas, nombres de compañeros y de cliente).

Cada carpeta contiene un proyecto independiente con su propio README detallando el problema resuelto, la lógica implementada y las tecnologías utilizadas.

## 📂 Proyectos

| Proyecto | Stack | Problema que resuelve | Impacto |
|---|---|---|---|
| [`falla-tecnica/`](./falla-tecnica) | Python (pandas, openpyxl) | Unificación y cruce diario de reportes de falla técnica en cobranzas | -93% tiempo de procesamiento (40 → 3 min) |
| [`kpis-ivr/`](./kpis-ivr) | Python + SQL Server | Pipeline completo de reportería de KPIs de campaña IVR | -83% tiempo de generación (60 → 10 min) |
| [`descuentos/`](./descuentos) | SQL Server (T-SQL) | Validación y conciliación automática de descuentos a clientes | -91% tiempo de procesamiento (90 → 8 min) |

## 🧠 Contexto general

Estos scripts fueron desarrollados para automatizar procesos operativos de cobranzas que antes se ejecutaban manualmente en Excel, dentro de un entorno de alto volumen de datos y múltiples fuentes (bases de datos, archivos planos, reportes de terceros). El objetivo común en los tres casos es el mismo: **reducir tiempos operativos, eliminar el error humano en la conciliación de datos y entregar información lista para decisión** (operativa o gerencial).

## 🛠️ Stack general utilizado

- **Lenguajes:** Python 3.x, T-SQL (SQL Server)
- **Librerías Python:** pandas, openpyxl, sqlalchemy, pyodbc, xlsxwriter, unicodedata, re
- **Base de datos:** SQL Server (Stored Procedures, BULK INSERT, tablas temporales)
- **Técnicas transversales:** normalización de texto y fechas, cruces multi-fuente, deduplicación, generación de reportes ejecutivos formateados

## ⚠️ Nota sobre los datos

Todo el código en este repositorio ha sido sanitizado: no contiene credenciales, direcciones IP internas, nombres de compañeros de trabajo ni datos de clientes reales. Las rutas de red y credenciales de conexión fueron reemplazadas por variables de entorno o placeholders genéricos.

## 👤 Autor

**Marioscar Diaz** — Analista de Automatización de Procesos
[LinkedIn](https://www.linkedin.com/in/marioscar-diaz-56561111b/)
