# Descuentos — Validación y Conciliación Automática de Ajustes (SQL Server)

Stored procedure en T-SQL que automatiza el proceso diario de validación de descuentos y ajustes ofrecidos a clientes de cobranzas, reemplazando una conciliación manual multi-fuente.

## 📌 Problema que resuelve

El área de cobranzas ofrece descuentos y ajustes a clientes bajo ciertas condiciones de pago. Validar diariamente qué clientes cumplieron esas condiciones (pagaron lo correspondiente) y cuáles no, cruzando contra pagos registrados, ajustes previos y reportes de otras áreas para evitar duplicidad, era un proceso manual propenso a errores y a doble conteo.

## ⚙️ Cómo funciona

1. **Carga de datos:** consolida datos desde un archivo plano vía `BULK INSERT`.
2. **Cruce multi-fuente:** concilia contra pagos registrados en el sistema, ajustes previos ya aplicados, y reportes de otras áreas, para evitar procesar dos veces el mismo caso.
3. **Validación matemática:** clasifica cada caso comparando saldo, monto pagado y ajuste realizado.
   - **Concretado:** el cliente pagó y el descuento aplica.
   - **Ofrecido:** el descuento está pendiente de cumplimiento.
4. **Generación de reportes:** exporta automáticamente tres reportes Excel diferenciados (resultado final, casos que cumplen, casos que no cumplen), además de un archivo plano para consumo de otro sistema.

## 📊 Resultado

| Métrica | Antes | Después | Mejora |
|---|---|---|---|
| Tiempo de procesamiento y conciliación | 90 min | 8 min | **-91%** |

## 🛠️ Tecnologías

- **Motor:** SQL Server (T-SQL)

## 🔧 Técnicas destacables

- Tablas temporales (`#`) y globales (`##`) para procesamiento por etapas
- `BULK INSERT` desde archivo plano delimitado
- `ROW_NUMBER() OVER (PARTITION BY...)` para deduplicación de registros
- Múltiples `JOIN` anidados para conciliación multi-fuente (pagos, ajustes, reportes externos)
- Automatización de exportación y transferencia de archivos hacia un sistema externo

## ⚠️ Nota sobre los datos

Este script fue sanitizado para el portfolio: todas las rutas de red, credenciales de base de datos y de transferencia de archivos fueron reemplazadas por placeholders genéricos (`SERVIDOR_DB`, `USUARIO_DB`, `PASSWORD_DB`, etc.). La lógica de negocio y estructura del stored procedure son idénticas a la versión en producción.
