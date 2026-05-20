# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Propósito del proyecto

Imputación automática de cuotas: tomar las transferencias bancarias del archivo `imputaciones.xlsx` (hoja semanal como "S 121", "S 122", etc. o "USD 5", "USD 6"...) y registrarlas en `deudores.xlsx` (monto real, número de cuota, fecha).

## Archivos clave

```
datos/
  IMPUTACIONES 02.03.26.xlsx   ← archivo de transferencias bancarias
  DEUDORES FINK-4-2026.xlsx    ← planilla de deudores/cuotas
  comprobantes_cache.json      ← cache local del Google Sheet "Datos comprobantes"
imputar_s120.py                ← script pesos (listo para S 121, DRY_RUN=True)
imputar_usd5.py                ← script USD (listo para USD 5, DRY_RUN=True)
comprobantes_helper.py         ← módulo de acceso al cache de comprobantes
exportar_comprobantes.py       ← refresca el cache desde Google Sheets
```

## Cache de comprobantes (Google Sheets)

El sheet "Datos comprobantes" (ID: `15p_2PwVOhABPg_UoEbkehci9Y8PE0nCWN88U_dO_uwU`) tiene los datos de clientes que mandaron comprobantes al bot. Se usa como fallback adicional cuando el CUIT no se encuentra en deudores ni en hojas anteriores de imputaciones.

**Antes de cada imputación semanal**, refrescar el cache:
```
python exportar_comprobantes.py
```
Requiere `pip install gspread`. La primera vez abre el navegador para autorizar (OAuth). Después queda guardado.

Si no se corre el exportar, igual funciona con el último cache guardado en `datos/comprobantes_cache.json`.

## Cómo correr la imputación de una nueva semana (pesos)

1. **Cambiar las variables en `imputar_s120.py`** (primeras líneas de CONFIG):
   - `IMP_SHEET = 'S 121'` → nombre de la hoja a procesar (ya configurado)
   - `DRY_RUN = True` → siempre empezar en True para revisar (ya configurado)

2. **Verificar las columnas de mes en `SHEETS_CFG`** si cambió el mes:
   - Cada mes nuevo agrega 4 columnas a la derecha en deudores
   - Buscar los headers `pago MES AÑO teorico/real/NUMERO DE CUOTA/fecha depo`
   - Actualmente configurado para **Mayo 2026**

3. **Correr dry-run**: `python imputar_s120.py`
4. **Revisar el reporte**, especialmente los casos ambiguos
5. **Cambiar `DRY_RUN = False`** y volver a correr para escribir

## Cómo correr la imputación USD

1. Usar `imputar_usd5.py` (ya configurado para "USD 5", DRY_RUN=True)
2. Para la siguiente semana USD, cambiar `IMP_SHEET = 'USD 6'` (etc.)
3. Mismo flujo que pesos: dry-run → revisar → DRY_RUN=False

## Estructura de deudores.xlsx

Hojas con cuotas en pesos (NO escribir en '$ USD fijo' desde el script pesos):

| Hoja | Header row | Data desde | CUIT col | Nombre col |
|------|-----------|------------|----------|-----------|
| INDICE CAC | fila 5 | fila 6 | L (12) | I (9) |
| INDICE CAC M. OBRA | fila 3 | fila 4 | K (11) | I (9) |
| BOLSA CEMENTO | fila 3 | fila 4 | N (14) | J (10) |

**Columnas de Mayo 2026** (actualizar cuando cambia el mes):

| Hoja | Teorico | Real | N° Cuota | Fecha |
|------|---------|------|----------|-------|
| INDICE CAC | EF (136) | EG (137) | EH (138) | EI (139) |
| INDICE CAC M. OBRA | EB (132) | EC (133) | ED (134) | EE (135) |
| BOLSA CEMENTO | EE (135) | EF (136) | EG (137) | EH (138) |

Hoja USD (`$  USD fijo`, dos espacios), Mayo 2026:

| Header row | Data desde | CUIT col | Nombre col | Lote col | Teorico | Real | N° Cuota | Fecha |
|-----------|------------|----------|-----------|---------|---------|------|----------|-------|
| fila 3 | fila 4 | M (13) | J (10) | I (9) | col 131 | col 132 | col 133 | col 134 |

CUITs en USD fijo pueden tener formato con guiones (`20-34658691-6`) o múltiples separados por `/`. El script los normaliza automáticamente.

## Cómo determina el script el número de cuota (regla actual)

**NO usar la columna A (MAYOR CUOTA)**. Esa fórmula no funciona correctamente para filas nuevas o cuando se imputan varias cuotas en la misma celda.

En cambio, el script:
1. Lee la columna "NUMERO DE CUOTA" del **mes actual** en la fila del cliente:
   - Si dice `"parte de cuota X"` → imputa como `X+1` (el cliente pagó parte de X, ahora completa)
   - Si es un número normal → ya está imputado este mes, no volver a imputar
2. Si está vacía, escanea **todas las columnas históricas** "NUMERO DE CUOTA" (meses anteriores) en la misma fila y toma el máximo + 1
3. Si no hay historial en deudores, busca en las últimas hojas de imputaciones (semanas previas) la cuota imputada para ese CUIT
4. Si aún no hay dato, usa `CUOTA_OVERRIDE` (dict manual en el CONFIG del script)

## Cómo resuelve el script clientes no encontrados por CUIT

Si el CUIT de una transferencia no está en deudores:
1. Busca ese CUIT en las últimas 8 hojas de imputaciones → extrae el nombre de col H
2. Busca ese nombre en deudores por palabras clave (case-insensitive)
3. Si hay 1 match → imputa; si hay 0 o varios → reporta como ambiguo

## Cómo imputa múltiples lotes del mismo cliente

Si un CUIT tiene N lotes en deudores y llegan N transferencias iguales:
- Si todos los lotes tienen el **mismo nombre**: agrega `l{LOTE}` en el label (ej: `Bustamante Gustavo l6 c13`)
- Si los lotes tienen **nombres distintos**: usa el nombre del lote, sin aclarar lote
- Asigna lotes en orden (por hoja+fila) sin reutilizar el mismo lote en la misma corrida

## Estructura de imputaciones.xlsx

Cada hoja semanal ("S 121", "USD 5", etc.) tiene:
- Col A: Fecha
- Col F: Monto transferido (Crédito)
- Col C: Concepto — contiene el CUIT como número de 11 dígitos (o 12 para cuentas tipo `402XXXXXXXXX`)
- Col H: Nombre del cliente + cuota (se escribe acá al imputar)
- Col I: "x" (marca de procesado, se escribe acá al imputar)
- Fondo amarillo en la fila = ya procesado

## Reglas de negocio

- Solo procesar filas NO amarillas y sin "PAGO MENOS" en col H
- Identificar cliente por CUIT/CUIL extraído del campo Concepto
- Si no hay CUIT extraíble → reportar como ambiguo
- Si pagó MENOS y diferencia > $3.000 (pesos) o > U$D 5 (USD) → escribir "PAGO MENOS" en col H
- Si pagó menos pero diferencia ≤ tolerancia → imputar normalmente
- Si pagó más → imputar normalmente
- Si ya fue imputado este mes → reportar como ambiguo (no sobreescribir)
- Al imputar: escribir en col H `{nombre} [l{lote}] c{cuota}`, en col I `x`, pintar fila amarilla

## Para cambiar de mes (ej: Junio 2026)

Buscar los headers exactos en deudores fila 5 (INDICE CAC) o fila 3 (otras hojas):
```python
python -c "
import openpyxl
wb = openpyxl.load_workbook('datos/DEUDORES FINK-4-2026.xlsx', data_only=True)
ws = wb['INDICE CAC']
for c in range(1, ws.max_column+1):
    v = ws.cell(5, c).value
    if v and 'junio' in str(v).lower() and '26' in str(v):
        print(c, ws.cell(5,c).column_letter, v)
"
```

Para USD fijo (header en fila 3):
```python
python -c "
import openpyxl
wb = openpyxl.load_workbook('datos/DEUDORES FINK-4-2026.xlsx', data_only=True)
ws = wb['\$  USD fijo']
for c in range(1, ws.max_column+1):
    v = ws.cell(3, c).value
    if v and 'junio' in str(v).lower() and '26' in str(v):
        print(c, v)
"
```
