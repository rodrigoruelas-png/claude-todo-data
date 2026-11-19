# Guía paso a paso: extraer datos de Noetix (para quien no conoce Noetix)

Quién hace qué: **tú** corres la extracción en tu PC (Noetix solo responde en tu red y con tu usuario).
Claude analiza el archivo que le pases. **Nunca compartas tu usuario/contraseña con nadie, ni con Claude.**

## Parte A — Actualizar el Excel que ya tienes (`Noetix_Tescom_Automatico_v4`)
1. Abre el Excel en tu PC (dentro de la red de la empresa o con VPN). Si pide "Habilitar contenido/edición", acepta.
2. Hoja **LOGIN**: escribe tu usuario en B2 y tu contraseña en B3.
3. Hoja **PARAMETROS** (celdas amarillas): deja la organización `ASCO_USA_IO_MC1`.
   - Para una **prueba rápida**: `Dias hacia atras` = 30 (rango de fechas de los últimos 30 días).
   - Para el **histórico**: 319, pero ver la nota de abajo.
4. Menú **Datos › Actualizar todo**. Tarda varios minutos; no cierres Excel. Si pregunta usuario/contraseña
   o responsabilidad, usa la de solo lectura (ver abajo).
5. Verifica que ya no diga "Sin datos todavia" en la fila 2 de cada hoja.
6. **Antes de enviarlo:** en LOGIN **borra B2 y B3**, guarda como `.xlsx` y súbelo a la sesión.

**Responsabilidad a usar (solo lectura):** `EMR NOETIX Global, Standard` → NOETIX Reporting. Si te pide
"Responsibility", escribe `%` (como indica el propio mensaje de la consulta) o esa responsabilidad.

**Nota — posible truncamiento:** en tu v4 completo, WIP_Completions trae ~124,989 filas y la última fecha es
10-sep aunque el rango llega al 17-nov. Es una **hipótesis**: el gateway podría cortar a ~125 mil filas.
Si pasa, divide el histórico en tramos (p. ej. ene–jun y jul–nov) o usa el archivo "parte 2" (últimos 30 días).

## Parte B — Las 3 consultas nuevas que faltan
Se crean en la página web de Noetix (`http://noetix.emersonprocess.com/nwq/nwq.htm`, abre en Edge dentro de tu red).
No conozco tu pantalla exacta: la lógica es **elegir vista → elegir columnas → poner filtro → ejecutar → exportar a Excel**.
Si un nombre no aparece, tómale captura a la lista de vistas y mándamela (sin datos sensibles).

1. **Existencias con fecha de recepción**
   - Vista: `INVG0_Onhand_Quantities` (la misma de OnHand).
   - Columnas: las de OnHand **más** cualquier columna tipo *Date Received / Orig Date Received / Receipt Date*.
   - Filtro: Organization_Name = `ASCO_USA_IO_MC1`.
   - Si la vista no tiene fecha de recepción, dímelo: se busca otra vista o se usa el `MC1 ONHAND Aging Report`.
2. **Último movimiento por P/N (todos los tipos)**
   - Vista: `INVG0_Transactions`.
   - Columnas: Organization/ORG, Item Number, Transaction Date, Transaction Type, Transaction Quantity, Subinventory.
   - Filtro: ORG = MC1 y fecha **entre** (hoy − 365 días) y hoy. **Sin filtro de tipo de transacción.**
   - Si pasa de ~125 mil filas, ejecútalo por tramos de fechas.
3. **Requerimientos de componentes de WIP**
   - Busca en la lista de vistas las que empiecen con `WIPG0_` y contengan *Requirement* o *Material*.
   - Columnas: Job Name, Job Status, Item Number del componente, Required Quantity, Quantity Issued, Subinventory.
   - Filtro: Org = MC1 y Job Status = Released.

Guarda cada una como `.xlsx` o `.csv` (nombre: `noetix_onhand_recepcion.xlsx`, `noetix_ult_mov.xlsx`,
`noetix_requerimientos_wip.xlsx`) y súbelas a esta sesión.

## Parte C — Lo que me debes pasar para cerrar el cruce con WIPREG
- El archivo/correo de Yahir (WIPREG): basta una hoja con ~20 filas de ejemplo, sin datos sensibles.
- El Excel `MC1 ONHAND Aging Report` más reciente (o sus columnas).

## Parte D — Qué haré yo con esos archivos
Calcular antigüedad real, P/N sin movimiento (por tipo de ítem), ítems críticos (requerimiento > existencia
o existencia < mínimo) y armar la gráfica de nivel de inventario + tabla de conciliación Oracle vs WIPREG vs Aging.
