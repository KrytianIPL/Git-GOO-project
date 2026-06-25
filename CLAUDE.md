# Proyecto Reabastecimiento ABC · Mezzanine · Cliente GOO

Esta carpeta contiene un sistema que genera tres tableros HTML a partir de dos
archivos Excel de inventario. Este archivo te dice (a Claude) cómo manejar el
proyecto cuando trabajes en esta carpeta desde Cowork.

## Qué hace el proyecto

A partir de dos Excel, un script de Python genera tres archivos HTML:
1. **Dashboard_Reabastecimiento_TipoA_GOO.html** — tablero interactivo para
   presentar a jefatura (con gráficos, KPIs y proyección de tiempo).
2. **Informe_Reabastecimiento_GOO_correo.html** — versión estática, sin
   JavaScript, lista para adjuntar a un correo.
3. **Herramienta_Reabastecimiento_GOO.html** — herramienta operativa por
   corrida, para que el personal de piso sepa qué cajas bajar y a dónde.

## Cómo regenerar (la tarea principal)

Cuando el usuario actualice los archivos Excel y pida regenerar, ejecuta:

```
python generar_dashboard.py
```

(en Windows puede ser `py generar_dashboard.py` si `python` no funciona).

Esto lee los dos Excel y reescribe los tres HTML en esta misma carpeta. No hay
que descargar nada: los archivos quedan actualizados aquí.

## Archivos de la carpeta

- `generar_dashboard.py` — el script generador (NO modificar sin pedirlo).
- `plantilla_dashboard.html` — plantilla del dashboard interactivo.
- `plantilla_herramienta.html` — plantilla de la herramienta operativa.
- `Para_analisis_de_GOO_proyecto.xlsx` — INVENTARIO. Columnas: Producto,
  Barcode, Categoria, Tipo (ABC), Ubicación, Serial EBS, CID, Saldo, Area.
- `Analisis_de_progreso_ABC_GOO.xlsx` — PENDIENTE POR PICK (hoja "Pendiente
  por Pick"). Columnas: Producto, Tipo (ABC), Ubic Pick, Pendiente, Ubic,
  Area, Pedido, OID, Pais, Categoría, Descripción.
- `corte_anterior.json` — lo crea el script solo. Guarda los totales y la
  FECHA del último corte, para calcular el avance y la proyección de tiempo.
  **NO BORRAR**: si se borra, se pierde el historial de avance.

## Para actualizar con datos nuevos

1. Reemplaza los dos Excel en esta carpeta por los nuevos (mismo nombre).
2. Pide regenerar (ejecutar `generar_dashboard.py`).
3. Los tres HTML quedan actualizados en la carpeta.

No cambies los nombres de los Excel; el script los busca por su nombre exacto.

## Definiciones del negocio (IMPORTANTES — no cambiar sin confirmar)

- **PISO** = una ubicación cuenta como piso si:
  - su Área es `MEZZANINE BINS`, O BIEN
  - su ubicación empieza con `MZ06` y termina en `-01` (estas físicamente ya
    están abajo, aunque el sistema las marque como RACK).
- **ALTURA** = todo lo demás (RACK que no sea MZ06-...-01, Paleta, etc.).
  Es de donde se baja producto para reabastecer.
- **Estándar de piso:** Tipo A = 2 cajas en piso; Tipo B = 1 caja; Tipo C =
  todo en altura (no va a piso).
- El conteo de reabastecimiento es por **cajas (CID distintos)**, no por
  unidades. Cualquier caja cuenta como válida, sin importar cuántas unidades
  tenga.
- **Corrida** = estilo + color = el SKU sin la última talla (lo que está
  después del último guion). Objetivo: todas las tallas de una corrida juntas
  en una sola ubicación de piso.

## Regla de selección de cajas (herramienta)

- Si UNA sola ubicación de altura tiene la corrida completa (suficientes cajas
  de todas las tallas que faltan), se bajan de ahí, aunque no sean las de más
  unidades (para no romper la agrupación).
- Si la corrida está dispersa, se bajan las cajas con MÁS unidades.

## Destino de cada corrida (herramienta)

1. Si la corrida ya tiene UNA ubicación en piso → llevar ahí ("fija").
2. Si está dispersa en varias de piso → consolidar en la que más tallas tiene.
3. Si no tiene nada en piso → el operario asigna una ubicación libre y la
   herramienta la recuerda ("nueva").

## Proyección de tiempo

El dashboard estima cuándo se llegará al 100% del reabastecimiento, midiendo
el ritmo real (cajas bajadas ÷ días entre cortes). La primera vez muestra un
aviso de espera; a partir del segundo corte con avance, muestra la proyección.

## Importante al modificar el código

Si hay que corregir algo de los HTML generados, el cambio debe hacerse en el
SCRIPT o en las PLANTILLAS, no solo en el HTML resultante. Si se parchea solo
el HTML, el cambio se pierde la próxima vez que se regenere.
