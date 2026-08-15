# Consolidador de consumos eléctricos

Herramienta de línea de comandos que unifica en un único CSV los registros horarios de consumo procedentes de las facturas descargadas desde el portal de tu distribuidora eléctrica.

---

## Requisitos

- Python 3.8 o superior
- Sin dependencias externas (solo librería estándar)

---

## Estructura de carpetas

```
consolidador_electricidad/
├── consolidar.py          ← script principal
├── input/                 ← coloca aquí tus CSV de facturas
│   └── _duplicados/       ← los duplicados detectados se mueven aquí
└── output/
    └── consolidado_consumo.csv   ← resultado generado
```

Las carpetas `input/` y `output/` deben existir antes de ejecutar el script. Si `output/` no existe, se crea automáticamente.

---

## Uso

1. Descarga tus facturas CSV desde el portal de tu distribuidora y cópialas en la carpeta `input/`.
2. Ejecuta el script desde la raíz del proyecto:

```bash
python consolidar.py
```

3. El archivo resultante estará en `output/consolidado_consumo.csv`.

---

## Formato de los CSV de entrada

El script espera el formato estándar exportado por distribuidoras como Endesa / e-distribución:

```
CUPS:,ES00316...
Fecha inicio:,18/04/2026
Fecha fin:,17/05/2026
Fecha y hora de extracción:,21/07/2026 23:43:31
Tarifa:,24T2D - TEMPO 24 HORAS 2.0TD

Fecha,Hora,Consumo (Wh),Precio (€/kWh),Coste por hora (€)
2026-04-18,00:00-01:00,104,0,0
2026-04-18,01:00-02:00,52,0,0
...
,Total (Wh):,320194
```

Características relevantes:
- Las 6 primeras filas son metadatos; los datos horarios empiezan en la fila 7.
- La última fila contiene el total y se ignora automáticamente.
- El separador de los datos es coma (`,`); las fechas de los registros están en formato `YYYY-MM-DD`.
- Las fechas de inicio/fin en la cabecera están en formato `DD/MM/YYYY`.

---

## Formato del CSV de salida

El archivo `consolidado_consumo.csv` usa punto y coma (`;`) como separador y coma (`,`) como separador decimal, por compatibilidad con Excel en configuración regional española.

| Columna | Ejemplo | Descripción |
|---|---|---|
| `Fecha` | `18/04/2026` | Fecha del registro en formato DD/MM/YYYY |
| `Hora` | `00:00 - 01:00` | Franja horaria |
| `Consumo_Wh` | `104,00` | Consumo en vatios-hora |
| `Archivo_origen` | `factura_49_...csv` | Nombre del CSV de origen |

Los registros se ordenan de más reciente a más antiguo.

---

## Detección de duplicados

Antes de consolidar, el script calcula una huella (SHA-256) de los datos de consumo de cada archivo, basada únicamente en `Fecha + Hora + Consumo`. Esto permite detectar archivos duplicados aunque tengan nombres distintos.

Cuando se detectan duplicados:
- Se conserva el archivo con la fecha de inicio más temprana (o, en caso de empate, el de nombre alfabéticamente menor).
- Los archivos duplicados se mueven a `input/_duplicados/` y no se eliminan, por si necesitas revisarlos.

---

## Ejemplo de salida en consola

```
============================================================
Consolidador de consumos eléctricos
============================================================

Archivos encontrados: 3
  OK  factura_49_20260418_20260517.csv
  OK  factura_49_COPIA.csv
  OK  factura_50_20260518_20260614.csv

Duplicados detectados: 1
  Duplicado movido: factura_49_COPIA.csv → input/_duplicados/factura_49_COPIA.csv

Consolidando 2 factura(s)...

============================================================
Proceso completado
  Facturas procesadas : 2
  Duplicados movidos  : 1
  Registros totales   : 1392
  Archivo generado    : .../output/consolidado_consumo.csv
============================================================
```

---

## Configuración avanzada

Al inicio de `consolidar.py` hay constantes que puedes modificar sin tocar el resto del código:

| Constante | Por defecto | Descripción |
|---|---|---|
| `INPUT_DIR` | `./input` | Carpeta de entrada |
| `OUTPUT_DIR` | `./output` | Carpeta de salida |
| `ARCHIVO_SALIDA` | `consolidado_consumo.csv` | Nombre del archivo de salida |
| `CARPETA_DUPLICADOS` | `_duplicados` | Subcarpeta dentro de `input/` para duplicados |
| `SEPARADOR_SALIDA` | `;` | Separador del CSV de salida |
| `DECIMAL_SALIDA` | `,` | Separador decimal del CSV de salida |
| `ENCODING` | `utf-8-sig` | Codificación (BOM para compatibilidad con Excel) |