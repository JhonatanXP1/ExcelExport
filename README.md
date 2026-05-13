# Implementación de Exportable de Excel

Genera archivos `.xlsx` de forma dinámica mediante una solicitud `POST` con un JSON estructurado.

**Path dentro del proyecto:** `EDUCAPP_v25/exportExcel.php`

---

## Requisitos

- Solo acepta solicitudes `POST`. Cualquier otro método retorna `401 Unauthorized`.
- Soporta preflight CORS (`OPTIONS`) para integraciones desde el front-end.
- Si el JSON viene codificado, incluir `"encode": true` para que el servidor aplique la decodificación UTF-8 antes de procesarlo.

---

## Estructura del JSON

### Propiedades raíz

| Propiedad  | Tipo     | Requerido | Descripción |
|------------|----------|-----------|-------------|
| `nameFile` | `string` | Sí        | Nombre base del archivo. Se le agrega automáticamente la fecha actual (`YYYY-MM-DD`) y la extensión `.xlsx`. |
| `title`    | `string` | Sí        | Texto del encabezado general que aparece en la primera fila, abarcando todas las columnas. Acepta saltos de línea con `\n`. |
| `data`     | `array`  | Sí        | Contenido del Excel. Puede ser asociativo (múltiples hojas) o no asociativo (una sola hoja). |
| `indice`   | `bool`   | No        | Si es `true`, agrega una columna de numeración automática al inicio de la tabla. Por defecto: `false`. |
| `fillter`  | `bool`   | No        | Si es `true`, activa el autofiltro de Excel en la fila de encabezados. Se desactiva automáticamente si alguna celda usa `mergeX`. Por defecto: `false`. |
| `encode`   | `bool`   | No        | Si es `true`, aplica decodificación UTF-8 al JSON recibido antes de procesarlo. |

---

## Hojas del Excel (`data`)

### Una sola hoja (array no asociativo)

Cuando `data` contiene un único bloque de datos con índice numérico, se genera un Excel con una sola hoja.

```php
[
    "nameFile" => "Reporte_Alumnos_",
    "title"    => "Reporte de Alumnos\nGrado: 1A | Nivel: Secundaria | Ciclo Escolar: 2025/1",
    "data"     => [
        [
            ["Nombre", "Matrícula", "Edad", "Estatus"],
            ["Ana Sofía Ramírez López", "RAM-234A9y", "15", "Activo"]
        ]
    ]
]
```

### Múltiples hojas (array asociativo)

Cuando `data` es asociativo, cada clave define el nombre de una hoja distinta dentro del Excel.

```php
[
    "nameFile" => "Reporte_Alumnos_",
    "title"    => "Reporte de Alumnos\nNivel: Secundaria | Ciclo Escolar: 2025/1",
    "data"     => [
        "1° A" => [
            ["Nombre", "Matrícula", "Edad", "Estatus"],
            ["Luis Miguel Torres García", "TOR-456B3z", "16", "Activo"]
        ],
        "2° B" => [
            ["Nombre", "Matrícula", "Edad", "Estatus"],
            ["Ana Sofía Ramírez López", "RAM-234A9y", "15", "Activo"]
        ]
    ]
]
```

El Excel resultante tendrá:
- **Hoja 1:** `1° A` con su información.
- **Hoja 2:** `2° B` con su información.

> Los nombres de hoja se sanitizan automáticamente: se eliminan espacios, se transliteran caracteres especiales y se truncan a 30 caracteres.

---

## Encabezados de tabla

El **primer array** dentro de cada bloque de datos se interpreta como la fila de encabezados de la tabla. Todos los arrays siguientes son el cuerpo.

```php
"1° A" => [
    // ↓ Encabezados
    ["Nombre", "Matrícula", "Edad", "Estatus"],
    // ↓ Cuerpo
    ["Luis Miguel Torres García", "TOR-456B3z", "16", "Activo"],
    ["Ana Sofía Ramírez López",   "RAM-234A9y", "15", "Activo"]
]
```

### Estilo de encabezado

Los encabezados se renderizan automáticamente con:
- Texto en **negrita**, tamaño 16, color blanco.
- Fondo azul corporativo (`#4472C4`).
- Texto centrado horizontal y verticalmente.

### Propiedad especial para encabezados: `mergeX`

Fusiona la celda con las columnas siguientes. El valor indica cuántas columnas **adicionales** se incluyen en la unión.

> El valor del encabezado debe ser siempre el primer elemento del array de la celda.

```php
["Nombre", "Matrícula", "Edad", ["Estatus", "mergeX" => 1]]
// "Estatus" ocupará 2 columnas (la propia + 1 adicional)
```

> **Nota:** Usar `mergeX` en encabezados desactiva el autofiltro (`fillter`) automáticamente, ya que Excel no admite filtros con celdas combinadas en la fila de encabezado.

---

## Cuerpo del Excel

Todos los arrays a partir del segundo son filas del cuerpo. Cada elemento del array representa una celda.

Las filas alternan color de fondo automáticamente (efecto zebra):
- Filas pares: `#FAFAFA` con texto `#181818`.
- Filas impares: `#C9D6EE` con texto `#303030`.

### Celdas simples

Un valor escalar (`string`) se escribe directamente en la celda.

```php
["Ana Sofía Ramírez López", "RAM-234A9y", "15", "Activo"]
```

### Celdas con propiedades especiales

Cuando una celda necesita estilo o comportamiento extra, se define como array. El **primer elemento** es siempre el valor de la celda.

| Propiedad   | Tipo     | Descripción |
|-------------|----------|-------------|
| `mergeX`    | `int`    | Fusiona la celda con las columnas siguientes. El valor indica cuántas columnas adicionales se incluyen. |
| `strong`    | `bool`   | Si es `true`, aplica negrita al texto. |
| `color`     | `string` | Color del texto en hexadecimal sin `#` (ej. `"FF0000"` para rojo). |
| `fillColor` | `string` | Color de fondo de la celda en hexadecimal sin `#` (ej. `"FFFF00"` para amarillo). |
| `align`     | `string` | Alineación horizontal del texto. Valores: `"left"` o `"right"`. Por defecto: centrado. |
| `currency`  | `mixed`  | Formatea la celda como moneda MXN (`$ #,##0.00`). El valor se almacena como número para permitir cálculos en Excel. |
| `indice`    | `bool`   | Si es `false` y la opción global `indice` está activa, suprime el número de fila para esa celda específica y no lo cuenta en el conteo. |

### Ejemplo completo del cuerpo

```php
[
    "Carlos Jhonatan Díaz Méndez",
    "JHO-118S7x",
    "17",
    [
        "Activo",
        "color"    => "0FFF00"   // Texto verde
    ],
    [
        "México / Puebla",
        "mergeX"   => 1,         // Ocupa 2 columnas
        "strong"   => true       // Texto en negrita
    ],
    "Masculino",
    [
        "70 kg.",
        "fillColor" => "FFF3CC"  // Fondo amarillo claro
    ],
    [
        "1500.50",
        "currency"  => true      // Formato $ 1,500.50
    ],
    [
        "Notas adicionales",
        "align"     => "left"    // Texto alineado a la izquierda
    ]
]
```

---

## Numeración automática (`indice`)

Cuando `"indice": true` se incluye en el JSON raíz, se agrega una columna `No.` al inicio de la tabla que numera cada fila automáticamente.

Para omitir el número en una fila específica (por ejemplo, una fila de subtotal o separador), se incluye `"indice": false` en cualquier celda de esa fila:

```php
[
    ["Subtotal", "mergeX" => 3, "indice" => false],
    "4500.00"
]
// Esta fila no mostrará número y no contará en el índice.
```

---

## Autofiltro (`fillter`)

Cuando `"fillter": true` se incluye en el JSON raíz, Excel agrega controles de filtro en la fila de encabezados, permitiendo filtrar los datos directamente desde la aplicación.

> El autofiltro se desactiva automáticamente si cualquier celda del encabezado o del cuerpo usa `mergeX`, ya que Excel no admite filtros con celdas combinadas.

---

## Comportamiento automático del archivo generado

- El nombre final del archivo es: `{nameFile}{YYYY-MM-DD}.xlsx`
- Las columnas se ajustan automáticamente al contenido (`autoSize`).
- La fila de encabezados de tabla queda **congelada** (freeze pane en fila 3), permitiendo desplazarse por los datos sin perder de vista los encabezados.
- El encabezado global (`title`) ocupa la fila 1 y abarca todas las columnas de la tabla.
- Los metadatos del archivo (autor, modificado por, tema) se establecen automáticamente como `"Ing. Jhonatan"`.

---

## Manejo de errores

Si ocurre un error durante la generación, el servidor responde con `HTTP 500` y un mensaje descriptivo en el cuerpo de la respuesta:

```
error en el servidor: {código} - {mensaje}
```

Los errores más comunes son:
- JSON vacío o mal formado.
- `data` nulo o sin contenido.
- Excepciones internas de PhpSpreadsheet.

---

## Ejemplo de uso completo

```php
[
    "nameFile" => "Reporte_Alumnos_",
    "title"    => "Reporte de Alumnos\nNivel: Secundaria | Ciclo Escolar: 2025/1",
    "indice"   => true,
    "fillter"  => true,
    "data"     => [
        "1° A" => [
            ["Nombre", "Matrícula", "Edad", "Estatus", "Observaciones"],
            ["Luis Miguel Torres García", "TOR-456B3z", "16", "Activo", ""],
            ["Ana Sofía Ramírez López",   "RAM-234A9y", "15", "Activo", ""],
            [
                ["Totales", "mergeX" => 3, "strong" => true, "indice" => false],
                ["2 alumnos", "color" => "4472C4"]
            ]
        ]
    ]
]
```
