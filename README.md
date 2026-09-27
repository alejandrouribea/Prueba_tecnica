# Procesamiento de ventas — `procesar_ventas.py`

Este script lee `datos_ventas.xlsx`, valida y limpia los datos, calcula totales de ventas por vendedor y por mes, y genera `resumen_ventas.xlsx` con dos hojas: `Resumen_Ventas` y `Ventas_Mensuales`.

## Requisitos

```bash
pip install pandas openpyxl
```

## Cómo ejecutarlo

```bash
python procesar_ventas.py
```

Debe ejecutarse en la misma carpeta donde está `datos_ventas.xlsx`. Si el archivo no existe o le faltan columnas, el script termina con un mensaje de error (`sys.exit(1)`) en lugar de continuar con datos incompletos.

## Constantes iniciales

```python
ARCHIVO_ENTRADA = "datos_ventas.xlsx"
ARCHIVO_SALIDA = "resumen_ventas.xlsx"
ANIO_OBJETIVO = 2023
COLUMNAS_REQUERIDAS = [...]
```

Centralizan los nombres de archivo, el año a filtrar y las columnas mínimas que debe tener el Excel de entrada, para no repetir literales sueltos por el código.

## Paso a paso del flujo (`main`)

`main()` ejecuta las funciones en orden, cada una recibe el DataFrame de la anterior y devuelve uno nuevo, ya limpio o transformado:

### 1. `cargar_datos`
Verifica que el archivo exista (`Path.exists`) y lo carga con `pd.read_excel`. Luego comprueba que estén presentes las 7 columnas requeridas (`ID_Venta`, `Fecha`, `Producto`, `Cantidad`, `Precio_Unitario`, `Total_Venta`, `Vendedor`). Si falta alguna, lanza `ValueError` antes de procesar nada.

### 2. `convertir_fecha`
Convierte la columna `Fecha` a tipo `datetime` con `pd.to_datetime(..., errors="coerce")`. Las fechas que no se puedan interpretar quedan como `NaT` y esas filas se descartan, porque sin fecha válida no se puede filtrar por año ni calcular el mes.

### 3. `normalizar_vendedor`
Convierte `Vendedor` a texto y le quita espacios en los bordes. Si queda vacío (celda nula, `None` o solo espacios), lo reemplaza por la etiqueta `"Sin Vendedor"` en vez de dejarlo en blanco o nulo.

Esto existe porque `pandas` descarta por defecto las filas con clave `NaN` al hacer `groupby`: si un vendedor llegara vacío y no se etiquetara, esa venta desaparecería silenciosamente del resumen por vendedor (aunque seguiría contando en el resumen mensual), y los dos totales dejarían de coincidir entre sí sin ningún aviso. Al etiquetarlo, la fila se sigue viendo — como una categoría `"Sin Vendedor"` explícita — y ambos resúmenes cuadran.

### 4. `normalizar_numero` (función auxiliar, no un paso del flujo)
No se llama directamente en `main`; la usan `validar_numericos` y `completar_total_venta` a través de `.map()` sobre cada celda, antes de convertirla con `pd.to_numeric`. Interpreta números escritos como texto con coma y/o punto:

- Si el valor ya es `int`/`float`, o está vacío/`NaN`, lo deja igual.
- Si tiene coma **y** punto (`"1.234,56"` o `"1,234.56"`): el separador más a la derecha (`str.rfind`) es el decimal; el otro se elimina por ser de miles.
- Si solo tiene coma (`"3,5"`): se asume convención española y se convierte en el punto decimal → `"3.5"`.
- Si tiene más de un punto y ninguna coma (`"12.345.678"`): todos son de miles, se eliminan.
- Si tiene un único punto (`"1.234"` vs `"3.5"`): si quedan exactamente 3 dígitos después del punto se asume separador de miles (`"1.234"` → `"1234"`); con 1, 2 o 4+ dígitos se respeta como decimal.

Esta heurística no es infalible (un `"34.255"` que en realidad quería decir "34 coma 255" se leerá como 34255), pero cubre los formatos más comunes de captura manual en español sin necesitar una librería de parseo de locales.

### 5. `validar_numericos`
Aplica `normalizar_numero` y luego `pd.to_numeric(..., errors="coerce")` a `Cantidad` y `Precio_Unitario`. Descarta las filas que sigan sin ser numéricas después de normalizar, y luego descarta las que queden con cantidad o precio menor o igual a 0 (una venta no puede tener cantidad o precio nulo/negativo).

### 6. `completar_total_venta`
Aplica `normalizar_numero` + `pd.to_numeric` también a `Total_Venta`, distinguiendo dos casos antes de tocar nada:
- Si la celda ya era nula (`NaN` real, sin valor), se marca para completar.
- Si tenía un valor pero no se pudo convertir a número ni siquiera normalizado, se avisa aparte (`"tenían un valor no numérico y se recalcularán"`) porque es un dato corrupto, no simplemente ausente.

Luego calcula `Cantidad * Precio_Unitario` y rellena con ese resultado únicamente las filas marcadas (nulas originales + no parseables). Si un `Total_Venta` válido no coincide con `Cantidad * Precio_Unitario`, no lo sobrescribe (para no alterar un dato ya cargado correctamente), solo imprime una advertencia para revisión manual.

### 7. `filtrar_por_año`
Filtra el DataFrame para quedarse solo con las filas cuyo año de `Fecha` sea `2023`. Si tras el filtro no queda ninguna fila, lanza `ValueError`, porque continuar generaría un reporte vacío sin avisar.

### 8. `agregar_columna_mes`
Crea la columna `Mes` extrayendo el número de mes (`1`-`12`) de la columna `Fecha` ya convertida a `datetime`.

### 9. `calcular_resumen_por_vendedor`
Agrupa por `Vendedor` y suma `Total_Venta`, ordenando de mayor a menor. Es el contenido de la hoja `Resumen_Ventas`. Gracias a `normalizar_vendedor`, las filas sin vendedor quedan agrupadas bajo `"Sin Vendedor"` en vez de perderse.

### 10. `calcular_resumen_mensual`
Agrupa por `Mes` y suma `Total_Venta`, ordenando de enero a diciembre. Es el contenido de la hoja `Ventas_Mensuales`.

### 11. `exportar_resumen`
Usa `pd.ExcelWriter` con motor `openpyxl` para escribir ambos resúmenes como hojas distintas dentro de un mismo archivo `resumen_ventas.xlsx`.

### Manejo de errores en `main`
Todo el flujo está dentro de un `try/except` que captura `FileNotFoundError` y `ValueError` (los errores esperables de datos/archivo), imprime el mensaje en `stderr` y termina con código de salida `1`, en vez de mostrar un traceback completo al usuario final.

## Validaciones incluidas (resumen)

| # | Validación | Por qué |
|---|---|---|
| 1 | Columnas requeridas presentes | Sin ellas el resto del script fallaría de forma poco clara |
| 2 | `Fecha` convertible a `datetime` | Necesaria para filtrar por año y extraer el mes |
| 3 | `Vendedor` vacío se etiqueta como `"Sin Vendedor"` (no se descarta) | Evita que `groupby` la elimine en silencio y descuadre `Resumen_Ventas` contra `Ventas_Mensuales` |
| 4 | `Cantidad` y `Precio_Unitario` normalizados (coma/punto) antes de convertir | Evita perder o mal-interpretar números capturados como texto en formato español o inglés |
| 5 | `Cantidad` y `Precio_Unitario` numéricos y positivos | Un valor no numérico o ≤ 0 invalidaría el cálculo del total |
| 6 | `Total_Venta` normalizado (coma/punto) antes de comparar/completar | Evita que un texto no convertible quede mezclado con números en la misma columna y rompa el `.sum()` del resumen |
| 7 | `Total_Venta` faltante o no parseable se completa con `Cantidad * Precio_Unitario` | Requisito explícito del enunciado, extendido a valores corruptos |
| 8 | `Total_Venta` inconsistente se reporta (no se sobrescribe) | Alerta de posibles errores de carga sin perder el dato original |
| 9 | Filtro por año 2023 | Requisito explícito del enunciado |
| 10 | DataFrame no vacío tras filtrar | Evita generar un Excel de salida vacío sin aviso |

## Salida esperada

Con el estado actual de `datos_ventas.xlsx` (100 filas, 25 con `Total_Venta` faltante y 21 con `Vendedor` vacío), el script:
- Etiqueta las 21 filas sin vendedor como `"Sin Vendedor"`.
- Completa los 25 valores faltantes de `Total_Venta`.
- No descarta ninguna fila (todas las fechas, cantidades y precios son válidos, y todas pertenecen a 2023).
- Genera `resumen_ventas.xlsx` con:
  - `Resumen_Ventas`: 5 filas (Ana, Carlos, `Sin Vendedor`, David, Beatriz).
  - `Ventas_Mensuales`: 12 filas (una por mes), cuyo total coincide con el de `Resumen_Ventas`.
