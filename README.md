# Limpiador de Excel/CSV

Script en Python que recibe un Excel o CSV desordenado y devuelve un `.xlsx` limpio.
Todos los datos de ejemplo son **inventados**.

## Qué hace

- Normaliza cabeceras (`" Fecha alta "` → `fecha_alta`) y elimina filas/columnas vacías.
- Nombres y ciudades: quita espacios sobrantes y pone mayúsculas correctas.
- Emails: minúsculas, sin espacios; avisa si no tienen formato válido.
- Teléfonos: formato `600 123 456` (quita `+34`, guiones y puntos); avisa si la longitud no cuadra.
- Fechas: entiende `05/03/2024`, `2024-03-07`, `5-3-2024`, fechas de Excel y números de serie; avisa si son imposibles (`31/02/2024`).
- Importes: entiende `1.234,50 €`, `89 EUR`, `1 200,00`, `99.99`; avisa si no son números.
- Duplicados: conserva la primera aparición (por email, o fila completa si no hay email).
- Las filas con problemas salen en amarillo con una columna `avisos`, y una hoja `Resumen` cuenta lo que se ha cambiado.
- No modifica el archivo original.

## Uso

```bash
pip install pandas openpyxl

python crear_ejemplo.py                      # genera ejemplo_sucio.xlsx
python limpiar_excel.py ejemplo_sucio.xlsx   # genera ejemplo_sucio_limpio.xlsx
python limpiar_excel.py datos.csv -o salida.xlsx
python limpiar_excel.py datos.xlsx --hoja "Clientes"
```

## Limitaciones conocidas

- Asume formato español: fechas día/mes/año, teléfonos de 9 dígitos, coma decimal.
- `2.500` se interpreta como 2500 (miles). Si tus datos usan punto decimal con tres decimales, revísalo.
- Reconoce el tipo de columna por su nombre (email, teléfono, fecha, importe...). Otras columnas solo se limpian de espacios.
- No corrige erratas ni inventa datos que faltan: los marca para revisión manual.
