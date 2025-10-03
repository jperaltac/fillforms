# Rellenar documentos DOCX a partir de un CSV

Este proyecto toma un archivo CSV y una plantilla DOCX con marcadores del tipo
`[[Campo]]` para generar un documento por cada fila del CSV.

Los nombres de columna se comparan con los marcadores sin distinguir mayúsculas
o minúsculas y también se ignoran los `#` iniciales. Por ejemplo, si tu CSV
incluye la columna `#Nombres`, podrás usar `[[NOMBRES]]`, `[[nombres]]` o
`[[Nombres]]` dentro del documento.

## Requisitos

- Python 3.9+
- Dependencias:
  ```bash
  pip install python-docx
  ```

## Uso rápido

```bash
python generate_forms.py alumnos.csv plantilla.docx --name-column Nombres
```

- `alumnos.csv`: archivo con encabezados en la primera fila.
- `plantilla.docx`: documento con los marcadores `[[...]]` que deseas reemplazar.
- `--name-column`: (opcional) nombre de la columna usada para generar el nombre
  de cada archivo. Si se omite, se utilizará el primer valor no vacío de la
  fila o, como último recurso, `row_001.docx`, `row_002.docx`, etc.
- `--outdir`: (opcional) carpeta de salida. Por defecto se usa `salida/`.
- `--encoding`: (opcional) codificación del CSV. Por defecto `utf-8-sig`.

Cada fila genera un archivo DOCX dentro del directorio de salida. Si el nombre
resultante ya existe, el script añade un sufijo incremental (`_1`, `_2`, ...).

## Preparar la plantilla

Escribe tus marcadores dentro de dobles corchetes. Algunos ejemplos:

- `Nombre completo: [[Nombres]] [[Apellidos]]`
- `RUN: [[rut]]`
- `Programa: [[ Programa ]]`  ← los espacios también se ignoran al comparar.

Dentro de la carpeta [`examples/`](examples/) encontrarás una plantilla de
ejemplo (`template.docx`) y el CSV correspondiente (`students.csv`). Puedes
probar el flujo completo con:

```bash
python generate_forms.py examples/students.csv examples/template.docx --name-column Nombres --outdir examples/output
```

## Notas

- Las filas completamente vacías del CSV se ignoran.
- Los valores se insertan tal cual aparecen en el CSV. Si necesitas formato
  adicional (por ejemplo, fechas), prepáralo previamente en el CSV.
- La codificación `utf-8-sig` permite abrir archivos guardados desde Excel que
  incluyen la marca BOM. Ajusta la opción `--encoding` si tu CSV usa otra
  codificación.

