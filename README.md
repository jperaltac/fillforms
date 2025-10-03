
# Generador de Formularios ANID 2026 (F1 & F2) desde CSV

Este mini-proyecto crea **dos archivos .docx por estudiante** a partir de tus plantillas y de un CSV con los datos.

## Requisitos

- Python 3.9+
- Paquetes:
  ```bash
  pip install python-docx pandas
  ```

## Archivos incluidos

- `generate_forms.py` — script principal
- `estudiantes_template.csv` — plantilla de columnas para tu CSV
- **Tus plantillas** (ya subidas por ti):
  - `Formulario_N1_2026 (1).docx`
  - `Formulario_N_2_2026.docx`

> **Tip:** Mantén los nombres exactos o pásalos con `--f1` y `--f2` si los cambias.

## CSV esperado

Abre `estudiantes_template.csv` y rellena una fila por estudiante.
Columnas:

- nombre, apellido, run, pasaporte
- universidad_pregrado, programa_pregrado, semestres_pregrado, region_pregrado
- promedio_pregrado, nota_final, posicion_egreso, total_generacion, ranking_porcentaje
- estado_postulacion (valores: `postulacion_formal` | `aceptado` | `alumno_regular`)
- programa_destino, mencion_destino, universidad_destino, region_postgrado, fecha_inicio_postgrado
- autoridad_nombre, autoridad_cargo

> Puedes dejar en blanco lo que no tengas. Para RUN/pasaporte, se usa primero `run` y si está vacío se usa `pasaporte`.

## Uso

Coloca tu CSV en la misma carpeta y ejecuta:

```bash
python generate_forms.py --csv estudiantes.csv
```

Se crearán archivos en `./salida` con el formato:

```
Apellido_Nombre_F1.docx
Apellido_Nombre_F2.docx
```

### Opciones

- `--f1 "Formulario_N1_2026 (1).docx"` — ruta a plantilla Formulario N°1
- `--f2 "Formulario_N_2_2026.docx"` — ruta a plantilla Formulario N°2
- `--outdir salida` — carpeta de salida

## ¿Qué rellena exactamente?

- **Formulario N°1** (Notas y Ranking): nombre, RUN/pasaporte, universidad y programa de pregrado, semestres, región, promedio, nota final, posición, total generación y ranking.
- **Formulario N°2** (Estado del/de la postulante): nombre, RUN/pasaporte, **marca el estado** (checkbox simulado con `[X]`), programa/mención/universidad de destino, región de postgrado, fecha de inicio y autoridad (nombre, cargo).

> El marcado del estado se hace por texto: `postulacion_formal`, `aceptado` o `alumno_regular`.

## Cómo funciona (rápido)

El script busca **líneas que empiezan con etiquetas** exactas del documento y las reescribe con el valor al final. Si una etiqueta no se encuentra (plantilla cambió, espacios, etc.), verás un aviso `[WARN]` pero el script seguirá.

Si quieres máxima robustez, otra alternativa es usar `docxtpl` con placeholders `{{...}}`. Este script evita esa edición previa y funciona **tal cual** con tus plantillas actuales.

---

Hecho con cariño para agilizar tu flujo ANID.
