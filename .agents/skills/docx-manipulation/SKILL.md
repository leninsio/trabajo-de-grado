---
name: docx-manipulation
description: Best practices for manipulating Word document (.docx) files using python-docx without layout regression or corruption.
---

# Manipulación Segura de Archivos Word (.docx) con python-docx

Esta guía documenta los aprendizajes y mejores prácticas metodológicas para editar y generar documentos de Word utilizando `python-docx` en el contexto de proyectos académicos y de tesis, evitando regresiones en el formato, pérdida de imágenes (como logos) o corrupción del archivo.

## 1. Patrones de Lectura e Inspección

Para evitar desalineaciones por índices variables, siempre inspecciona la estructura del documento de forma no destructiva antes de intentar aplicar cambios.

### Mapeo de Párrafos y Tablas en el Árbol del Documento
El orden en que aparecen los párrafos (`doc.paragraphs`) y las tablas (`doc.tables`) no siempre coincide con el orden visual si están anidados en celdas. Para identificar el flujo exacto de los elementos en el cuerpo del documento, recorre la colección de hijos del `body`:

```python
import docx

doc = docx.Document("documento.docx")
body = doc.element.body

for idx, child in enumerate(body):
    if child.tag.endswith('p'):
        p = docx.text.paragraph.Paragraph(child, doc)
        print(f"Index {idx} [Párrafo]: '{p.text[:60]}'")
    elif child.tag.endswith('tbl'):
        t = docx.table.Table(child, doc)
        print(f"Index {idx} [Tabla] con {len(t.rows)} filas")
```

---

## 2. Trampas y Errores Críticos (Gotchas)

### A. La Trampa de la Evaluación Dinámica en `insert_paragraph_before()`
Al insertar varios párrafos de forma consecutiva antes de un determinado punto, nunca uses una referencia dinámica de índice como `paragraphs[0]`.

* **El Error (LIFO Inversion)**:
  ```python
  # Esto causará que los párrafos se inserten en orden inverso (de abajo hacia arriba)
  p = cell.paragraphs[0]
  p.insert_paragraph_before().text = "Línea 1"
  p.insert_paragraph_before().text = "Línea 2"  # Se inserta antes de la nueva línea 1
  ```
* **La Solución (Referencia Estática)**:
  Guarda una referencia estática al párrafo original y realiza todas las inserciones sobre esa misma variable de referencia:
  ```python
  target_p = cell.paragraphs[0]
  p1 = target_p.insert_paragraph_before()
  p1.text = "Línea 1"
  p2 = target_p.insert_paragraph_before()
  p2.text = "Línea 2"  # Se coloca correctamente después de la línea 1 y antes del objetivo
  ```

### B. El Estiramiento de Espaciado por Justificación (`JUSTIFY`)
Si estás insertando líneas de completado (como `Nombre: _____________________`), nunca apliques justificación completa (`WD_ALIGN_PARAGRAPH.JUSTIFY`) al párrafo.
* **Consecuencia**: Word intentará estirar los espacios entre las palabras a lo largo de toda la página, separando de forma exagerada las letras (ej: `Nombre     :   ________`).
* **Solución**: Usa alineación izquierda (`WD_ALIGN_PARAGRAPH.LEFT`) para cualquier párrafo que contenga líneas de firma, completado o datos del trabajador.

### C. Preservación del Estilo de Lista (Viñetas / Bullets)
Al reemplazar texto en párrafos existentes, ten en cuenta que el párrafo conserva su estilo de Word original.
* **El Problema**: Si el párrafo original tenía viñetas (`•`), el nuevo texto seguirá teniendo viñetas automáticamente a menos que limpies el estilo.
* **La Solución**: Cambia explícitamente el estilo a `'Normal'` para eliminar viñetas no deseadas:
  ```python
  p.style = 'Normal'
  p.paragraph_format.left_indent = Inches(0.0)
  p.paragraph_format.first_line_indent = Inches(0.0)
  ```

### D. Ruptura de Relaciones de Medios al Copiar XML entre Documentos
Si un elemento (como una celda de tabla) contiene una imagen o logo, **no** intentes clonar o copiar su elemento XML directamente a un nuevo documento en blanco.
* **El Problema**: Las imágenes dependen de identificadores de relación (`rId`) almacenados en los metadatos del documento fuente. Si copias el XML a un archivo nuevo, el logo aparecerá como una imagen rota porque los metadatos de relación no existen en el destino.
* **La Solución (Destrucción Selectiva)**:
  En lugar de copiar partes de un documento a otro, abre una copia completa del documento original que ya tiene todas las imágenes y relaciones cargadas, y **elimina sistemáticamente** todo lo que no necesites:
  ```python
  # Para extraer solo una sección (Anexo A) y guardarla como otro archivo
  doc = docx.Document("tesis_completa.docx")
  body = doc.element.body
  
  # Remover elementos anteriores al inicio de la sección
  for idx in range(start_idx - 1, -1, -1):
      body.remove(list(body)[idx])
      
  # Remover elementos posteriores al final de la sección
  # (Se evalúa dinámicamente el final tras la primera poda)
  ```

---

## 3. Recetario de Funciones Útiles

### Formatear un Párrafo a Estilo Académico Estándar (Times New Roman, 1.5)
```python
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def format_academic_paragraph(p, text, is_bold=False, align_left=False):
    p.text = ""
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run.bold = is_bold
    
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    
    if align_left:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Inches(0.0)
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Inches(0.5)
```

### Inserción Segura de Líneas de Firma
```python
def add_signature_line(p_name, signature_line="__________________________________"):
    # Inserta una línea de firma inmediatamente encima del nombre indicado
    p_line = p_name.insert_paragraph_before()
    p_line.text = signature_line
    p_line.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_line.paragraph_format.space_before = Pt(12)
    p_line.paragraph_format.space_after = Pt(2)
    p_line.paragraph_format.first_line_indent = Inches(0.0)
    for run in p_line.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
```
