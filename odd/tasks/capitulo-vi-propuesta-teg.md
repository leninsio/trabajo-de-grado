# Tareas: Desarrollo del Capítulo VI (La Propuesta) y Completitud Metodológica TEG

## Objetivo
Desarrollar y consolidar el Capítulo VI ("La Propuesta") conforme al modelo oficial del CUAM para Proyectos Factibles en Informática, integrando la arquitectura técnica real de los subsistemas desarrollados (`app-empenos` y `ENTRADAS-Y-SALIDAS-TIENDA`), y corrigiendo las deficiencias de páginas preliminares y Fases de Proyecto Factible en el documento de trabajo de grado.

## Alcance y Restricciones
- Archivo objetivo principal: `/Users/leninsio/Downloads/trabajo de grados/proyecto_de_grado/trabajo_de_grado.docx`
- Archivo de respaldo/sincronización: `proyecto_de_grado/final.docx`
- Cumplimiento estricto de normas UPEL/CUAM: Times New Roman 12 pt, interlineado 1.5, sangría de primera línea de 1.27 cm (0.5"), márgenes izquierdo 4 cm, superior/derecho/inferior 3 cm.
- Respetar intactas las imágenes de los anexos (validación de expertos en Anexo B).
- No inventar stacks o componentes ficticios: documentar la arquitectura real descubierta (Next.js/React + Prisma/PostgreSQL/Supabase para empeños; Django REST Framework/Python + Vite/React para caja/inventario).

## Lista de Tareas

- [x] **TASK-01**: Consolidar especificación técnica de la solución (Arquitectura, Modelo de Datos Relacional, Roles de Usuario, Reglas Financieras y Fórmulas de Empeño/Compra).
  - *Ruta*: Inline
  - *Verificación*: Contenido técnico extraído y cotejado contra los repositorios reales `app-empenos` y `ENTRADAS-Y-SALIDAS-TIENDA`.

- [x] **TASK-02**: Redactar el texto completo del Capítulo VI según la estructura oficial del CUAM (Presentación, Justificación, Objetivos de la Propuesta, Fundamentación, Estudio de Factibilidad Técnica/Operativa/Económica con cuadros 9, 10, 11 y 12, Estructura del Sistema, y Diseño con especificación funcional y manual de usuario operativo).
  - *Ruta*: Inline
  - *Verificación*: Coherencia formal y completitud según `modelo_proyecto_factible.pdf` (pp. 95–128).

- [x] **TASK-03**: Redactar la sección "Fases del Proyecto Factible" para el Capítulo III y las tablas de "Lista de Cuadros" y "Lista de Gráficos" para las páginas preliminares.
  - *Ruta*: Inline
  - *Verificación*: Coincidencia exacta de títulos y números de página con los cuadros 1–12 y gráficos 1–8.

- [x] **TASK-04**: Implementar el script de inyección en `trabajo_de_grado.docx` utilizando `python-docx` bajo las reglas de `docx-manipulation` (preservando estilos, márgenes, numeración y relaciones de imágenes de Anexo B).
  - *Ruta*: Delegated direct / script python
  - *Verificación*: Generación sin corrupción del archivo `.docx` tanto en Downloads como en workspace.

- [x] **TASK-05**: Auditoría y verificación final del documento modificado (inspección de párrafos, tablas, imágenes conservadas y cálculo de páginas).
  - *Ruta*: Inline
  - *Verificación*: Inspección con python-docx confirmando 659 párrafos, 15 tablas y 12 imágenes del Anexo B intactas.
