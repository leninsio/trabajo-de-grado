import docx

doc = docx.Document("proyecto_de_grado/TRABAJO_LENIN_ALVARADO.docx")
print("--- CHAPTERS IN THE THESIS ---")
for idx, p in enumerate(doc.paragraphs):
    txt = p.text.strip().upper()
    if "CAPÍTULO" in txt or "CAPITULO" in txt:
        print(f"P{idx}: '{p.text}'")
