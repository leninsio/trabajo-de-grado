import re

manual_path = "guias_y_manuales/SEMINARIO_II_UNIDAD_I.txt"
with open(manual_path, "r", encoding="utf-8") as f:
    text = f.read()

print("=== SEARCHING CUAM MANUAL FOR RULES ===")

keywords = [
    "antecedentes", "teórica", "teorica", "legal", "ley", 
    "factible", "campo", "descriptiva", "explicativa",
    "manual", "norma"
]

lines = text.split("\n")
for idx, line in enumerate(lines):
    line_upper = line.upper()
    for kw in keywords:
        if kw.upper() in line_upper:
            # Print the line and some context
            start = max(0, idx - 2)
            end = min(len(lines), idx + 3)
            print(f"\nLine {idx} (matched keyword '{kw}'):")
            for i in range(start, end):
                print(f"  {i}: {lines[i]}")
            break
