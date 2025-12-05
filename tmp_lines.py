from pathlib import Path
text = Path('index.html').read_text().splitlines()
for i in range(220, 260):
    print(f"{i+1:04}: {text[i]}")
