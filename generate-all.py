from docx.api import Document
import docx.document, docx.table
import os
from pathlib import Path
from xwlg import generate_word_list

def get_table(document: docx.document.Document) -> docx.table.Table:
    tables = document.tables
    if len(tables) == 1:
        return tables[0]
    if len(tables) > 1:
        selected = -1
        while selected < 1 or selected > len(tables):
            try:
                print("There is more than one table, please select which table you would like to use.")
                selected = int(input(f"Selected table (1-{len(tables)}): "))
                print()
                if selected < 1 or selected > len(tables):
                    raise ValueError
            except ValueError:
                print("Invalid input.")
        return tables[selected-1]
    print("Could not find any tables")
    exit()

# Get documents in cwd
files = [path for path in os.listdir() if path.endswith(".docx")]
if len(files) == 0:
    print("Could not find any .docx files")
    exit()
for path in files:
    print(f"Document found: {path}")
    document = Document(path)
    table = get_table(document)
    output_path = Path(Path(path).name.rsplit('.', 1)[0] + ".xlsx")

    _ = generate_word_list(table, output_path)
    print(f"Saved to {output_path}")