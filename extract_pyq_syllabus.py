from pathlib import Path

import docx
from pypdf import PdfReader


PDFS = [
    r"C:\Users\Paras\AppData\Roaming\dragfile_file\tbc201_tbi201_endsem_2022_jul.pdf",
    r"C:\Users\Paras\AppData\Roaming\dragfile_file\tbc201_tbi201_tbs201_endsem_2023_jun.pdf",
    r"C:\Users\Paras\AppData\Roaming\dragfile_file\tbc201_endsem_2024_jun (2).pdf",
    r"C:\Users\Paras\AppData\Roaming\dragfile_file\tbc301_tbd301_tbi303_tbl305_endsem_2025_dec.pdf",
]

DOCX = r"C:\Users\Paras\AppData\Roaming\dragfile_file\TBD 301 Introdcution to Data structures-1.docx"
OUT = Path(r"C:\xampp\htdocs\lost-found\_extracted_pyq_syllabus.txt")


def main() -> None:
    chunks = []
    for file_name in PDFS:
        path = Path(file_name)
        chunks.append("\n" + "=" * 80 + f"\n{path.name}\n" + "=" * 80)
        reader = PdfReader(file_name)
        for page_number, page in enumerate(reader.pages, start=1):
            chunks.append(f"\n--- page {page_number} ---\n{page.extract_text() or ''}")

    doc_path = Path(DOCX)
    document = docx.Document(DOCX)
    chunks.append("\n" + "=" * 80 + f"\n{doc_path.name}\n" + "=" * 80)
    chunks.extend(paragraph.text for paragraph in document.paragraphs if paragraph.text.strip())
    for table_number, table in enumerate(document.tables, start=1):
        chunks.append(f"\n--- table {table_number} ---")
        for row in table.rows:
            cells = [" ".join(cell.text.split()) for cell in row.cells]
            chunks.append(" | ".join(cells))

    OUT.write_text("\n".join(chunks), encoding="utf-8")
    print(OUT)


if __name__ == "__main__":
    main()
