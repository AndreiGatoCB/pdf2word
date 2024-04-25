import fitz
import os
import re
import shutil
from docxtpl import DocxTemplate
from docx import Document

pdf_path = 'pdf/Alicia_en_la_ciudad_de_la_furia.pdf'
plantilla_path = 'word/plantilla.docx'

extracted_text_ = ''
with fitz.open(pdf_path) as pdf_document:
    for page_number in range(pdf_document.page_count):
        page = pdf_document.load_page(page_number)
        page_text = page.get_text()
        extracted_text_ += page_text

extracted_text_ = extracted_text_.replace('\n', '\n\n')
extracted_text_ = extracted_text_.replace('-\n\n', '- ')
extracted_text_ = extracted_text_.replace('+\n\n', '  + ')
extracted_text_ = extracted_text_.replace(': ', ':\n')
extracted_text_ = extracted_text_.replace('\n(Desaparece por pata izquierda!', '\n(Desaparece por pata izquierda!)')
extracted_text_ = extracted_text_.replace(r'\n\n([a-z]+)', '\n\1')
extracted_text_ = re.sub(r'\n\n([a-z]+)', r'\n\1', extracted_text_)
extracted_text_ = re.sub('\n([a-zA-Z0-9áéíóúÁÉÍÓÚüÜ ]+):', r'\n<b>\1:</b>', extracted_text_)
extracted_text_ = re.sub(r'\(([^)]+)[!.]', r'\n<i>(\1)</i>', extracted_text_)
extracted_text_ = re.sub(r'\((.*?)\)', r'<i>(\1</i>', extracted_text_)
context = {'text': extracted_text_}
# print(context)

output_dir, output_filename = os.path.split(pdf_path)
output_name, output_ext = os.path.splitext(output_filename)
output_filename = f'{output_name}.md'
word_output_filename = f'word/{output_name}.docx'
output_path = os.path.join(output_dir, output_filename)
shutil.copy(plantilla_path, word_output_filename)
doc_ = DocxTemplate(word_output_filename)
doc_.save(word_output_filename)

doc = Document(word_output_filename)
doc.add_paragraph(extracted_text_)
doc.save(word_output_filename)

archivo_md = output_path

# Abrir el archivo en modo escritura
with open(archivo_md, "w", encoding="utf-8") as archivo:
    # Escribir el contenido en el archivo
    archivo.write(extracted_text_)
# print(extracted_text_)
print(f"El archivo {archivo_md} se ha creado correctamente.")
