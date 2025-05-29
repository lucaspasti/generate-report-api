import os

import docx

# Caminho da pasta onde estão os arquivos .docx
folder_path = "3. ANTAQ"

# Lista todos os arquivos .docx na pasta
docx_files = [
    os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(".docx")
]

# Itera sobre cada arquivo .docx
for file_path in docx_files:
    # Abre o documento
    doc = docx.Document(file_path)

    # Substitui colchetes por chaves em todos os parágrafos
    for para in doc.paragraphs:
        para.text = para.text.replace("[", "{").replace("]", "}")

    # Cria um nome novo para o arquivo tratado
    file_name = os.path.basename(file_path)
    treated_file_name = f"tratado_{file_name}"
    new_file_path = os.path.join(folder_path, treated_file_name)

    # Salva o novo arquivo
    doc.save(new_file_path)

print("Tratamento concluído.")
