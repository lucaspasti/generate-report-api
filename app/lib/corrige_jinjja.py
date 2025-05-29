import re

from docx import Document


def extrair_jinja_de_docx(caminho_docx):
    doc = Document(caminho_docx)

    # Junta todo o texto dos parágrafos em uma string
    texto_total = "\n".join(paragraph.text for paragraph in doc.paragraphs)

    # Encontra todos os marcadores Jinja do tipo {{ ... }}
    jinja_matches = re.findall(r"{{(.*?)}}", texto_total)

    # Verifica se os nomes são válidos (letras, números, underscores)
    marcadores_invalidos = []
    for match in jinja_matches:
        token = match.strip()
        if not re.match(r"^[A-Za-z0-9_]+$", token):
            marcadores_invalidos.append(f"{{{{ {token} }}}}")

    # Resultados
    print("✅ Total de marcadores encontrados:", len(jinja_matches))
    if marcadores_invalidos:
        print("❌ Marcadores Jinja com problemas de sintaxe:")
        for marcador in marcadores_invalidos:
            print(" -", marcador)
    else:
        print("🎉 Nenhum marcador inválido encontrado!")


# Exemplo de uso
extrair_jinja_de_docx(
    "C:/Users/ecpro_dhl3wmn/relatorio-ec-infra/relatorios/monitoramento_ruidos_subaquaticos.docx"
)
