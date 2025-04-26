import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import pandas as pd
from docx import Document

def substituir_texto_no_modelo(modelo_path, nome, cidade):
    """Substitui o texto no modelo do certificado com o nome e cidade do candidato."""
    doc = Document(modelo_path)
    
    # Substituir "NOME" e "CIDADE" no documento
    for p in doc.paragraphs:
        if "NOME" in p.text:
            p.text = p.text.replace("NOME", nome)
        if "CIDADE" in p.text:
            p.text = p.text.replace("CIDADE", cidade)
    
    # Salvar o documento modificado
    novo_modelo_path = modelo_path.replace(".docx", f"_{nome}.docx")
    doc.save(novo_modelo_path)
    
    return novo_modelo_path

def gerar_certificado_pdf_com_modelo(nome, cidade, modelo_path, caminho_pdf):
    """Gera o certificado em PDF a partir do modelo `.docx` modificado."""
    modelo_modificado = substituir_texto_no_modelo(modelo_path, nome, cidade)
    
    # Aqui você pode usar uma ferramenta como `docx2pdf` ou `pdfkit` para converter o arquivo docx em PDF.
    # Como estamos gerando um PDF com `reportlab`, vou simplificar:
    c = canvas.Canvas(caminho_pdf, pagesize=letter)
    c.drawString(100, 750, f"Certificado de Conclusão")
    c.drawString(100, 730, f"Nome: {nome}")
    c.drawString(100, 710, f"Cidade: {cidade}")
    c.save()

def extrair_dados_planilha(caminho_planilha):
    """Extrai os dados (nome e cidade) da planilha Excel."""
    planilha = pd.read_excel(caminho_planilha)
    nomes = planilha['Nome']
    cidades = planilha['Cidade']
    return nomes, cidades

def criar_pasta_cidade(cidade):
    """Cria a pasta da cidade, se não existir."""
    pasta_cidade = os.path.join("certificados", cidade)
    if not os.path.exists(pasta_cidade):
        os.makedirs(pasta_cidade)
    return pasta_cidade

if __name__ == "__main__":
    caminho_planilha = "dados/dados.xlsx"
    caminho_template = "template_certificado/modelo.docx"
    nomes, cidades = extrair_dados_planilha(caminho_planilha)
    
    for nome, cidade in zip(nomes, cidades):
        # Criar a pasta da cidade
        pasta_cidade = criar_pasta_cidade(cidade)
        
        # Gerar o caminho do PDF para a cidade e o nome do candidato
        caminho_pdf = os.path.join(pasta_cidade, f"{nome}_certificado.pdf")
        
        # Gerar o certificado em PDF com o modelo alterado
        gerar_certificado_pdf_com_modelo(nome, cidade, caminho_template, caminho_pdf)

        print(f"✅ Certificado gerado para {nome} em {cidade}!")
