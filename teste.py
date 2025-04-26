import os
import pypandoc
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

def substituir_texto_no_modelo(modelo_path, nome, cidade):
    """Substitui o texto no modelo do certificado com o nome e cidade do candidato."""
    doc = Document(modelo_path)
    
    # Definir o estilo da fonte
    font_name = 'Times New Roman'
    font_size = Pt(16)
    
    # Substituir "NOME" e "CIDADE" no documento e aplicar formatação
    for p in doc.paragraphs:
        if "NOME" in p.text:
            for run in p.runs:
                run.text = run.text.replace("NOME", nome.upper())  # Nome em maiúsculas
                run.font.name = font_name
                run.font.size = font_size
                # Para garantir que funcione em todos os idiomas
                r = run._r
                r.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        if "CIDADE" in p.text:
            for run in p.runs:
                run.text = run.text.replace("CIDADE", cidade)
                run.font.name = font_name
                run.font.size = font_size
                # Para garantir que funcione em todos os idiomas
                r = run._r
                r.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    
    # Salvar o documento modificado
    novo_modelo_path = modelo_path.replace(".docx", f"_{nome}.docx")
    doc.save(novo_modelo_path)
    
    return novo_modelo_path


def gerar_certificado_pdf_com_modelo(nome, cidade, modelo_path, caminho_pdf):
    """Gera o certificado em PDF a partir do modelo `.docx` modificado."""
    modelo_modificado = substituir_texto_no_modelo(modelo_path, nome, cidade)
    
    try:
        # Usando pypandoc para converter o arquivo .docx para .pdf
        output = pypandoc.convert_file(modelo_modificado, 'pdf', outputfile=caminho_pdf)
        
        # O caminho do PDF gerado será o especificado em `caminho_pdf`
        return caminho_pdf
    
    except Exception as e:
        print(f"Erro ao gerar o PDF: {e}")
        return None
    

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
        pdf_path = gerar_certificado_pdf_com_modelo(nome, cidade, caminho_template, caminho_pdf)

        if pdf_path:
            print(f"✅ Certificado gerado para {nome} em {cidade}!")
        else:
            print(f"❌ Erro ao gerar o certificado para {nome} em {cidade}.")
