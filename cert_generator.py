import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter , A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO
import os
import platform

# Caminhos
TEMPLATE_PATH = "templates/modelo.pdf"
EXCEL_PATH = "data/dados.xlsx"
OUTPUT_DIR = "output"

# Detectar sistema operacional e definir o caminho da fonte
system = platform.system()

print("Valor de system é " + system)
if system == 'Darwin':  # macOS
    FONT_PATH = "/System/Library/Fonts/Supplemental/Arial.ttf"
    FONT_NAME = 'Arial'
elif system == 'Windows':  # Windows
    FONT_PATH = r"C:\Windows\Fonts\arial.ttf"
    FONT_NAME = 'Arial'
else:  # Linux e outros aqui deve ser manjaro linux
    # Tente DejaVu Sans, disponível na maioria das distros Linux
    try:
        FONT_PATH = "/usr/share/fonts/TTF/DejaVuSans.ttf"
        FONT_NAME = 'DejaVuSans'
    except:
        FONT_PATH = "/usr/share/fonts/TTF/DejaVuSans-Bold.ttf"
        FONT_NAME = 'DejaVuSans-Bold'



# Posições dos textos no PDF (ajuste conforme seu template)
POS_NOME = (300, 400)
POS_CPF = (200, 370)
POS_CIDADE = (200, 340)

# Registrar fonte TrueType
try:
    pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
except Exception as e:
    print(f"Aviso: não foi possível carregar {FONT_PATH}, usando fonte interna padrão.\nErro: {e}")
    FONT_NAME = 'Helvetica'


def gerar_certificado(nome, cpf, cidade, output_filename):
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)

    # Configurar fonte
    can.setFont(FONT_NAME, 18)

    # Inserir dados
    can.drawString(*POS_NOME, f"Nome: {nome}")
    can.drawString(*POS_CPF, f"CPF: {cpf}")
    can.drawString(*POS_CIDADE, f"Cidade: {cidade}")
    can.save()

    # Combinar overlay com template
    packet.seek(0)
    overlay_pdf = PdfReader(packet)
    template_pdf = PdfReader(open(TEMPLATE_PATH, "rb"))
    output_pdf = PdfWriter()

    page = template_pdf.pages[0]
    page.merge_page(overlay_pdf.pages[0])
    output_pdf.add_page(page)

    # Criar pasta da cidade
    cidade_dir = os.path.join(OUTPUT_DIR, cidade.replace(" ", "_"))
    os.makedirs(cidade_dir, exist_ok=True)

    # Salvar PDF
    with open(os.path.join(cidade_dir, output_filename), "wb") as f_out:
        output_pdf.write(f_out)


def main():
    # Ler dados do Excel
    df = pd.read_excel(EXCEL_PATH)

    for _, row in df.iterrows():
        nome = str(row['Nome'])
        cpf = str(row['CPF'])
        cidade = str(row['Cidade'])
        filename = f"{nome.replace(' ', '_')}.pdf"
        gerar_certificado(nome, cpf, cidade, filename)


if __name__ == "__main__":
    main()

