from pathlib import Path  # core library

import os, sys 
import pandas as pd  # pip install pandas
import win32com.client as win32  # pip install pywin32
import xlwings as xw  # pip install xlwings
from docxtpl import DocxTemplate  # pip install docxtpl


os.chdir(sys.path[0])
current_dir = Path(__file__).parent
template_path = current_dir / "modelo.docx"


doc = DocxTemplate('modelo.docx')

# Abrir a planilha Excel
wb = xw.Book('lista.xlsx')
ws = wb.sheets[0]  # Acessa a primeira aba da planilha
#sht_sales = wb.sheets["nome"]
# Ler os dados da planilha
# Supondo que os dados comecem na linha 2 e as colunas sejam Nome, Idade, Endereço
dados = ws.range('A2').expand('table').value
#df = sht_sales.range("A13").options(pd.DataFrame, index=False, expand="table").value

doc = DocxTemplate(str(template_path))
doc.render(dados)
doc.save("rendered_lista")