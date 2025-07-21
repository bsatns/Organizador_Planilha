from flask import Flask, render_template, request, send_file
import pandas as pd
import os
import uuid
from werkzeug.utils import secure_filename
import zipfile
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
import numpy as np
from openpyxl.formatting.rule import FormulaRule


app = Flask(__name__)

# Aumentar o limite de upload (exemplo: 100 MB)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
RESULT_FOLDER = os.path.abspath(os.path.join(BASE_DIR, '..', 'results'))

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

def aplicar_formatacao_condicional(ws, coluna_obs):
    fill_cores = {
        "ATIVO WHATSAPP": "C6EFCE",
        "SEM INTERESSE": "FFC7CE",
        "VENDA": "00B050",
        "TEL NÃO ATENDEU": "FFEB9C",
        "SEM POSSIBILIDADES": "FF0000",
        "SEM CONTATO": "BDB76B",
        "SEM WHATSAPP": "8B008B",
        "SENDO TRAB": "FFA500"
    }

    inicio = 2
    fim = ws.max_row
    intervalo = f"{coluna_obs}{inicio}:{coluna_obs}{fim}"

    for status, cor in fill_cores.items():
        formula = f'${coluna_obs}{inicio}="{status}"'
        rule = FormulaRule(formula=[formula],
                           fill=PatternFill(start_color=cor, end_color=cor, fill_type="solid"))
        ws.conditional_formatting.add(intervalo, rule)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            num_parts = int(request.form['num_parts'])
            formatar = request.form.get('formatar') == 'sim'

            arquivos = request.files.getlist('Planilhas')
            if not arquivos:
                return "Nenhum arquivo anexado", 400

            zip_name = f"planilhas_{uuid.uuid4().hex}.zip"
            zip_path = os.path.join(RESULT_FOLDER, zip_name)

            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for file in arquivos:
                    filename = secure_filename(file.filename)
                    original_path = os.path.join(UPLOAD_FOLDER, filename)
                    file.save(original_path)

                    df = pd.read_excel(original_path)
                    splits = np.array_split(df, num_parts)

                    for i, part_df in enumerate(splits):
                        part_name = f"{filename.rsplit('.', 1)[0]}_part{i+1}.xlsx"
                        part_path = os.path.join(UPLOAD_FOLDER, part_name)
                        part_df.to_excel(part_path, index=False)

                        if formatar:
                            wb = openpyxl.load_workbook(part_path)
                            ws = wb.active

                            col_letter = get_column_letter(ws.max_column + 1)
                            ws[f'{col_letter}1'] = "OBSERVAÇÃO"

                            opcoes = [
                                "ATIVO WHATSAPP", "SEM INTERESSE", "VENDA", "TEL NÃO ATENDEU",
                                "SEM POSSIBILIDADES", "SEM CONTATO", "SEM WHATSAPP", "SENDO TRAB"
                            ]
                            dv = DataValidation(type="list", formula1='"' + ",".join(opcoes) + '"')
                            ws.add_data_validation(dv)
                            dv.add(f"{col_letter}2:{col_letter}{ws.max_row}")

                            # Garante células vazias, aplica validação e cor
                            for row in range(2, ws.max_row + 1):
                                ws[f'{col_letter}{row}'] = ""

                            aplicar_formatacao_condicional(ws, col_letter)

                            wb.save(part_path)

                        zipf.write(part_path, os.path.basename(part_path))
                        os.remove(part_path)

                    os.remove(original_path)

            return send_file(zip_path, as_attachment=True)

        except Exception as e:
            return f"Erro no processamento: {str(e)}", 500

    return render_template("index.html")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
