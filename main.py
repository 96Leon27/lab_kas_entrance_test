from flask import Flask, request, jsonify, send_file
import os
from werkzeug.utils import secure_filename
import openpyxl
import uuid

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'txt'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/public/report/export', methods=['POST'])
def export_report():
    temp_file = None
    excel_file = None

    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Файл не найден'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'Файл не выбран'}), 400

        if not allowed_file(file.filename):
            return jsonify({'error': 'Поддерживаются только .txt файлы'}), 400

        filename = secure_filename(file.filename)
        temp_file = os.path.join(UPLOAD_FOLDER, f'temp_{uuid.uuid4().hex}_{filename}')
        file.save(temp_file)

        with open(temp_file, 'r', encoding='utf-8') as f:
            text = f.read()

        words_to_find = ['житель', 'жителем']
        results = {word: {'total': text.lower().count(word),
                          'per_line': ','.join(list(map(lambda x: str(x.lower().count(word)), text.split('\n'))))}
                   for word in words_to_find}

        excel_file = os.path.join(UPLOAD_FOLDER, f'report_{uuid.uuid4().hex}.xlsx')

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Отчет"

        ws.append(['Словоформа', 'Кол-во во всем документе', 'Кол-во в каждой строке'])

        for word, data in results.items():
            ws.append([word, data['total'], data['per_line']])

        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width

        wb.save(excel_file)

        response = send_file(
            excel_file,
            as_attachment=True,
            download_name='report.xlsx'
        )

        @response.call_on_close
        def cleanup():
            try:
                if temp_file and os.path.exists(temp_file):
                    os.remove(temp_file)
                if excel_file and os.path.exists(excel_file):
                    os.remove(excel_file)
            except Exception as e:
                pass

        return response

    except Exception as e:
        if temp_file and os.path.exists(temp_file):
            try:
                os.remove(temp_file)
            except:
                pass
        if excel_file and os.path.exists(excel_file):
            try:
                os.remove(excel_file)
            except:
                pass
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
