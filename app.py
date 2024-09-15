from flask import Flask, request, render_template
import pandas as pd
import openpyxl

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "Вы не загрузили файлы"

    file = request.files['file']

    if file.filename == '':
        return "No selected file"

    if file:
        df = pd.read_excel(file)
        return render_template('data.html', tables=[df.to_html(classes='data')], titles=df.columns.values)

    return "Не удалось обработать файл"


if __name__ == '__main__':
    app.run(debug=True)
