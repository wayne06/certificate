import os
import pandas as pd
from docxtpl import DocxTemplate
from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
from zipfile import ZipFile

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        excel_file = request.files['excel']
        word_file = request.files['word']
        
        # 保存上传文件
        excel_path = os.path.join(UPLOAD_FOLDER, secure_filename(excel_file.filename))
        word_path = os.path.join(UPLOAD_FOLDER, secure_filename(word_file.filename))
        excel_file.save(excel_path)
        word_file.save(word_path)

        # 读取 Excel
        df = pd.read_excel(excel_path, engine='xlrd', skiprows=[0], header=0, dtype=str)
        df = df.dropna(subset=['姓名']).reset_index(drop=True)

        # 加载 Word 模板
        template = DocxTemplate(word_path)

        # 清空 outputs 目录
        for f in os.listdir(OUTPUT_FOLDER):
            os.remove(os.path.join(OUTPUT_FOLDER, f))

        # 渲染并保存
        output_files = []
        for _, row in df.iterrows():
            context = {
                '姓名': row['姓名'],
                '学号': row['学号'],
                '专业': row['专业'],
                '成绩': row['成绩'],
                '及格分': row['及格分'],
                '满分': row['满分'],
                '考核结论': row['考核结论'],
            }
            template.render(context)
            out_name = f"{row['姓名']}_成绩报告.docx"
            out_path = os.path.join(OUTPUT_FOLDER, out_name)
            template.save(out_path)
            output_files.append(out_path)

        # 打包 zip
        zip_path = os.path.join(OUTPUT_FOLDER, '成绩报告打包.zip')
        with ZipFile(zip_path, 'w') as zipf:
            for file in output_files:
                zipf.write(file, arcname=os.path.basename(file))

        return send_file(zip_path, as_attachment=True)

    return render_template('index.html')

if __name__ == '__main__':
	app.run(debug=True)
