import os
import io
import json
from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook
import openai
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

openai.api_key = os.getenv('OPENAI_API_KEY')

FIVE_DOMAINS = [
    "健康・生活",
    "運動・感覚",
    "認知・行動",
    "言語・コミュニケーション",
    "人間関係・社会性"
]

def generate_text_with_gpt(prompt, max_tokens=500):
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "あなたは放課後等デイサービスの専門支援計画書作成の専門家です。"},
                {"role": "user", "content": prompt}
            ],
            max_tokens=max_tokens,
            temperature=0.7
        )
        return response['choices'][0]['message']['content']
    except Exception as e:
        print(f"GPT API Error: {str(e)}")
        return f"エラー: {str(e)}"

def process_excel(file_content):
    excel_file = io.BytesIO(file_content)
    wb = load_workbook(excel_file)
    ws = wb.active
    
    v33_value = ws['V33'].value or ""
    ws['B12'].value = v33_value
    print(f"✓ Step 1: V33 → B12 転記完了")
    
    b12_content = ws['B12'].value or ""
    if b12_content:
        prompt_p12 = f"""以下のアセスメント概要を基に、放課後等デイサービスの5領域（{', '.join(FIVE_DOMAINS)}）を踏まえて、特に専門的な支援を要する領域とその具体的な内容を500字以内で生成してください。

アセスメント概要:
{b12_content}

回答は�条書き形式で、各領域ごとに支援が必要な内容を明記してください。"""
        p12_content = generate_text_with_gpt(prompt_p12)
        ws['P12'].value = p12_content
        print(f"✓ Step 2: P12（5領域分析）自動生成完了")
    
    p12_content = ws['P12'].value or ""
    if p12_content:
        prompt_l19 = f"""以下の専門的支援が必要な領域の情報から、1年間の長期目標を300字以内で生成してください。目標は具体的で測定可能な形で記述してください。

支援が必要な領域:
{p12_content}

長期目標:"""
        l19_content = generate_text_with_gpt(prompt_l19, max_tokens=300)
        ws['L19'].value = l19_content
        print(f"✓ Step 3: L19（長期目標）自動生成完了")
    
    if p12_content:
        prompt_l21 = f"""以下の専門的支援が必要な領域の情報から、3ヶ月の短期目標を300字以内で生成してください。目標は具体的で測定可能な形で記述してください。

支援が必要な領域:
{p12_content}

短期目標:"""
        l21_content = generate_text_with_gpt(prompt_l21, max_tokens=300)
        ws['L21'].value = l21_content
        print(f"✓ Step 4: L21（短期目標）自動生成完了")
    
    l19_content = ws['L19'].value or ""
    l21_content = ws['L21'].value or ""
    
    support_contents = ['C25', 'C27', 'C29']
    for i, cell in enumerate(support_contents, 1):
        if l19_content and l21_content:
            prompt_support = f"""以下の長期目標と短期目標に基づいて、第{i}項目の具体的な支援の内容及び支援の実施方法を400字以内で生成してください。実施方法は週単位での活動内容を含めてください。

長期目標: {l19_content}
短期目標: {l21_content}

第{i}項目の具体的な支援内容と実施方法:"""
            support_content = generate_text_with_gpt(prompt_support, max_tokens=400)
            ws[cell].value = support_content
            print(f"✓ Step 5-{i}: {cell}（支援内容）自動生成完了")
    
    caution_cells = ['U25', 'U27', 'U29']
    for i, (support_cell, caution_cell) in enumerate(zip(support_contents, caution_cells), 1):
        support_content = ws[support_cell].value or ""
        if support_content:
            prompt_caution = f"""以下の支援内容に対する留意点を250字以内で生成してください。安全性、効果性、個別対応の観点から記述してください。

支援内容:
{support_content}

留意点:"""
            caution_content = generate_text_with_gpt(prompt_caution, max_tokens=250)
            ws[caution_cell].value = caution_content
            print(f"✓ Step 6-{i}: {caution_cell}（留意点）自動生成完了")
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_file():
    try:
        if 'file' not in request.files:
            return {'error': 'ファイルが選択されていません'}, 400
        
        file = request.files['file']
        if file.filename == '':
            return {'error': 'ファイルが選択されていません'}, 400
        
        if not file.filename.endswith('.xlsx'):
            return {'error': 'Excelファイル（.xlsx）のみ対応しています'}, 400
        
        file_content = file.read()
        processed_file = process_excel(file_content)
        
        return send_file(
            processed_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='専門的支援実施計画書-処理済み.xlsx'
        )
    
    except Exception as e:
        print(f"Error: {str(e)}")
        return {'error': f'処理中にエラーが発生しました: {str(e)}'}, 500

@app.route('/health')
def health():
    return {'status': 'ok'}

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)), debug=False)