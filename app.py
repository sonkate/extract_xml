import os
import pandas as pd
import xml.etree.ElementTree as ET
from flask import Flask, request, redirect, url_for, render_template, send_file
from werkzeug.utils import secure_filename
import zipfile

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'zip', 'xml'}

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        return redirect(url_for('process_file', filename=filename))
    return redirect(request.url)

@app.route('/process/<filename>')
def process_file(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    extract_folder = os.path.join(app.config['UPLOAD_FOLDER'], 'extracted')
    if not os.path.exists(extract_folder):
        os.makedirs(extract_folder)

    if filename.endswith('.zip'):
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(extract_folder)
    else:
        os.rename(file_path, os.path.join(extract_folder, filename))

    data = []
    for root, dirs, files in os.walk(extract_folder):
        for file in files:
            if file.endswith('.xml'):
                file_path = os.path.join(root, file)
                file_data = extract_all_info_from_xml(file_path)
                data.extend(file_data)

    df = pd.DataFrame(data)
    output_file = os.path.join(app.config['UPLOAD_FOLDER'], 'output.xlsx')
    df.to_excel(output_file, index=False)

    return send_file(output_file, as_attachment=True)

def extract_all_info_from_xml(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    data = []
    for hdon in root.findall('DLHDon'):
        record = {}
        ttchung = hdon.find('TTChung')
        record['Loại Hóa Đơn'] = ttchung.find('THDon').text if ttchung.find('THDon') is not None else None
        KHHDon = ttchung.find('KHHDon').text if ttchung.find('KHHDon') is not None else None
        SHDon = ttchung.find('SHDon').text if ttchung.find('SHDon') is not None else None
        record['Số hoá đơn/Tờ khai Hải quan điện tử'] = KHHDon + SHDon
        record['Ngày hóa đơn/ Tờ khai Hải quan điện tử'] = ttchung.find('NLap').text if ttchung.find('NLap') is not None else None
        record['Mã số thuế'] = ttchung.find('MSTTCGP').text if ttchung.find('MSTTCGP') is not None else None

        ttkhac = ttchung.find('TTKhac')
        if ttkhac is not None:
            for ttin in ttkhac.findall('TTin'):
                ttruong = ttin.find('TTruong').text
                dlieu = ttin.find('DLieu').text
                if ttruong == 'BankAccount':
                    record['Số tài khoản'] = dlieu
                if ttruong == 'BankName':
                    record['Tại Ngân hàng'] = dlieu

        ndh_don = hdon.find('NDHDon')
        nban = ndh_don.find('NBan')
        record['Người thụ hưởng'] = nban.find('Ten').text if nban.find('Ten') is not None else None
        record['Mã số thuế'] = nban.find('MST').text if nban.find('MST') is not None else None
        ttoan = ndh_don.find('TToan')
        record['Giá trị theo hoá đơn/Tờ khai Hải quan điện tử'] = ttoan.find('TgTTTBSo').text if ttoan.find('TgTTTBSo') is not None else None
        record['Số tiền giải ngân'] = ttoan.find('TgTTTBSo').text if ttoan.find('TgTTTBSo') is not None else None
        data.append(record)
    return data

if __name__ == '__main__':
    app.run(debug=True)
