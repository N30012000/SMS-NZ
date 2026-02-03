from flask import Flask, render_template, request, send_file, jsonify
import os, tempfile, shutil, uuid
from werkzeug.utils import secure_filename
from ocr_processor import process_files_to_dataframe
from exporter import create_audit_excel
from dashboard import generate_dashboard_from_excel

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024 * 1024  # 1GB limit (adjust as needed)
ALLOWED_EXT = {'png','jpg','jpeg','tiff','pdf'}

def allowed(filename):
    return '.' in filename and filename.rsplit('.',1)[1].lower() in ALLOWED_EXT

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    files = request.files.getlist('files[]')
    if not files:
        return jsonify({'error':'No files uploaded'}), 400
    tmpdir = tempfile.mkdtemp(prefix='airsial_')
    saved = []
    for f in files:
        if f and allowed(f.filename):
            fname = secure_filename(f.filename)
            path = os.path.join(tmpdir, fname)
            f.save(path)
            saved.append(path)
    return jsonify({'tmpdir': tmpdir, 'count': len(saved)})

@app.route('/extract', methods=['POST'])
def extract():
    # Expect JSON: {"tmpdir": "..."}
    data = request.get_json()
    tmpdir = data.get('tmpdir')
    if not tmpdir or not os.path.isdir(tmpdir):
        return jsonify({'error':'Invalid tmpdir'}), 400
    try:
        df = process_files_to_dataframe(tmpdir)  # returns pandas DataFrame with one row per form
        out_path = os.path.join(tempfile.gettempdir(), f'airsial_extracted_{uuid.uuid4().hex}.xlsx')
        create_audit_excel(df, out_path)
        # remove tmpdir contents immediately
        shutil.rmtree(tmpdir, ignore_errors=True)
        return jsonify({'excel_path': out_path})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download', methods=['GET'])
def download():
    path = request.args.get('path')
    if not path or not os.path.isfile(path):
        return jsonify({'error':'File not found'}), 404
    # send file and then delete
    response = send_file(path, as_attachment=True)
    try:
        os.remove(path)
    except:
        pass
    return response

@app.route('/dashboard', methods=['POST'])
def dashboard():
    # Accept uploaded Excel file or path to previously generated file
    file = request.files.get('file')
    month = int(request.form.get('month'))
    year = int(request.form.get('year'))
    tmpfile = None
    if file:
        tmpfile = os.path.join(tempfile.gettempdir(), secure_filename(file.filename))
        file.save(tmpfile)
    else:
        return jsonify({'error':'No file uploaded'}), 400
    try:
        out = generate_dashboard_from_excel(tmpfile, month, year)
        # out is dict with 'html' preview path and 'excel' and 'pdf' paths
        return jsonify(out)
    finally:
        try:
            os.remove(tmpfile)
        except:
            pass

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
