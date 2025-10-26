#!/usr/bin/env python3
import os
import tempfile
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from extract_pptx_ocr import extract_text_from_pptx

app = Flask(__name__)
app.config['SECRET_KEY'] = os.urandom(24)
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
ALLOWED_EXTENSIONS = {'pptx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file selected')
        return redirect(request.url)
    
    file = request.files['file']
    
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))
    
    if file and allowed_file(file.filename):
        try:
            # Save uploaded file to temporary location
            filename = secure_filename(file.filename)
            upload_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(upload_path)
            
            # Process the PPTX file
            ocr_langs = request.form.get('ocr_langs', 'eng+ara')
            
            # Create output filename
            base_name = os.path.splitext(filename)[0]
            output_filename = f"{base_name}_extracted.txt"
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
            
            # Extract text
            extract_text_from_pptx(upload_path, output_path, ocr_langs)
            
            # Send the file as download
            return send_file(
                output_path,
                as_attachment=True,
                download_name=output_filename,
                mimetype='text/plain'
            )
            
        except Exception as e:
            flash(f'Error processing file: {str(e)}')
            return redirect(url_for('index'))
        finally:
            # Cleanup temporary files
            if os.path.exists(upload_path):
                os.remove(upload_path)
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except:
                    pass
    else:
        flash('Invalid file type. Please upload a .pptx file')
        return redirect(url_for('index'))

if __name__ == '__main__':
    import sys
    # Disable debug mode in production
    debug = os.environ.get('FLASK_DEBUG', 'False').lower() == 'true'
    host = os.environ.get('HOST', '0.0.0.0')
    port = int(os.environ.get('PORT', 5000))
    
    # For production, use Waitress instead of Flask dev server
    if not debug and '--dev' not in sys.argv:
        try:
            from waitress import serve
            print(f"Starting Waitress server on {host}:{port}")
            serve(app, host=host, port=port)
        except ImportError:
            print("Waitress not installed, using Flask dev server")
            app.run(debug=debug, host=host, port=port)
    else:
        app.run(debug=debug, host=host, port=port)

