from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import traceback
from formatter import format_document

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'

# Create folders if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return send_file('index.html')

@app.route('/style.css')
def style():
    return send_file('style.css')

@app.route('/script.js')
def script():
    return send_file('script.js')

@app.route('/format', methods=['POST'])
def format_doc():
    try:
        print("\n" + "="*50)
        print("📥 NEW UPLOAD REQUEST RECEIVED")
        print("="*50)
        
        if 'file' not in request.files:
            print("❌ Error: No file in request")
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        print(f"📄 File received: {file.filename}")
        
        if file.filename == '':
            print("❌ Error: Empty filename")
            return jsonify({'error': 'No file selected'}), 400
        
        if not file.filename.endswith('.docx'):
            print("❌ Error: Wrong file type")
            return jsonify({'error': 'Only .docx files are supported'}), 400
        
        # Save uploaded file
        input_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(input_path)
        print(f"✅ File saved to: {input_path}")
        
        # Format the document
        output_filename = f"formatted_{file.filename}"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        
        print(f"🔄 Starting formatting process...")
        format_document(input_path, output_path)
        print(f"✅ Formatting completed!")
        print(f"📦 Output saved to: {output_path}")
        
        # Send the formatted file back
        print(f"📤 Sending formatted file to user...")
        return send_file(output_path, as_attachment=True, download_name=output_filename)
    
    except Exception as e:
        print("\n" + "="*50)
        print("❌ ERROR OCCURRED:")
        print("="*50)
        print(f"Error message: {str(e)}")
        print("\nFull traceback:")
        print(traceback.format_exc())
        print("="*50 + "\n")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    print("\n🚀 Starting KG College Document Formatter...")
    print("📍 Server running at: http://localhost:5000")
    print("="*50 + "\n")
    app.run(debug=True, host='0.0.0.0', port=5000)