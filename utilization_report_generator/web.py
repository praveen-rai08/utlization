"""
Web UI for the Utilization Report Generator using Flask
"""

import os
import sys
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import tempfile
from datetime import datetime

from .core import ReportGenerator

app = Flask(__name__, template_folder='templates')
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'dev-secret-change-in-prod')
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()


@app.route('/')
def index():
    """Home page"""
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Handle file upload and generate reports"""
    
    try:
        # Check if file is in request
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not file.filename.endswith(('.xlsx', '.xls')):
            return jsonify({'error': 'Only Excel files (.xlsx, .xls) are supported'}), 400
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Generate reports
        try:
            generator = ReportGenerator(filepath)
            result = generator.generate()
            
            return jsonify({
                'success': True,
                'message': 'Reports generated successfully!',
                'excel_path': result['excel_path'],
                'html_path': result['html_path'],
                'summary': {
                    'total_associates': result['total_associates'],
                    'overall_avg_util': f"{result['overall_avg_util']}%",
                    'high_util_count': result['high_util_count'],
                    'medium_util_count': result['medium_util_count'],
                    'low_util_count': result['low_util_count'],
                }
            })
        
        except Exception as e:
            return jsonify({'error': f'Report generation failed: {str(e)}'}), 500
        
        finally:
            # Clean up uploaded file
            if os.path.exists(filepath):
                os.remove(filepath)
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/download/<file_type>')
def download_file(file_type):
    """Download or view a generated report.

    Excel -> served as attachment (triggers download).
    HTML  -> served inline so it opens in the browser tab.
    """
    try:
        file_path = request.args.get('path')

        if not file_path or not os.path.exists(file_path):
            return jsonify({'error': 'File not found'}), 404

        as_attachment = (file_type != 'html')
        return send_file(file_path, as_attachment=as_attachment)

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.errorhandler(404)
def not_found(error):
    """Handle 404 errors"""
    return jsonify({'error': 'Not found'}), 404


@app.errorhandler(500)
def server_error(error):
    """Handle 500 errors"""
    return jsonify({'error': 'Server error'}), 500


def run_web_app(host='127.0.0.1', port=5000, debug=False):
    """Run the Flask web application"""
    app.run(host=host, port=port, debug=debug)


if __name__ == '__main__':
    run_web_app(debug=True)
