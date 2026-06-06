import os
from flask import Blueprint, request, jsonify, send_from_directory, current_app
from app.services.file_service import FileService
from app.services.convert_service import ConvertService
from app.services.test_service import TestService
from app.utils.database import DatabaseManager

api_bp = Blueprint('api', __name__)

@api_bp.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    file_service = FileService(current_app.config)
    file_id, error = file_service.upload_file(file)
    
    if error:
        return jsonify({'error': error}), 400
    
    return jsonify({'file_id': file_id, 'status': 'pending'}), 200

@api_bp.route('/convert/<file_id>', methods=['POST'])
def convert_file(file_id):
    file_service = FileService(current_app.config)
    file_info = file_service.get_file_info(file_id)
    
    if not file_info:
        return jsonify({'error': 'File not found'}), 404
    
    data = request.get_json() or {}
    use_llm = data.get('use_llm', False)
    
    source_path = file_service.get_upload_path(file_id)
    
    def progress_callback(progress, stage, message):
        file_service.update_file_status(file_id, 'processing', progress, stage)
    
    convert_service = ConvertService(current_app.config)
    output_path, error = convert_service.convert(file_id, source_path, use_llm, progress_callback)
    
    if error:
        file_service.update_file_status(file_id, 'failed', 0, 'failed', error)
        return jsonify({'error': error}), 500
    
    file_service.update_converted_path(file_id, output_path)
    
    return jsonify({
        'file_id': file_id,
        'status': 'completed',
        'download_url': f'/api/download/{file_id}'
    }), 200

@api_bp.route('/convert/status/<file_id>', methods=['GET'])
def get_convert_status(file_id):
    file_service = FileService(current_app.config)
    file_info = file_service.get_file_info(file_id)
    
    if not file_info:
        return jsonify({'error': 'File not found'}), 404
    
    return jsonify({
        'file_id': file_info['id'],
        'original_name': file_info['original_name'],
        'status': file_info['status'],
        'progress': file_info['progress'],
        'stage': file_info['stage'],
        'error_message': file_info['error_message']
    }), 200

@api_bp.route('/download/<file_id>', methods=['GET'])
def download_file(file_id):
    file_service = FileService(current_app.config)
    file_info = file_service.get_file_info(file_id)
    
    if not file_info or file_info['status'] != 'completed':
        return jsonify({'error': 'File not available'}), 404
    
    converted_path = file_info['converted_path']
    if not converted_path or not os.path.exists(converted_path):
        return jsonify({'error': 'File not found'}), 404
    
    filename = f"{os.path.splitext(file_info['original_name'])[0]}_wps_js.xlsx"
    
    return send_from_directory(
        os.path.dirname(converted_path),
        os.path.basename(converted_path),
        as_attachment=True,
        download_name=filename
    )

@api_bp.route('/config/llm', methods=['POST'])
def configure_llm():
    data = request.get_json()
    
    if not data or 'endpoint' not in data or 'model' not in data:
        return jsonify({'error': 'Missing required parameters'}), 400
    
    db_path = os.path.join(current_app.config['BASE_DIR'], 'data', 'converter.db')
    db = DatabaseManager(db_path)
    
    db.save_llm_config(
        api_key=data.get('api_key'),
        endpoint=data['endpoint'],
        model=data['model']
    )
    
    return jsonify({'success': True, 'message': 'LLM configuration updated'}), 200

@api_bp.route('/config/llm', methods=['GET'])
def get_llm_config():
    db_path = os.path.join(current_app.config['BASE_DIR'], 'data', 'converter.db')
    db = DatabaseManager(db_path)
    
    config = db.get_active_llm_config()
    
    if config:
        return jsonify({
            'endpoint': config['endpoint'],
            'model': config['model'],
            'has_api_key': config['api_key'] is not None
        }), 200
    
    return jsonify({
        'endpoint': current_app.config.get('LLM_ENDPOINT', ''),
        'model': current_app.config.get('LLM_MODEL', ''),
        'has_api_key': current_app.config.get('LLM_API_KEY') is not None
    }), 200

@api_bp.route('/test/js', methods=['POST'])
def test_js_code():
    data = request.get_json()
    
    if not data or 'code' not in data:
        return jsonify({'error': 'Missing code parameter'}), 400
    
    test_service = TestService(current_app.config)
    result = test_service.run_full_test(data['code'])
    
    return jsonify(result), 200

@api_bp.route('/health', methods=['GET'])
def health_check():
    return jsonify({'status': 'ok', 'service': 'vba-converter'}), 200