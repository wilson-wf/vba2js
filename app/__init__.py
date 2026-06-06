from flask import Flask, send_from_directory
from flask_cors import CORS
from app.api.routes import api_bp
from config import config
import os

def create_app(config_name='default'):
    app = Flask(__name__)
    app.config.from_object(config[config_name])
    
    CORS(app)
    
    app.register_blueprint(api_bp, url_prefix='/api')
    
    @app.route('/')
    def serve_frontend():
        frontend_dist = os.path.join(app.config['BASE_DIR'], 'frontend', 'dist')
        return send_from_directory(frontend_dist, 'index.html')
    
    @app.route('/<path:path>')
    def serve_static(path):
        frontend_dist = os.path.join(app.config['BASE_DIR'], 'frontend', 'dist')
        file_path = os.path.join(frontend_dist, path)
        if os.path.exists(file_path):
            return send_from_directory(frontend_dist, path)
        return send_from_directory(frontend_dist, 'index.html')
    
    return app