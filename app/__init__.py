from flask import Flask
from flask_cors import CORS
from app.api.routes import api_bp
from config import config

def create_app(config_name='default'):
    app = Flask(__name__)
    app.config.from_object(config[config_name])
    
    CORS(app)
    
    app.register_blueprint(api_bp, url_prefix='/api')
    
    return app