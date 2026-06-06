import os
from datetime import timedelta

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'vba-converter-secret-key'
    
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))
    
    UPLOAD_FOLDER = os.path.join(BASE_DIR, 'data', 'uploads')
    CONVERTED_FOLDER = os.path.join(BASE_DIR, 'data', 'converted')
    LOG_FOLDER = os.path.join(BASE_DIR, 'data', 'logs')
    
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024
    
    ALLOWED_EXTENSIONS = {'xlsm'}
    
    LLM_API_KEY = os.environ.get('LLM_API_KEY') or None
    LLM_ENDPOINT = os.environ.get('LLM_ENDPOINT') or 'https://api.openai.com/v1/chat/completions'
    LLM_MODEL = os.environ.get('LLM_MODEL') or 'gpt-4'
    
    DATABASE_URI = 'sqlite:///' + os.path.join(BASE_DIR, 'data', 'converter.db')
    
    SESSION_PERMANENT = True
    PERMANENT_SESSION_LIFETIME = timedelta(hours=24)

class DevelopmentConfig(Config):
    DEBUG = True

class ProductionConfig(Config):
    DEBUG = False

config = {
    'development': DevelopmentConfig,
    'production': ProductionConfig,
    'default': DevelopmentConfig
}