import os
from dotenv import load_dotenv

# Chargement des variables d'environnement
load_dotenv()

class Config:
    # Configuration de base
    SECRET_KEY = os.getenv('FLASK_SECRET_KEY')
    DEBUG = False
    
    # Configuration des dossiers
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
    LOGS_FOLDER = os.path.join(BASE_DIR, 'logs')
    BACKUP_FOLDER = os.path.join(BASE_DIR, 'backups')
    
    # Configuration des fichiers
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max
    ALLOWED_EXTENSIONS = {'xlsx'}
    
    # Configuration de la journalisation
    LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    LOG_LEVEL = 'INFO'
    LOG_FILE = os.path.join(LOGS_FOLDER, 'app.log')
    
    # Configuration de la sauvegarde
    BACKUP_INTERVAL = 24  # heures
    MAX_BACKUPS = 7  # nombre de sauvegardes à conserver
    
    # Configuration de la sécurité
    SESSION_COOKIE_SECURE = True
    SESSION_COOKIE_HTTPONLY = True
    REMEMBER_COOKIE_SECURE = True
    REMEMBER_COOKIE_HTTPONLY = True
    
    # Configuration du serveur
    HOST = '0.0.0.0'
    PORT = int(os.getenv('PORT', 5000))

class DevelopmentConfig(Config):
    DEBUG = True
    LOG_LEVEL = 'DEBUG'
    ENV = "development"

class ProductionConfig(Config):
    DEBUG = False
    LOG_LEVEL = 'INFO'
    ENV = "production"
# Configuration par défaut
config = {
    'development': DevelopmentConfig,
    'production': ProductionConfig,
    'default': DevelopmentConfig
}

def get_config():
    env = os.getenv('FLASK_ENV', 'default')
    return config[env] 