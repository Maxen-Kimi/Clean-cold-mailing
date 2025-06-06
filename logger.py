import logging
import os
from logging.handlers import RotatingFileHandler
from config import get_config

def setup_logger():
    config = get_config()
    
    # Créer le dossier de logs s'il n'existe pas
    os.makedirs(config.LOGS_FOLDER, exist_ok=True)
    
    # Configuration du logger
    logger = logging.getLogger('clean_cold_mailing')
    logger.setLevel(getattr(logging, config.LOG_LEVEL))
    
    # Format du log
    formatter = logging.Formatter(config.LOG_FORMAT)
    
    # Handler pour le fichier
    file_handler = RotatingFileHandler(
        config.LOG_FILE,
        maxBytes=10*1024*1024,  # 10MB
        backupCount=5
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # Handler pour la console
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    return logger

# Création du logger global
logger = setup_logger()

def log_error(error, context=None):
    """Journalise une erreur avec son contexte."""
    error_msg = f"Erreur: {str(error)}"
    if context:
        error_msg += f" | Contexte: {context}"
    logger.error(error_msg)

def log_info(message, context=None):
    """Journalise un message d'information avec son contexte."""
    info_msg = message
    if context:
        info_msg += f" | Contexte: {context}"
    logger.info(info_msg)

def log_warning(message, context=None):
    """Journalise un avertissement avec son contexte."""
    warning_msg = message
    if context:
        warning_msg += f" | Contexte: {context}"
    logger.warning(warning_msg)

def log_debug(message, context=None):
    """Journalise un message de débogage avec son contexte."""
    debug_msg = message
    if context:
        debug_msg += f" | Contexte: {context}"
    logger.debug(debug_msg) 