from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for, flash
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from flask_talisman import Talisman
import pandas as pd
import os
import sys
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash   
from dotenv import load_dotenv
import mimetypes
from config import get_config
from logger import log_info, log_error, log_warning, log_debug
from backup import backup_manager

from clean_cold_mailing_script import (
    clean_name, generate_email, step1_extract_companies,
    step3_clean_and_complete, analyze_email_patterns,
    normalize_special_characters, capture_output
)

# Chargement des variables d'environnement
load_dotenv()

# Configuration de l'application
config = get_config()

app = Flask(__name__)
app.config.from_object(config)

# Configuration de la sécurité
Talisman(app,
    content_security_policy={
        'default-src': "'self'",
        'script-src': "'self' 'unsafe-inline' 'unsafe-eval' https:",
        'style-src': "'self' 'unsafe-inline' https:",
        'img-src': "'self' data: https:",
        'font-src': "'self' https:",
    },
    force_https=app.config.get('ENV') == 'production',
    strict_transport_security=True,
    session_cookie_secure=True,
    session_cookie_http_only=True
)

# Configuration de Flask-Login
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

# Classe utilisateur pour Flask-Login
class User(UserMixin):
    def __init__(self, id):
        self.id = id

@login_manager.user_loader
def load_user(user_id):
    return User(user_id)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def validate_excel_file(filepath):
    try:
        # Vérifier l'extension
        if not filepath.endswith('.xlsx'):
            return False
            
        # Essayer d'ouvrir le fichier avec pandas
        df = pd.read_excel(filepath)
        
        # Vérifier que le fichier contient des données
        if df.empty:
            return False
            
        return True
    except Exception as e:
        log_error(e, f"Erreur de validation du fichier Excel: {filepath}")
        return False

# Créer les dossiers nécessaires
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        if username == os.getenv('ADMIN_USERNAME') and \
           check_password_hash(os.getenv('ADMIN_PASSWORD'), password):
            user = User(username)
            login_user(user)
            log_info(f"Connexion réussie pour l'utilisateur: {username}")
            next_page = request.args.get('next')
            return redirect(next_page or url_for('index'))
        
        log_warning(f"Tentative de connexion échouée pour l'utilisateur: {username}")
        flash('Nom d\'utilisateur ou mot de passe incorrect')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    log_info(f"Déconnexion de l'utilisateur: {current_user.id}")
    logout_user()
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
@login_required
def upload_file():
    if 'file' not in request.files:
        log_warning("Tentative d'upload sans fichier")
        return jsonify({'error': 'Aucun fichier n\'a été envoyé'}), 400
    
    file = request.files['file']
    if file.filename == '':
        log_warning("Tentative d'upload avec un nom de fichier vide")
        return jsonify({'error': 'Aucun fichier sélectionné'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        try:
            file.save(filepath)
            if not validate_excel_file(filepath):
                os.remove(filepath)
                log_error(f"Fichier Excel invalide: {filename}")
                return jsonify({'error': 'Le fichier n\'est pas un fichier Excel valide'}), 400
            
            log_info(f"Fichier uploadé avec succès: {filename}")
            return jsonify({'message': 'Fichier téléchargé avec succès', 'filename': filename})
        except Exception as e:
            if os.path.exists(filepath):
                os.remove(filepath)
            log_error(e, f"Erreur lors de l'upload du fichier: {filename}")
            return jsonify({'error': f'Erreur lors du téléchargement: {str(e)}'}), 500
    
    log_warning(f"Tentative d'upload d'un fichier non autorisé: {file.filename}")
    return jsonify({'error': 'Format de fichier non supporté'}), 400

@app.route('/process', methods=['POST'])
@login_required
def process_file():
    data = request.json
    action = data.get('action')
    filename = data.get('filename')
    
    if not filename:
        log_warning("Tentative de traitement sans nom de fichier")
        return jsonify({'error': 'Nom de fichier manquant'}), 400
    
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    if not os.path.exists(filepath):
        log_error(f"Fichier non trouvé: {filename}")
        return jsonify({'error': 'Fichier non trouvé'}), 404
    
    try:
        log_info(f"Début du traitement {action} sur le fichier: {filename}")
        
        if action == 'extract_companies':
            output, success = capture_output(step1_extract_companies)
            message = 'Entreprises extraites avec succès' if success else 'Erreur lors de l\'extraction'
        
        elif action == 'analyze_patterns':
            output, success = capture_output(analyze_email_patterns, filepath)
            message = 'Patterns analysés avec succès' if success else 'Erreur lors de l\'analyse'
        
        elif action == 'clean_complete':
            output, success = capture_output(step3_clean_and_complete, filepath)
            message = 'Nettoyage et complétion terminés' if success else 'Erreur lors du nettoyage'
        
        elif action == 'normalize':
            output, success = capture_output(normalize_special_characters, filepath)
            message = 'Normalisation terminée' if success else 'Erreur lors de la normalisation'
        
        else:
            log_warning(f"Action non reconnue: {action}")
            return jsonify({'error': 'Action non reconnue'}), 400
        
        if success:
            log_info(f"Traitement {action} réussi pour le fichier: {filename}")
        else:
            log_error(f"Échec du traitement {action} pour le fichier: {filename}")
        
        return jsonify({
            'message': message,
            'output': output
        })
    
    except Exception as e:
        log_error(e, f"Erreur lors du traitement {action} du fichier: {filename}")
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
@login_required
def download_file(filename):
    if not allowed_file(filename):
        log_warning(f"Tentative de téléchargement d'un fichier non autorisé: {filename}")
        return jsonify({'error': 'Format de fichier non autorisé'}), 400
        
    # Vérifier d'abord dans le dossier uploads
    upload_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(upload_path):
        log_info(f"Téléchargement du fichier: {filename}")
        return send_file(
            upload_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    # Si non trouvé, vérifier dans le répertoire courant
    current_path = os.path.join(os.getcwd(), filename)
    if os.path.exists(current_path):
        log_info(f"Téléchargement du fichier: {filename}")
        return send_file(
            current_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    log_error(f"Fichier non trouvé lors du téléchargement: {filename}")
    return jsonify({'error': f'Le fichier {filename} n\'a pas été trouvé'}), 404

@app.route('/download-cleaned-contacts')
@login_required
def download_cleaned_contacts():
    filename = 'cleaned_contacts.xlsx'
    file_path = os.path.join(os.getcwd(), filename)
    
    if os.path.exists(file_path):
        log_info("Téléchargement du fichier cleaned_contacts.xlsx")
        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    log_error("Fichier cleaned_contacts.xlsx non trouvé")
    return jsonify({'error': 'Le fichier cleaned_contacts.xlsx n\'a pas été trouvé'}), 404

@app.errorhandler(404)
def not_found_error(error):
    log_error(f"Page non trouvée: {request.url}")
    return render_template('404.html'), 404

@app.errorhandler(500)
def internal_error(error):
    log_error(f"Erreur serveur: {str(error)}")
    return render_template('500.html'), 500

def init_app():
    """Initialise l'application."""
    # Démarrage du planificateur de sauvegardes
    backup_manager.start_scheduler()
    log_info("Application initialisée")

if __name__ == '__main__':
    init_app()
    app.run(
        host=config.HOST,
        port=config.PORT,
        debug=config.DEBUG
    ) 