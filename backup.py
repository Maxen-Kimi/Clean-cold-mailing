import os
import shutil
from datetime import datetime
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.interval import IntervalTrigger
from config import get_config
from logger import log_info, log_error

class BackupManager:
    def __init__(self):
        self.config = get_config()
        self.scheduler = BackgroundScheduler()
        self.setup_backup_folder()
        
    def setup_backup_folder(self):
        """Crée le dossier de sauvegarde s'il n'existe pas."""
        os.makedirs(self.config.BACKUP_FOLDER, exist_ok=True)
        
    def create_backup(self):
        """Crée une sauvegarde des fichiers importants."""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_dir = os.path.join(self.config.BACKUP_FOLDER, f'backup_{timestamp}')
            os.makedirs(backup_dir)
            
            # Sauvegarde des fichiers Excel
            for file in os.listdir(self.config.UPLOAD_FOLDER):
                if file.endswith('.xlsx'):
                    src = os.path.join(self.config.UPLOAD_FOLDER, file)
                    dst = os.path.join(backup_dir, file)
                    shutil.copy2(src, dst)
            
            # Sauvegarde des logs
            if os.path.exists(self.config.LOG_FILE):
                shutil.copy2(
                    self.config.LOG_FILE,
                    os.path.join(backup_dir, 'app.log')
                )
            
            log_info(f"Sauvegarde créée avec succès: {backup_dir}")
            self.cleanup_old_backups()
            
        except Exception as e:
            log_error(e, "Erreur lors de la création de la sauvegarde")
    
    def cleanup_old_backups(self):
        """Supprime les anciennes sauvegardes."""
        try:
            backups = sorted([
                os.path.join(self.config.BACKUP_FOLDER, d)
                for d in os.listdir(self.config.BACKUP_FOLDER)
                if d.startswith('backup_')
            ])
            
            while len(backups) > self.config.MAX_BACKUPS:
                oldest_backup = backups.pop(0)
                shutil.rmtree(oldest_backup)
                log_info(f"Ancienne sauvegarde supprimée: {oldest_backup}")
                
        except Exception as e:
            log_error(e, "Erreur lors du nettoyage des anciennes sauvegardes")
    
    def start_scheduler(self):
        """Démarre le planificateur de sauvegardes."""
        self.scheduler.add_job(
            self.create_backup,
            trigger=IntervalTrigger(hours=self.config.BACKUP_INTERVAL),
            id='backup_job',
            replace_existing=True
        )
        self.scheduler.start()
        log_info("Planificateur de sauvegardes démarré")
    
    def stop_scheduler(self):
        """Arrête le planificateur de sauvegardes."""
        self.scheduler.shutdown()
        log_info("Planificateur de sauvegardes arrêté")

# Instance globale du gestionnaire de sauvegardes
backup_manager = BackupManager() 