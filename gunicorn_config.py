import multiprocessing
import os

# Nombre de workers
workers = multiprocessing.cpu_count() * 2 + 1

# Configuration des workers
worker_class = 'sync'
worker_connections = 1000
timeout = 30
keepalive = 2

# Configuration du logging
accesslog = 'logs/gunicorn-access.log'
errorlog = 'logs/gunicorn-error.log'
loglevel = 'info'

# Configuration de la sécurité
limit_request_line = 4094
limit_request_fields = 100
limit_request_field_size = 8190

# Configuration des performances
max_requests = 1000
max_requests_jitter = 50
graceful_timeout = 30

# Configuration du serveur
bind = f"0.0.0.0:{os.getenv('PORT', '5000')}"
preload_app = True 