# Application de Nettoyage de Contacts

Cette application Flask permet de nettoyer et normaliser des fichiers Excel contenant des contacts pour le cold mailing.

## Fonctionnalités

- Interface web sécurisée avec authentification
- Upload et validation de fichiers Excel
- Nettoyage et normalisation des données
- Génération d'emails
- Analyse de patterns
- Sauvegarde automatique des données
- Journalisation des opérations

## Prérequis

- Python 3.9+
- pip (gestionnaire de paquets Python)

## Installation

1. Cloner le dépôt :
```bash
git clone [URL_DU_REPO]
cd [NOM_DU_DOSSIER]
```

2. Créer un environnement virtuel :
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows
```

3. Installer les dépendances :
```bash
pip install -r requirements.txt
```

4. Configurer les variables d'environnement :
- Copier le fichier `templates/doc.env` vers `.env`
- Modifier les valeurs dans `.env` selon vos besoins

## Structure du Projet

```
.
├── app.py                 # Application principale
├── config.py             # Configuration
├── logger.py             # Gestion des logs
├── backup.py             # Gestion des sauvegardes
├── clean_cold_mailing_script.py  # Script de traitement
├── tests.py              # Tests unitaires
├── requirements.txt      # Dépendances
├── Procfile             # Configuration pour Render
├── gunicorn_config.py   # Configuration Gunicorn
├── runtime.txt          # Version Python
├── uploads/             # Dossier des fichiers uploadés
├── logs/                # Dossier des logs
├── backups/             # Dossier des sauvegardes
└── templates/           # Templates HTML
```

## Déploiement sur Render

1. Créer un nouveau service Web sur Render
2. Connecter votre dépôt Git
3. Configurer les variables d'environnement sur Render
4. Déployer l'application

## Tests

Exécuter les tests unitaires :
```bash
python tests.py
```

## Sécurité

- Authentification requise pour toutes les routes
- Protection CSRF
- Validation des fichiers
- Limitation de la taille des fichiers
- Sauvegarde automatique
- Journalisation des opérations

## Maintenance

- Les logs sont automatiquement archivés
- Les sauvegardes sont effectuées quotidiennement
- Les anciennes sauvegardes sont automatiquement nettoyées

## Licence

[VOTRE_LICENCE]

## Contact

[VOS_INFORMATIONS_DE_CONTACT] 