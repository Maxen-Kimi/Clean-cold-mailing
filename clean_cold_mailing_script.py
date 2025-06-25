import pandas as pd
import re
import os
from collections import Counter
import unidecode
import openpyxl
import io
import sys
from contextlib import redirect_stdout

print(f"Pandas version: {pd.__version__}")
print(f"Openpyxl version: {openpyxl.__version__}")

def clean_name(name):
    if pd.isna(name):
        return ""
    # Convertir en string si ce n'est pas déjà le cas
    name = str(name)
    # Supprimer les caractères spéciaux
    name = re.sub(r'[^a-zA-ZÀ-ÿ\-\s]', '', name)
    # Ne garder que la première partie si trait d'union
    name = name.split('-')[0]
    # Nettoyer les espaces
    name = name.strip()
    # Vérifier la longueur minimale
    if len(name) <= 1:
        return ""
    return name.lower()

def generate_email(row, pattern):
    if pd.isna(pattern):
        return ""
    
    firstname = clean_name(row['Prénom'])
    lastname = clean_name(row['Nom'])
    company = row['Société'].lower()
    
    if not firstname or not lastname:
        return ""
        
    # Extraire la première lettre pour les initiales
    firstname_initial = firstname[0] if firstname else ""
    lastname_initial = lastname[0] if lastname else ""
    
    email = pattern.lower()
    # Remplacer d'abord les patterns d'initiales
    email = email.replace('firstnameinitial', firstname_initial)
    email = email.replace('lastnameinitial', lastname_initial)
    # Puis remplacer les noms complets
    email = email.replace('firstname', firstname)
    email = email.replace('lastname', lastname)
    email = email.replace('company', company)
    
    return email

def step1_extract_companies():
    print("🔹 Étape 1: Extraction des entreprises uniques")
    
    # Charger le fichier d'entrée
    df = pd.read_excel('input.xlsx')
    
    # Filtrer les lignes avec email qualification 'catch_all@pro' ou vide
    mask = (df['Email Qualification'].str.contains('catch_all@pro', na=False)) | (df['Email Qualification'].isna())
    df_filtered = df[mask]
    
    # Extraire les entreprises uniques
    companies = df_filtered['Société'].unique()
    companies = sorted([c for c in companies if pd.notna(c)])
    
    # Créer et sauvegarder le fichier des entreprises
    companies_df = pd.DataFrame({'Société': companies})
    companies_df.to_excel('companies.xlsx', index=False)
    
    print(f"✅ Fichier companies.xlsx créé avec succès ({len(companies)} entreprises)")

def capture_output(func, *args, **kwargs):
    output = io.StringIO()
    with redirect_stdout(output):
        try:
            func(*args, **kwargs)
            success = True
        except Exception as e:
            print(f"❌ Erreur: {str(e)}")
            success = False
    return output.getvalue(), success

def step3_clean_and_complete(filename='input.xlsx'):
    print("🔹 Étape 3: Nettoyage et complétion des emails")
    
    # Vérifier que les fichiers nécessaires existent
    if not os.path.exists('detected_patterns.xlsx'):
        print("❌ Erreur: detected_patterns.xlsx non trouvé")
        return False
        
    if not os.path.exists(filename):
        print(f"❌ Erreur: {filename} non trouvé")
        return False
    
    try:
        # Charger les fichiers
        patterns_df = pd.read_excel('detected_patterns.xlsx')
        input_df = pd.read_excel(filename)
        
        # Créer un dictionnaire des patterns
        patterns_dict = dict(zip(patterns_df['Société'], patterns_df['Pattern']))
        
        # Ajouter une colonne pour les patterns
        input_df['Email Pattern'] = input_df['Société'].map(patterns_dict)
        
        # Nettoyer les noms
        input_df['Prénom'] = input_df['Prénom'].apply(clean_name)
        input_df['Nom'] = input_df['Nom'].apply(clean_name)
        
        # Supprimer les lignes où les 4 colonnes valent 0
        cols_to_check = [
            'Years in Position',
            'Months in Position',
            'Years in Company',
            'Months in Company'
        ]
        if all(col in input_df.columns for col in cols_to_check):
            mask_zero = (input_df[cols_to_check] == 0).all(axis=1)
            input_df = input_df[~mask_zero].copy()

        # Créer une nouvelle colonne 'New Email' avec les valeurs de 'Email' existantes
        input_df['New Email'] = input_df['Email'].copy()
        
        # Générer les nouveaux emails là où nécessaire dans la nouvelle colonne
        mask = (input_df['Email'].isna()) | (input_df['Email Qualification'].astype(str).str.contains('catch_all@pro', na=False))
        input_df.loc[mask, 'New Email'] = input_df[mask].apply(
            lambda row: generate_email(row, row['Email Pattern']), axis=1
        )
        
        # Définir les différents cas
        generated_mask = (mask) & (input_df['New Email'] != '') & (input_df['New Email'] != input_df['Email'])
        failed_mask = (mask) & ((input_df['New Email'].isna()) | (input_df['New Email'] == ''))
        sac_mask = (mask) & (input_df['New Email'] == input_df['Email'])
        
        # Mettre à jour Email qualification selon les cas
        input_df.loc[generated_mask, 'Email Qualification'] = 'Generated'
        input_df.loc[failed_mask, 'Email Qualification'] = 'Not find'
        input_df.loc[sac_mask, 'Email Qualification'] = 'SAC'
        
        # Supprimer uniquement la colonne temporaire de pattern
        if 'Email Pattern' in input_df.columns:
            input_df = input_df.drop('Email Pattern', axis=1)
        
        # Vérifier que la colonne New Email est toujours présente
        if 'New Email' not in input_df.columns:
            print("⚠️ La colonne New Email a disparu!")
            return False
            
        # Sauvegarder le résultat final en spécifiant explicitement toutes les colonnes
        columns_to_save = [col for col in input_df.columns]
        input_df[columns_to_save].to_excel('cleaned_contacts.xlsx', index=False)

        # Calculer et afficher les statistiques détaillées
        total_generated = (input_df['Email Qualification'] == 'Generated').sum()
        total_not_find = (input_df['Email Qualification'] == 'Not find').sum()
        total_sac = (input_df['Email Qualification'] == 'SAC').sum()
        
        print("\n📊 Statistiques des emails générés :")
        print(f"✅ Generated : {total_generated}")
        print(f"❌ Not find : {total_not_find}")
        print(f"ℹ️ SAC : {total_sac}")
        print(f"📝 Total traité : {total_generated + total_not_find + total_sac}")
        
        return True
        
    except Exception as e:
        print(f"❌ Erreur lors du nettoyage: {str(e)}")
        return False

def analyze_email_patterns(filename=None):
    print("🔹 Analyse des patterns d'emails")
    
    if not filename:
        # Fallback pour l'utilisation en script indépendant si nécessaire
        filename = input("Entrez le nom du fichier Excel à analyser (ex: contacts.xlsx): ")
    
    if not os.path.exists(filename):
        print(f"❌ Erreur: {filename} non trouvé")
        return False
    
    try:
        df = pd.read_excel(filename)
        required_columns = ['Email', 'Prénom', 'Nom', 'Société']
        if not all(col in df.columns for col in required_columns):
            print("❌ Erreur: Le fichier doit contenir les colonnes: Email, Prénom, Nom, Société")
            return False
        
        df = df.dropna(subset=['Email', 'Prénom', 'Nom', 'Société'])
        
        patterns = []
        new_companies = set()  # Pour suivre les nouvelles entreprises
        
        for index, row in df.iterrows():
            try:
                if not isinstance(row['Email'], str):
                    print(f"⚠️ Ligne {index}: Email non valide")
                    continue
                    
                email = str(row['Email']).lower().strip()
                firstname = clean_name(row['Prénom'])
                lastname = clean_name(row['Nom'])
                company = str(row['Société']).lower().strip()
                
                if not email or not firstname or not lastname or not company:
                    print(f"⚠️ Ligne {index}: Données manquantes")
                    continue
                
                if '@' not in email:
                    print(f"⚠️ Ligne {index}: Format d'email invalide")
                    continue

                # Vérifier si l'email a la qualification "nominative@pro" ou "Generated"
                if not (('nominative@pro' in str(row['Email Qualification'])) or ('Generated' in str(row['Email Qualification']))):
                    continue
                
                local_part = email.split('@')[0]
                domain = email.split('@')[1]
                
                firstname_initial = firstname[0] if firstname else ''
                lastname_initial = lastname[0] if lastname else ''
                
                # On va tester chaque pattern dans l'ordre et s'arrêter au premier qui correspond
                pattern = local_part
                pattern_found = False
                
                patterns_to_try = [
                    # Avec des points (.)
                    (f"{firstname}.{lastname}", "firstname.lastname"),
                    (f"{firstname_initial}.{lastname}", "firstnameinitial.lastname"),
                    (f"{firstname}.{lastname_initial}", "firstname.lastnameinitial"),
                    (f"{firstname_initial}.{lastname_initial}", "firstnameinitial.lastnameinitial"),
                    (f"{lastname}.{firstname}", "lastname.firstname"),
                    (f"{lastname}.{firstname_initial}", "lastname.firstnameinitial"),
                    (f"{lastname_initial}.{firstname}", "lastnameinitial.firstname"),
                    (f"{lastname_initial}.{firstname_initial}", "lastnameinitial.firstnameinitial"),
                    
                    # Sans séparateurs (concaténation directe)
                    (f"{firstname}{lastname}", "firstnamelastname"),
                    (f"{firstname_initial}{lastname}", "firstnameinitiallastname"),
                    (f"{firstname}{lastname_initial}", "firstnamelastnameinitial"),
                    (f"{firstname_initial}{lastname_initial}", "firstnameinitiallastnameinitial"),
                    (f"{lastname}{firstname}", "lastnamefirstname"),
                    (f"{lastname_initial}{firstname}", "lastnameinitialfirstname"),
                    (f"{lastname}{firstname_initial}", "lastnamefirstnameinitial"),
                    (f"{lastname_initial}{firstname_initial}", "lastnameinitialfirstnameinitial"),
                    
                    # Pattern combiné avec points
                    (f"{firstname}.{lastname}.{firstname_initial}{lastname_initial}", "firstname.lastname.firstnameinitiallastnameinitial"),
                    
                    # Patterns individuels (de base)
                    (firstname, "firstname"),
                    (lastname, "lastname"),
                    (f"{firstname_initial}", "firstnameinitial"),
                    (f"{lastname_initial}", "lastnameinitial")
                ]
                
                for old, new in patterns_to_try:
                    if old == local_part:
                        pattern = new
                        pattern_found = True
                        break
                
                if not pattern_found:
                    print(f"⚠️ Ligne {index}: Pattern non reconnu pour {email}")
                    continue
                
                full_pattern = f"{pattern}@{domain.replace(company, 'company')}"
                
                patterns.append({
                    'Société': row['Société'],
                    'Pattern': full_pattern
                })
                
                # Ajouter l'entreprise à l'ensemble des nouvelles entreprises
                new_companies.add(row['Société'])
                
            except Exception as e:
                print(f"⚠️ Erreur ligne {index}: {str(e)}")
                continue
        
        if not patterns:
            print("❌ Aucun pattern valide n'a été trouvé")
            return False
        
        # Charger les patterns existants s'ils existent
        output_file = 'detected_patterns.xlsx'
        existing_companies = set()
        
        if os.path.exists(output_file):
            existing_patterns_df = pd.read_excel(output_file)
            existing_companies = set(existing_patterns_df['Société'])
            new_patterns_df = pd.DataFrame(patterns)
            
            # Combiner les anciens et nouveaux patterns
            patterns_df = pd.concat([existing_patterns_df, new_patterns_df])
            # Supprimer les doublons en gardant la première occurrence
            patterns_df = patterns_df.drop_duplicates(subset=['Société'], keep='first')
        else:
            # Si le fichier n'existe pas, créer un nouveau DataFrame
            patterns_df = pd.DataFrame(patterns).drop_duplicates(subset=['Société'])
        
        # Sauvegarder le résultat
        patterns_df.to_excel(output_file, index=False)
        
        # Calculer les nouvelles entreprises ajoutées
        newly_added_companies = new_companies - existing_companies
        
        print(f"✅ Patterns détectés et sauvegardés dans {output_file}")
        print(f"   {len(patterns_df)} patterns uniques au total")
        
        if newly_added_companies:
            print("\n📋 Nouvelles entreprises ajoutées :")
            for company in sorted(newly_added_companies):
                print(f"   • {company}")
        else:
            print("\nℹ️ Aucune nouvelle entreprise n'a été ajoutée")
        
        return True
        
    except Exception as e:
        print(f"❌ Erreur lors de l'analyse: {str(e)}")
        return False

def detect_email_pattern(emails, first_names, last_names):
    patterns = []
    for email, firstname, lastname in zip(emails, first_names, last_names):
        if pd.isna(email) or pd.isna(firstname) or pd.isna(lastname):
            continue
            
        email = email.lower()
        firstname = clean_name(firstname)
        lastname = clean_name(lastname)
        
        if not firstname or not lastname:
            continue
            
        firstname_initial = firstname[0]
        lastname_initial = lastname[0]
        
        local_part = email.split('@')[0]
        domain = email.split('@')[1]
        
        patterns_to_try = [
            # Avec des points (.)
            (f"{firstname}.{lastname}", "firstname.lastname"),
            (f"{firstname_initial}.{lastname}", "firstnameinitial.lastname"),
            (f"{firstname}.{lastname_initial}", "firstname.lastnameinitial"),
            (f"{firstname_initial}.{lastname_initial}", "firstnameinitial.lastnameinitial"),
            (f"{lastname}.{firstname}", "lastname.firstname"),
            (f"{lastname}.{firstname_initial}", "lastname.firstnameinitial"),
            (f"{lastname_initial}.{firstname}", "lastnameinitial.firstname"),
            (f"{lastname_initial}.{firstname_initial}", "lastnameinitial.firstnameinitial"),
            
            # Sans séparateurs (concaténation directe)
            (f"{firstname}{lastname}", "firstnamelastname"),
            (f"{firstname_initial}{lastname}", "firstnameinitiallastname"),
            (f"{firstname}{lastname_initial}", "firstnamelastnameinitial"),
            (f"{firstname_initial}{lastname_initial}", "firstnameinitiallastnameinitial"),
            (f"{lastname}{firstname}", "lastnamefirstname"),
            (f"{lastname_initial}{firstname}", "lastnameinitialfirstname"),
            (f"{lastname}{firstname_initial}", "lastnamefirstnameinitial"),
            (f"{lastname_initial}{firstname_initial}", "lastnameinitialfirstnameinitial"),
            
            # Pattern combiné avec points
            (f"{firstname}.{lastname}.{firstname_initial}{lastname_initial}", "firstname.lastname.firstnameinitiallastnameinitial"),
            
            # Patterns individuels (de base)
            (firstname, "firstname"),
            (lastname, "lastname"),
            (f"{firstname_initial}", "firstnameinitial"),
            (f"{lastname_initial}", "lastnameinitial")
        ]
        pattern_found = False
        for test_pattern, pattern_name in patterns_to_try:
            if test_pattern == local_part:
                patterns.append(f"{pattern_name}@{domain}")
                pattern_found = True
                break
                
        # End of Selection
    # Retourner le pattern le plus fréquent
    if patterns:
        return Counter(patterns).most_common(1)[0][0]
    return ''

def normalize_special_characters(filename=None, output_filename_web='normalized_contacts.xlsx'):
    print("🔹 Normalisation des caractères spéciaux :")
    
    is_web_call = filename is not None
    
    if not is_web_call:
        # Fallback pour l'utilisation en script indépendant si nécessaire
        filename = input("Entrez le nom du fichier Excel : ")
    
    if not os.path.exists(filename):
        print(f"❌ Erreur: {filename} non trouvé")
        return False
        
    try:
        # Charger le fichier
        df = pd.read_excel(filename)
        
    
        
        # Normaliser les colonnes First name et Last name
        df['Prénom'] = df['Prénom'].apply(lambda x: unidecode.unidecode(str(x)) if pd.notna(x) else x)
        df['Nom'] = df['Nom'].apply(lambda x: unidecode.unidecode(str(x)) if pd.notna(x) else x)
        
        # Déterminer le nom du fichier de sortie
        if is_web_call:
            final_output_filename = output_filename_web
        else:
            final_output_filename = f"{os.path.splitext(filename)[0]}_normalized.xlsx"

        # Sauvegarder le résultat
        df.to_excel(final_output_filename, index=False)
        print(f"✅ Caractères spéciaux normalisés.")
        
        return True
        
    except Exception as e:
        print(f"❌ Erreur lors de la normalisation: {str(e)}")
        return False

def main():
    while True:
        print("\n📋 Menu Principal:")
        print("1. Extraire les entreprises uniques")
        print("2. Analyser les patterns d'emails")
        print("3. Compléter les emails et recevoir le fichier prêt")
        print("4. Normaliser les caractères spéciaux")
        print("5. Quitter")
        
        choice = input("\nChoisissez une étape (1-5): ")
        
        if choice == '1':
            step1_extract_companies()
        elif choice == '2':
            analyze_email_patterns()
        elif choice == '3':
            step3_clean_and_complete()
        elif choice == '4':
            normalize_special_characters()
        elif choice == '5':
            print("Au revoir!")
            break
        else:
            print("❌ Choix invalide. Veuillez réessayer.")

if __name__ == "__main__":
    main()