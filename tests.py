import unittest
import os
import pandas as pd
from app import app, allowed_file, validate_excel_file
from clean_cold_mailing_script import (
    clean_name, generate_email, normalize_special_characters
)

class TestApp(unittest.TestCase):
    def setUp(self):
        self.app = app.test_client()
        self.app.testing = True
        self.test_file = 'test.xlsx'
        
        # Créer un fichier Excel de test
        df = pd.DataFrame({
            'Name': ['John Doe', 'Jane Smith'],
            'Email': ['john@example.com', 'jane@example.com']
        })
        df.to_excel(self.test_file, index=False)
    
    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
    
    def test_allowed_file(self):
        """Test de la validation des extensions de fichiers."""
        self.assertTrue(allowed_file('test.xlsx'))
        self.assertFalse(allowed_file('test.txt'))
        self.assertFalse(allowed_file('test'))
    
    def test_validate_excel_file(self):
        """Test de la validation des fichiers Excel."""
        self.assertTrue(validate_excel_file(self.test_file))
        
        # Créer un fichier texte et tester
        with open('test.txt', 'w') as f:
            f.write('test')
        self.assertFalse(validate_excel_file('test.txt'))
        os.remove('test.txt')
    
    def test_clean_name(self):
        """Test de la fonction de nettoyage des noms."""
        self.assertEqual(clean_name('John Doe'), 'John Doe')
        self.assertEqual(clean_name('  John  Doe  '), 'John Doe')
        self.assertEqual(clean_name('John-Doe'), 'John Doe')
    
    def test_generate_email(self):
        """Test de la génération d'emails."""
        self.assertEqual(
            generate_email('John', 'Doe', 'example.com'),
            'john.doe@example.com'
        )
        self.assertEqual(
            generate_email('John', 'Doe', 'example.com', 'jdoe'),
            'jdoe@example.com'
        )
    
    def test_normalize_special_characters(self):
        """Test de la normalisation des caractères spéciaux."""
        test_str = "Café & Crème"
        normalized = normalize_special_characters(test_str)
        self.assertEqual(normalized, "Cafe & Creme")
    
    def test_login_required(self):
        """Test de la protection des routes."""
        response = self.app.get('/')
        self.assertEqual(response.status_code, 302)  # Redirection vers login
    
    def test_upload_file(self):
        """Test de l'upload de fichier."""
        with open(self.test_file, 'rb') as f:
            response = self.app.post(
                '/upload',
                data={'file': (f, self.test_file)},
                content_type='multipart/form-data'
            )
        self.assertEqual(response.status_code, 401)  # Non authentifié

if __name__ == '__main__':
    unittest.main() 