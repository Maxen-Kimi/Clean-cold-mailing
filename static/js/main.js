
let currentFilename = '';

document.getElementById('uploadForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    const uploadStatus = document.getElementById('uploadStatus');
    
    if (!file) {
        showAlert(uploadStatus, 'Veuillez sélectionner un fichier', 'danger');
        return;
    }
    
    const formData = new FormData();
    formData.append('file', file);
    
    try {
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (response.ok) {
            currentFilename = data.filename;
            showAlert(uploadStatus, data.message, 'success');
            document.getElementById('downloadButtons').style.display = 'block';
        } else {
            showAlert(uploadStatus, data.error, 'danger');
        }
    } catch (error) {
        showAlert(uploadStatus, 'Erreur lors du téléchargement du fichier', 'danger');
    }
});

async function processFile(action) {
    if (!currentFilename) {
        showAlert(document.getElementById('status'), 'Veuillez d\'abord télécharger un fichier', 'danger');
        return;
    }
    
    const status = document.getElementById('status');
    showAlert(status, 'Traitement en cours...', 'info');
    
    try {
        const response = await fetch('/process', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                action: action,
                filename: currentFilename
            })
        });
        
        const data = await response.json();
        
        if (response.ok) {
            showAlert(status, data.message, 'success');
        } else {
            showAlert(status, data.error, 'danger');
        }
    } catch (error) {
        showAlert(status, 'Erreur lors du traitement du fichier', 'danger');
    }
}

async function downloadFile(filename) {
    try {
        const response = await fetch(`/download/${filename}`);
        
        if (response.ok) {
            // Créer un lien temporaire pour le téléchargement
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
        } else {
            const data = await response.json();
            showAlert(document.getElementById('status'), data.error, 'danger');
        }
    } catch (error) {
        showAlert(document.getElementById('status'), 'Erreur lors du téléchargement du fichier', 'danger');
    }
}

function showAlert(element, message, type) {
    element.textContent = message;
    element.className = `alert alert-${type}`;
    element.style.display = 'block';
    
    // Masquer l'alerte après 5 secondes
    setTimeout(() => {
        element.style.display = 'none';
    }, 5000);
} 
