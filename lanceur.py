import sys
import os
import subprocess
import time
import webbrowser

# --- 1. CONFIGURATION ---
nom_fichier_app = "app_bourse.py"
dossier_actuel = os.getcwd()
chemin_app = os.path.join(dossier_actuel, nom_fichier_app)

print(f"--- DÉMARRAGE ROBUSTE DE {nom_fichier_app} ---")

# --- 2. FORÇAGE DES PARAMÈTRES (Anti-Email) ---
os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"

# --- 3. LANCEMENT ---
if not os.path.exists(chemin_app):
    print(f"ERREUR : Le fichier {nom_fichier_app} est introuvable !")
    sys.exit()

# Commande simplifiée
cmd = [sys.executable, "-m", "streamlit", "run", nom_fichier_app]

try:
    print("1. Démarrage du serveur...")
    # On lance le processus
    process = subprocess.Popen(cmd, cwd=dossier_actuel)
    
    print("2. Attente de l'initialisation (attendez 5 secondes)...")
    time.sleep(5) 
    
    # --- 4. OUVERTURE NAVIGATEUR ---
    url = "http://localhost:8501"
    print(f"3. Ouverture de {url}")
    webbrowser.open(url)
    
    print("\n✅ APPLICATION LANCÉE !")
    print("Laissez ce script tourner (bouton rouge allumé) pendant que vous naviguez.")

except Exception as e:
    print(f"Erreur : {e}")