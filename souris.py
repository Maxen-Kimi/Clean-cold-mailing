import pyautogui
import time

# Positions proches des coins (à ajuster selon la taille de ton écran)
positions = [
    (20, 20),        # Haut gauche
    (1900, 20),      # Haut droit
    (20, 1060),      # Bas gauche
    (1900, 1060)     # Bas droit
]

interval = 0.1  # secondes
duree_max = 2 * 60 * 60  # 2 heures en secondes
start_time = time.time()

while time.time() - start_time < duree_max:
    for pos in positions:
        pyautogui.moveTo(pos[0], pos[1], duration=0.25)
        time.sleep(interval)
        if time.time() - start_time > duree_max:
            break
