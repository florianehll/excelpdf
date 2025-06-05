import os
import io
import pandas as pd
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# === Enregistrer la police Comfortaa ===
# Placez le fichier TTF de Comfortaa dans un dossier 'fonts' à la racine du projet
BASE_DIR = os.getcwd()
fonts_dir = os.path.join(BASE_DIR, 'fonts')
comfortaa_ttf = os.path.join(fonts_dir, 'Comfortaa-Regular.ttf')
if not os.path.exists(comfortaa_ttf):
    raise FileNotFoundError("Le fichier Comfortaa-Regular.ttf doit se trouver dans le dossier 'fonts'.")
pdfmetrics.registerFont(TTFont('Comfortaa', comfortaa_ttf))

# === Chemins de base ===
DATA_DIR = os.path.join(BASE_DIR, 'data')
PHOTOS_DIR = os.path.join(DATA_DIR, 'photos')
COURBES_DIR = os.path.join(DATA_DIR, 'courbes')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
EXCEL_FILE = os.path.join(DATA_DIR, 'visiteurs-aresia.xlsx')
TEMPLATE_FILE = os.path.join(DATA_DIR, 'template.pdf')  # ton template de base
os.makedirs(OUTPUT_DIR, exist_ok=True)

# === Charger le fichier Excel ===
df = pd.read_excel(EXCEL_FILE)
required_columns = ['ID', 'Nom', 'Prénom', 'Date d\'enregistrement']
if not all(col in df.columns for col in required_columns):
    raise ValueError(f"Le fichier Excel doit contenir les colonnes : {required_columns}")
print("Colonnes détectées dans Excel :", df.columns.tolist())

# === Fonction utilitaire pour formater la date ===
def format_date_from_excel(date_str):
    """
    Convertit une date du format Excel (2025-05-28T14:14:21.712Z) 
    au format français (JJ/MM/AAAA)
    """
    try:
        # Si c'est déjà un objet datetime pandas
        if isinstance(date_str, pd.Timestamp):
            return date_str.strftime("%d/%m/%Y")
        
        # Si c'est une chaîne de caractères
        if isinstance(date_str, str):
            # Supprimer le 'Z' à la fin et parser la date ISO
            date_clean = date_str.replace('Z', '')
            if 'T' in date_clean:
                # Format ISO avec heure
                dt = datetime.fromisoformat(date_clean)
            else:
                # Format date simple
                dt = datetime.strptime(date_clean, "%Y-%m-%d")
            return dt.strftime("%d/%m/%Y")
        
        # Si c'est un objet datetime
        if isinstance(date_str, datetime):
            return date_str.strftime("%d/%m/%Y")
            
        # Fallback - retourner une date par défaut
        print(f"Format de date non reconnu: {date_str}, utilisation de la date actuelle")
        return datetime.now().strftime("%d/%m/%Y")
        
    except Exception as e:
        print(f"Erreur lors du formatage de la date '{date_str}': {e}")
        return datetime.now().strftime("%d/%m/%Y")

# === Fonction utilitaire pour retrouver la bonne extension de courbe ===
def find_courbe_file(base_dir, identifiant):
    for ext in ['jpg', 'png']:
        path = os.path.join(base_dir, f"{identifiant}_courbe.{ext}")
        if os.path.exists(path):
            return path
    return None

# === Créer un PDF en mémoire avec ReportLab pour overlay ===
def create_overlay(nom, prenom, avion, map_name, mission, instructeur, photo_path, courbe_path, date_formatted):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    
    # === Page 1 ===
    # Couleur de texte blanche pour les informations fixes
    can.setFillColorRGB(1, 1, 1)
    font_size = 16
    can.setFont("Comfortaa", font_size)
    
    # Nom Pilote (Nom + Prénom) : X=99, Y=530
    can.drawString(99, 533, f"{nom} {prenom}")
    # Instructeur : X=99, Y=458
    can.drawString(99, 458, f"{instructeur}")
    # Avion (juste le modèle) : X=98, Y=309
    can.drawString(98, 309, f"{avion}")
    # Map (seulement le nom) : X=83, Y=284
    can.drawString(83, 284, f"{map_name}")
    # Mission Type : AIR-GROUND : X=160, Y=360
    can.drawString(163, 360, "AIR-GROUND")
    # Mission Name : Suippes : X=180, Y=342
    can.drawString(179, 338, "Suippes")
    # Date : utilisation de la date formatée depuis l'Excel : X=84, Y=490
    can.drawString(86, 385, date_formatted)
    
    # Photo : coins (360,278) bas-gauche -> (565,560) haut-droit
    if os.path.exists(photo_path):
        try:
            img = ImageReader(photo_path)
            can.drawImage(img, x=360, y=278, width=205, height=280, mask='auto')
        except Exception as e:
            print(f"Erreur chargement photo: {e}")
    else:
        print(f"Photo non trouvée: {photo_path}")
    
    # Passer à la page 2 pour insérer la courbe et les titres
    can.showPage()
    
    # Sur la page 2, écrire les titres
    # Training Simulation Report - Shots Details en blanc : size 20, X=50, Y=800
    can.setFillColorRGB(1, 1, 1)  # Couleur blanche
    can.setFont("Comfortaa", 20)
    can.drawString(50, 800, "Training Simulation Report - Shots Details")
    
    # Round 1 en couleur #1C3062 (RGB : 28/255, 48/255, 98/255) : size 16, X=50, Y=765
    can.setFillColorRGB(28/255, 48/255, 98/255)
    can.setFont("Comfortaa", 16)
    can.drawString(50, 765, "Round 1")
    
    # Courbe : X=50, Y=500, width=500, height=280
    if courbe_path and os.path.exists(courbe_path):
        try:
            img2 = ImageReader(courbe_path)
            can.drawImage(img2, x=50, y=465, width=500, height=280, mask='auto')
        except Exception as e:
            print(f"Erreur chargement courbe: {e}")
    else:
        print(f"Courbe non trouvée: {courbe_path}")
    
    can.save()
    packet.seek(0)
    return packet

# === Fusionner overlay avec le template ===
def merge_overlay(template_path, overlay_stream, output_path):
    reader_template = PdfReader(template_path)
    writer = PdfWriter()
    overlay_reader = PdfReader(overlay_stream)

    # Page 1 fusion
    base_page1 = reader_template.pages[0]
    overlay_page1 = overlay_reader.pages[0]
    base_page1.merge_page(overlay_page1)
    writer.add_page(base_page1)
    
    # Page 2 fusion: si le template a page 2
    if len(reader_template.pages) > 1:
        base_page2 = reader_template.pages[1]
        overlay_page2 = overlay_reader.pages[1]
        base_page2.merge_page(overlay_page2)
        writer.add_page(base_page2)
    
    # Copier les pages restantes du template si besoin
    for p in range(2, len(reader_template.pages)):
        writer.add_page(reader_template.pages[p])
    
    with open(output_path, 'wb') as f_out:
        writer.write(f_out)

# === Boucle de génération ===
for idx, row in df.iterrows():
    identifiant = str(row['ID'])
    nom = row['Nom']
    prenom = row['Prénom']
    mission = row['Mission'] if 'Mission' in row and not pd.isna(row['Mission']) else ''
    
    # Récupération et formatage de la date d'enregistrement
    date_enregistrement = row['Date d\'enregistrement'] if 'Date d\'enregistrement' in row else None
    date_formatted = format_date_from_excel(date_enregistrement)
    
    avion = "M-2000C"
    map_name = "Caucasus"
    instructeur = "ARESIA"

    photo_path = os.path.join(PHOTOS_DIR, f"{identifiant}.jpg")
    courbe_path = find_courbe_file(COURBES_DIR, identifiant)
    
    print(f"\n--- Traitement de: {nom} {prenom} (ID: {identifiant})")
    print(f"Date d'enregistrement: {date_enregistrement} -> {date_formatted}")
    print(f"Photo: {photo_path}")
    print(f"Courbe: {courbe_path if courbe_path else 'non trouvée'}")

    overlay_pdf = create_overlay(nom, prenom, avion, map_name, mission, instructeur, photo_path, courbe_path, date_formatted)
    output_file = os.path.join(OUTPUT_DIR, f"{identifiant}_{nom}_{prenom}.pdf")
    merge_overlay(TEMPLATE_FILE, overlay_pdf, output_file)
    print(f"PDF généré : {output_file}")

print("\nTous les PDFs ont été générés avec le template.")