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
required_columns = ['ID', 'Nom', 'Prénom', 'Date d\'enregistrement', 'Type de mission']
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

# === Fonction utilitaire pour retrouver toutes les courbes d'un pilote ===
def find_courbe_files(base_dir, identifiant):
    """
    Trouve tous les fichiers de courbes pour un identifiant donné.
    Retourne un dictionnaire {1: path1, 2: path2, 3: path3, 4: path4}
    """
    courbes = {}
    for i in range(1, 5):  # Courbes 1 à 4
        for ext in ['jpg', 'png']:
            path = os.path.join(base_dir, f"{identifiant}_courbe{i}.{ext}")
            if os.path.exists(path):
                courbes[i] = path
                break
    return courbes

# === Fonction pour déterminer le nom de mission selon le type ===
def get_mission_name(mission_type):
    """
    Retourne le nom de la mission selon le type :
    - AIR - GROUND -> Suippes
    - AIR - AIR -> Taxan
    """
    if mission_type and isinstance(mission_type, str):
        mission_type_clean = mission_type.strip().upper()
        if "AIR - GROUND" in mission_type_clean or "AIR-GROUND" in mission_type_clean:
            return "Suippes"
        elif "AIR - AIR" in mission_type_clean or "AIR-AIR" in mission_type_clean:
            return "Taxan"
    
    # Valeur par défaut si le type n'est pas reconnu
    return "Suippes"

# === Créer un PDF en mémoire avec ReportLab pour overlay ===
def create_overlay(nom, prenom, avion, map_name, mission_type, mission_name, instructeur, photo_path, courbes_dict, date_formatted):
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
    # Mission Type : utilise la valeur depuis l'Excel : X=160, Y=360
    can.drawString(163, 360, mission_type)
    # Mission Name : Suippes ou Taxan selon le type : X=180, Y=342
    can.drawString(179, 334, mission_name)
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
    
    # === Page 2 - Courbes 1 et 2 ===
    can.showPage()
    
    # Titres sur la page 2
    # Training Simulation Report - Shots Details en blanc : size 20, X=50, Y=800
    can.setFillColorRGB(1, 1, 1)  # Couleur blanche
    can.setFont("Comfortaa", 20)
    can.drawString(50, 800, "Training Simulation Report - Shots Details")
    
    # Round 1 en couleur #1C3062 (RGB : 28/255, 48/255, 98/255) : size 16, X=50, Y=765
    can.setFillColorRGB(28/255, 48/255, 98/255)
    can.setFont("Comfortaa", 16)
    can.drawString(50, 765, "Round 1")
    
    # Courbe 1 : X=50, Y=465, width=500, height=280
    if 1 in courbes_dict and os.path.exists(courbes_dict[1]):
        try:
            img1 = ImageReader(courbes_dict[1])
            can.drawImage(img1, x=50, y=465, width=500, height=280, mask='auto')
        except Exception as e:
            print(f"Erreur chargement courbe 1: {e}")
    
    # Courbe 2 : X=50, Y=150 (un peu en dessous de la courbe 1)
    if 2 in courbes_dict and os.path.exists(courbes_dict[2]):
        try:
            # Round 2 titre
            can.setFillColorRGB(28/255, 48/255, 98/255)
            can.setFont("Comfortaa", 16)
            can.drawString(50, 450, "Round 2")
            
            img2 = ImageReader(courbes_dict[2])
            can.drawImage(img2, x=50, y=150, width=500, height=280, mask='auto')
        except Exception as e:
            print(f"Erreur chargement courbe 2: {e}")
    
    # === Page 3 - Courbes 3 et 4 (seulement si elles existent) ===
    if 3 in courbes_dict or 4 in courbes_dict:
        can.showPage()
        
        # Titre page 3
        can.setFillColorRGB(1, 1, 1)  # Couleur blanche
        can.setFont("Comfortaa", 20)
        can.drawString(50, 800, "Training Simulation Report - Shots Details (Suite)")
        
        # Courbe 3 : X=50, Y=465, width=500, height=280 (même position que courbe 1)
        if 3 in courbes_dict and os.path.exists(courbes_dict[3]):
            try:
                # Round 3 titre
                can.setFillColorRGB(28/255, 48/255, 98/255)
                can.setFont("Comfortaa", 16)
                can.drawString(50, 765, "Round 3")
                
                img3 = ImageReader(courbes_dict[3])
                can.drawImage(img3, x=50, y=465, width=500, height=280, mask='auto')
            except Exception as e:
                print(f"Erreur chargement courbe 3: {e}")
        
        # Courbe 4 : X=50, Y=150 (même position que courbe 2)
        if 4 in courbes_dict and os.path.exists(courbes_dict[4]):
            try:
                # Round 4 titre
                can.setFillColorRGB(28/255, 48/255, 98/255)
                can.setFont("Comfortaa", 16)
                can.drawString(50, 450, "Round 4")
                
                img4 = ImageReader(courbes_dict[4])
                can.drawImage(img4, x=50, y=150, width=500, height=280, mask='auto')
            except Exception as e:
                print(f"Erreur chargement courbe 4: {e}")
    
    can.save()
    packet.seek(0)
    return packet

# === Fusionner overlay avec le template ===
def merge_overlay(template_path, overlay_stream, output_path, nb_courbes=0):
    reader_template = PdfReader(template_path)
    writer = PdfWriter()
    overlay_reader = PdfReader(overlay_stream)

    # Page 1 fusion
    base_page1 = reader_template.pages[0]
    overlay_page1 = overlay_reader.pages[0]
    base_page1.merge_page(overlay_page1)
    writer.add_page(base_page1)
    
    # Page 2 fusion
    if len(reader_template.pages) > 1:
        base_page2 = reader_template.pages[1]
        overlay_page2 = overlay_reader.pages[1]
        base_page2.merge_page(overlay_page2)
        writer.add_page(base_page2)
    
    # Gestion de la page 3
    if nb_courbes > 2:
        # Si on a plus de 2 courbes, inclure la page 3 avec overlay
        if len(overlay_reader.pages) > 2:
            if len(reader_template.pages) > 2:
                base_page3 = reader_template.pages[2]
                overlay_page3 = overlay_reader.pages[2]
                base_page3.merge_page(overlay_page3)
                writer.add_page(base_page3)
            else:
                # Si le template n'a pas de page 3, on prend juste l'overlay
                writer.add_page(overlay_reader.pages[2])
        
        # Copier les pages restantes du template (4, 5, etc.)
        for p in range(3, len(reader_template.pages)):
            writer.add_page(reader_template.pages[p])
    else:
        # Si on a <= 2 courbes, IGNORER la page 3 du template mais inclure les pages 4, 5, etc.
        for p in range(3, len(reader_template.pages)):
            writer.add_page(reader_template.pages[p])
    
    with open(output_path, 'wb') as f_out:
        writer.write(f_out)

# === Boucle de génération ===
for idx, row in df.iterrows():
    identifiant = str(row['ID'])
    nom = row['Nom']
    prenom = row['Prénom']
    
    # Récupération du type de mission depuis l'Excel
    type_mission = row['Type de mission'] if 'Type de mission' in row and not pd.isna(row['Type de mission']) else 'AIR - GROUND'
    
    # Détermination du nom de mission selon le type
    nom_mission = get_mission_name(type_mission)
    
    # Récupération et formatage de la date d'enregistrement
    date_enregistrement = row['Date d\'enregistrement'] if 'Date d\'enregistrement' in row else None
    date_formatted = format_date_from_excel(date_enregistrement)
    
    avion = "M-2000C"
    map_name = "Caucasus"
    instructeur = "ARESIA"

    photo_path = os.path.join(PHOTOS_DIR, f"{identifiant}.jpg")
    courbes_dict = find_courbe_files(COURBES_DIR, identifiant)
    nb_courbes = len(courbes_dict)
    
    print(f"\n--- Traitement de: {nom} {prenom} (ID: {identifiant})")
    print(f"Type de mission: {type_mission} -> Nom de mission: {nom_mission}")
    print(f"Date d'enregistrement: {date_enregistrement} -> {date_formatted}")
    print(f"Photo: {photo_path}")
    print(f"Courbes trouvées: {list(courbes_dict.keys())} (Total: {nb_courbes})")
    print(f"Page 3 sera {'incluse' if nb_courbes > 2 else 'supprimée'}")

    overlay_pdf = create_overlay(nom, prenom, avion, map_name, type_mission, nom_mission, instructeur, photo_path, courbes_dict, date_formatted)
    output_file = os.path.join(OUTPUT_DIR, f"{identifiant}_{nom}_{prenom}.pdf")
    merge_overlay(TEMPLATE_FILE, overlay_pdf, output_file, nb_courbes)
    print(f"PDF généré : {output_file}")

print("\nTous les PDFs ont été générés avec le template.")