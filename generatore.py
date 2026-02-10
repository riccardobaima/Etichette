import pandas as pd
import qrcode
from fpdf import FPDF
import os

# 1. CARICAMENTO
try:
    df = pd.read_excel('torino.xlsx')
except FileNotFoundError:
    df = pd.read_csv('torino.xlsx - Foglio1.csv')

# --- PULIZIA DATE (Rimuove l'orario 00:00:00) ---
# Trasformiamo le colonne in date vere e poi in testo formatato GG/MM/AAAA
for col in ['DataRilascio', 'DataScadenza']:
    df[col] = pd.to_datetime(df[col]).dt.strftime('%d/%m/%Y')

# 2. Setup PDF
pdf = FPDF(orientation='P', unit='mm', format='A4')
pdf.set_auto_page_break(auto=True, margin=10)
pdf.add_page()

if not os.path.exists('temp_qr'):
    os.makedirs('temp_qr')

x_attuale, y_attuale = 5, 10
contatore_colonna = 0

for index, row in df.iterrows():
    # --- DATI ---
    ditta = str(row['Ditta'])
    cellula_info = f"Cellula: {row['CellulaID']} - {row['Posizione']}"
    descrizione = str(row['Descrizione'])
    
    # Ora row['DataRilascio'] è già pulita senza l'orario
    testo_qr = (
        f"Ente rilasciante: {row['Comune']}\n"
        f"N° autorizzazione: {row['ProtocolloEnte']}\n"
        f"Data rilascio: {row['DataRilascio']}\n"
        f"Data scadenza: {row['DataScadenza']}"
    )

    # Genera QR
    img = qrcode.make(testo_qr)
    qr_path = f"temp_qr/qr_{index}.png"
    img.save(qr_path)

    # --- DISEGNO ---
    pdf.rect(x_attuale, y_attuale, 40, 100) 

    # 1. LOGO AZIENDALE
    if os.path.exists('logo.jpg'):
        pdf.image('logo.jpg', x=x_attuale + 10, y=y_attuale + 3, w=20)
    
    # 2. DITTA
    pdf.set_font("Helvetica", 'B', 9)
    pdf.set_xy(x_attuale, y_attuale + 18)
    pdf.multi_cell(40, 5, text=ditta, align='C')
    
    # 3. CELLULA
    pdf.set_font("Helvetica", 'B', 10)
    pdf.set_xy(x_attuale + 2, y_attuale + 32) 
    pdf.cell(36, 5, text=cellula_info, align='L')
        
    # 4. DESCRIZIONE
    pdf.set_font("Helvetica", '', 8)
    pdf.set_xy(x_attuale + 2, y_attuale + 40)
    pdf.multi_cell(36, 4, text=descrizione, align='L')
    
    # 5. QR CODE
    pdf.image(qr_path, x=x_attuale + 5, y=y_attuale + 68, w=30)

    # Logica Griglia
    contatore_colonna += 1
    if contatore_colonna < 5:
        x_attuale += 40
    else:
        contatore_colonna = 0
        x_attuale = 5
        y_attuale += 100
        if y_attuale + 100 > 280:
            pdf.add_page()
            y_attuale = 10

pdf.output("etichette_torino_finale.pdf")
print("PDF GENERATO! Date pulite correttamente.")