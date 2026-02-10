import pandas as pd
import qrcode
from fpdf import FPDF
import os

# --- 1. SCELTA DEL FILE ---
nome_input = input("Inserisci il nome del file Excel (es. torino.xlsx): ")
base_name = os.path.splitext(nome_input)[0]

try:
    df = pd.read_excel(nome_input)
except Exception as e:
    print(f"Errore: Impossibile trovare o leggere {nome_input}.")
    exit()

# --- PULIZIA DATE ---
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
    
    # --- LOGICA LINK DINAMICO ---
    # Qui costruiamo il link che la pagina HTML leggerà
    link_personalizzato = (
        f"https://riccardobaima.github.io/Etichette/?"
        f"ditta={ditta}&"
        f"cellula={row['CellulaID']}&"
        f"scadenza={row['DataScadenza']}"
    ).replace(" ", "%20")

    # --- GENERAZIONE QR ---
    img = qrcode.make(link_personalizzato)
    qr_path = f"temp_qr/qr_{index}.png"
    img.save(qr_path)

    # --- DISEGNO ETICHETTA ---
    pdf.rect(x_attuale, y_attuale, 40, 100) 

    if os.path.exists('logo.jpg'):
        pdf.image('logo.jpg', x=x_attuale + 10, y=y_attuale + 3, w=20)
    
    pdf.set_font("Helvetica", 'B', 9)
    pdf.set_xy(x_attuale, y_attuale + 18)
    pdf.multi_cell(40, 5, text=ditta, align='C')
    
    pdf.set_font("Helvetica", 'B', 10)
    pdf.set_xy(x_attuale + 2, y_attuale + 32) 
    pdf.cell(36, 5, text=cellula_info, align='L')
        
    pdf.set_font("Helvetica", '', 8)
    pdf.set_xy(x_attuale + 2, y_attuale + 40)
    pdf.multi_cell(36, 4, text=descrizione, align='L')
    
    # Inserimento del QR generato col link
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

nome_output = f"etichette_{base_name}.pdf"
pdf.output(nome_output)
print(f"PDF GENERATO: {nome_output}")