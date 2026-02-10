import pandas as pd
import qrcode
from fpdf import FPDF
import os

# --- 1. SCELTA DEL FILE ---
# Ora puoi scrivere 'torino.xlsx' o 'milano.xlsx' quando lanci il programma
nome_input = input("Inserisci il nome del file Excel (es. torino.xlsx): ")

# Ricaviamo il nome senza estensione per usarlo nel nome del PDF finale
base_name = os.path.splitext(nome_input)[0]

try:
    df = pd.read_excel(nome_input)
except Exception as e:
    print(f"Errore: Impossibile trovare o leggere {nome_input}. Assicurati che sia nella cartella Etichette.")
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
    # --- 1. DEFINISCI IL LINK ---
    # Sostituisci 'tuo-utente' con il tuo nome GitHub e 'tuo-repo' con il nome del progetto
    link_sito = "https://riccardobaima.github.io/Etichette/" 

    # --- 2. GENERA IL QR CON IL LINK (Invece del testo semplice) ---
    img = qrcode.make(link_sito)
    qr_path = f"temp_qr/qr_{index}.png"
    img.save(qr_path)

    # --- 3. DISEGNO (Questa è la sezione che hai indicato tu) ---
    # Non serve cambiare nulla qui, perché qr_path ora punta a un QR col link!
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

# Salviamo il PDF col nome dinamico (es: etichette_torino.pdf)
nome_output = f"etichette_{base_name}.pdf"
pdf.output(nome_output)
print(f"PDF GENERATO: {nome_output}")