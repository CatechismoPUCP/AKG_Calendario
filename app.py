import streamlit as st
import pandas as pd
import re
from datetime import datetime, time
from io import BytesIO

def parse_lesson_text(text):
    """
    Parsa il testo delle lezioni e estrae le informazioni necessarie
    """
    lessons = []
    lines = text.strip().split('\n')
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Pattern per estrarre le informazioni dalla riga
        # Formato: "Materia - Data ora_inizio - ora_fine - Modalità"
        pattern = r'(.+?)\s-\s(\d{2}/\d{2}/\d{4})\s(\d{2}:\d{2})\s-\s(\d{2}:\d{2})\s-\s(.+)'
        match = re.match(pattern, line)
        
        if match:
            materia = match.group(1).strip()
            data = match.group(2)
            ora_inizio = match.group(3)
            ora_fine = match.group(4)
            modalita = match.group(5).strip()
            
            lessons.append({
                'materia': materia,
                'data': data,
                'ora_inizio': ora_inizio,
                'ora_fine': ora_fine,
                'modalita': modalita
            })
    
    return lessons

def generate_hourly_rows(lessons, id_sezione, codice_fiscale):
    """
    Genera le righe orarie per l'Excel basandosi sulle lezioni
    """
    rows = []
    
    for lesson in lessons:
        # Converte gli orari in oggetti time per calcolare le ore
        start_time = datetime.strptime(lesson['ora_inizio'], '%H:%M').time()
        end_time = datetime.strptime(lesson['ora_fine'], '%H:%M').time()
        
        # Converte in minuti per calcolare la durata
        start_minutes = start_time.hour * 60 + start_time.minute
        end_minutes = end_time.hour * 60 + end_time.minute
        
        # Genera una riga per ogni ora
        current_minutes = start_minutes
        while current_minutes < end_minutes:
            current_hour = current_minutes // 60
            current_minute = current_minutes % 60
            next_hour = (current_minutes + 60) // 60
            next_minute = (current_minutes + 60) % 60
            
            # Se l'ora successiva supera l'ora di fine, usa l'ora di fine
            if current_minutes + 60 > end_minutes:
                next_hour = end_time.hour
                next_minute = end_time.minute
            
            # Salta l'ora di pausa pranzo (13:00-14:00)
            if current_hour == 13:
                current_minutes += 60
                continue
                
            ora_inizio_slot = f"{current_hour:02d}:{current_minute:02d}"
            ora_fine_slot = f"{next_hour:02d}:{next_minute:02d}"
            
            # Determina tipologia e sede in base alla modalità della lezione
            if lesson['modalita'].lower() == 'ufficio':
                tipologia = '1'
                sede_lezione = '1'
            else:
                tipologia = '4'
                sede_lezione = ''
            
            row = {
                'ID_SEZIONE': str(id_sezione),
                'DATA LEZIONE': lesson['data'],
                'TOTALE_ORE': '1',
                'ORA_INIZIO': ora_inizio_slot,
                'ORA_FINE': ora_fine_slot,
                'TIPOLOGIA': tipologia,
                'CODICE FISCALE DOCENTE': codice_fiscale,
                'MATERIA': lesson['materia'],
                'CONTENUTI MATERIA': lesson['materia'],
                'SVOLGIMENTO SEDE LEZIONE': sede_lezione
            }
            
            rows.append(row)
            current_minutes += 60
    
    return rows

def create_excel_file(rows):
    """
    Crea il file Excel con la formattazione corretta (tutte le celle come testo)
    """
    df = pd.DataFrame(rows)
    
    # Crea un buffer in memoria per il file Excel
    output = BytesIO()
    
    # Usa xlsxwriter per avere più controllo sulla formattazione
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Lezioni', index=False)
        
        # Ottieni il workbook e il worksheet
        workbook = writer.book
        worksheet = writer.sheets['Lezioni']
        
        # Definisci il formato per le celle come testo
        text_format = workbook.add_format({'num_format': '@'})
        
        # Applica il formato testo a tutte le celle con dati
        for row_num in range(len(df) + 1):  # +1 per includere l'header
            for col_num in range(len(df.columns)):
                worksheet.write(row_num, col_num, 
                              df.iloc[row_num-1, col_num] if row_num > 0 else df.columns[col_num], 
                              text_format)
        
        # Adatta la larghezza delle colonne
        for i, col in enumerate(df.columns):
            max_length = max(
                df[col].astype(str).map(len).max(),
                len(col)
            )
            worksheet.set_column(i, i, max_length + 2)
    
    output.seek(0)
    return output

def main():
    st.set_page_config(
        page_title="Generatore Excel Lezioni",
        page_icon="📅",
        layout="wide"
    )
    
    # Stile CSS personalizzato
    st.markdown("""
    <style>
        /* Stile generale */
        .main {
            background-color: #ffffff;
        }
        
        /* Header */
        .stApp > header {
            background-color: #4CAF50;
            color: white;
        }
        
        /* Titolo */
        h1 {
            color: #2E7D32;
        }
        
        /* Sottotitoli */
        h2, h3, h4, h5, h6 {
            color: #388E3C;
        }
        
        /* Pulsanti */
        .stButton > button {
            background-color: #4CAF50;
            color: white;
            border: none;
            border-radius: 4px;
            padding: 0.5rem 1rem;
        }
        
        .stButton > button:hover {
            background-color: #388E3C;
            color: white;
        }
        
        /* Input fields */
        .stTextInput > div > div > input,
        .stTextArea > textarea {
            border: 1px solid #81C784;
            border-radius: 4px;
        }
        
        /* Sidebar */
        .css-1d391kg {
            background-color: #E8F5E9;
        }
        
        /* Tabella */
        .stDataFrame {
            border: 1px solid #81C784;
            border-radius: 4px;
        }
        
        /* Footer */
        footer {
            visibility: hidden;
        }
    </style>
    """, unsafe_allow_html=True)
    
    st.title("📅 Generatore Excel Lezioni")
    st.markdown("---")
    
    # Input del testo delle lezioni
    st.subheader("📝 Inserisci il testo delle lezioni")
    lesson_text = st.text_area(
        "Incolla qui il testo delle lezioni:",
        height=200,
        placeholder="AI: Intelligenza Artificiale 100% FAD - Modulo 1 - 21/08/2025 09:00 - 16:00 - Online\n..."
    )
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🆔 ID Sezione")
        id_sezione = st.text_input("Inserisci l'ID sezione:", placeholder="141230")
    
    with col2:
        st.subheader("👤 Codice Fiscale")
        codice_fiscale = st.text_input("Inserisci il codice fiscale:", placeholder="MLSNRS97S25F205C")
    
    if st.button("🚀 Genera Excel", type="primary"):
        if not lesson_text.strip():
            st.error("⚠️ Inserisci il testo delle lezioni!")
            return
        
        if not id_sezione.strip():
            st.error("⚠️ Inserisci l'ID sezione!")
            return
            
        if not codice_fiscale.strip():
            st.error("⚠️ Inserisci il codice fiscale!")
            return
        
        try:
            # Parsa le lezioni
            lessons = parse_lesson_text(lesson_text)
            
            if not lessons:
                st.error("⚠️ Nessuna lezione trovata nel testo. Verifica il formato!")
                return
            
            st.success(f"✅ Trovate {len(lessons)} lezioni!")
            
            # Genera le righe per l'Excel
            rows = generate_hourly_rows(lessons, id_sezione, codice_fiscale)
            
            st.info(f"📊 Generate {len(rows)} righe per l'Excel")
            
            # Crea il file Excel
            excel_file = create_excel_file(rows)
            
            # Mostra anteprima
            st.subheader("👀 Anteprima dati")
            df_preview = pd.DataFrame(rows)
            st.dataframe(df_preview.head(10), use_container_width=True)
            
            if len(rows) > 10:
                st.info(f"Mostrate prime 10 righe di {len(rows)} totali")
            
            # Download button
            st.download_button(
                label="📥 Scarica Excel",
                data=excel_file,
                file_name=f"lezioni_{id_sezione}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"❌ Errore durante la generazione: {str(e)}")
    
    # Sezione di aiuto
    with st.expander("ℹ️ Formato del testo richiesto"):
        st.markdown("""
        **Formato richiesto per ogni riga:**
        ```
        Materia - Data ora_inizio - ora_fine - Modalità
        ```
        
        **Esempio:**
        ```
        AI: Intelligenza Artificiale 100% FAD - Modulo 1 - 21/08/2025 09:00 - 16:00 - Online
        ```
        
        **Note:**
        - La data deve essere nel formato DD/MM/YYYY
        - Gli orari devono essere nel formato HH:MM
        - L'ora di pausa pranzo (13:00-14:00) viene automaticamente saltata
        - Ogni ora viene divisa in slot di 1 ora per l'Excel
        """)

if __name__ == "__main__":
    main()
