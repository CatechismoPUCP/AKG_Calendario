# 📅 Generatore Excel Lezioni

Applicazione Streamlit per convertire testi di lezioni in file Excel formattati.

## 🚀 Installazione e Avvio

### 1. Installa le dipendenze
```bash
pip install -r requirements.txt
```

### 2. Avvia l'applicazione
```bash
streamlit run app.py
```

## 📋 Come usare l'applicazione

1. **Inserisci il testo delle lezioni** nel formato:
   ```
   Materia - Data ora_inizio - ora_fine - Modalità
   ```
   
   Esempio:
   ```
   AI: Intelligenza Artificiale 100% FAD - Modulo 1 - 21/08/2025 09:00 - 16:00 - Online
   ```

2. **Inserisci l'ID Sezione** (es. 141230)

3. **Inserisci il Codice Fiscale** (es. MLSNRS97S25F205C)

4. **Clicca "Genera Excel"** per creare il file

## 📦 Esportazione in .exe

Per creare un file .exe eseguibile:

### 1. Installa PyInstaller
```bash
pip install pyinstaller
```

### 2. Crea il file .exe
```bash
pyinstaller --onefile --windowed --add-data "C:\Users\[USERNAME]\AppData\Local\Programs\Python\Python[VERSION]\Lib\site-packages\streamlit\static;streamlit\static" app.py
```

### 3. Alternativa più semplice
Crea un file `build.bat`:
```batch
@echo off
echo Creazione file .exe in corso...
pyinstaller --onefile --noconsole --name "GeneratoreExcelLezioni" app.py
echo File .exe creato nella cartella dist/
pause
```

Poi esegui:
```bash
build.bat
```

## 🔧 Funzionalità

- ✅ Parsing automatico del testo delle lezioni
- ✅ Divisione automatica in slot orari di 1 ora
- ✅ Esclusione automatica della pausa pranzo (13:00-14:00)
- ✅ Formattazione Excel con celle di testo
- ✅ Download diretto del file Excel
- ✅ Interfaccia utente intuitiva

## 📊 Formato Output Excel

Il file Excel generato contiene le seguenti colonne:
- ID_SEZIONE
- DATA LEZIONE
- TOTALE_ORE (sempre 1)
- ORA_INIZIO
- ORA_FINE
- TIPOLOGIA (1 se ufficio 4 se online)
- CODICE FISCALE DOCENTE
- MATERIA
- CONTENUTI MATERIA
- SVOLGIMENTO SEDE LEZIONE (1 se ufficio; 4 se vuoto)

## ⚠️ Note Importanti

- Il formato della data deve essere DD/MM/YYYY
- Gli orari devono essere nel formato HH:MM
- L'applicazione salta automaticamente l'ora di pranzo (13:00-14:00)
- Tutte le celle Excel sono formattate come testo
