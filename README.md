# CVC – Scheda Valutazione Allievi

App web per la gestione delle valutazioni degli allievi del corso di vela CVC.

---

## Struttura
```
cvc_app/
├── app.py              # Backend Flask
├── requirements.txt    # Dipendenze Python
├── Procfile            # Comando di avvio (Railway/Render)
├── static/
│   └── index.html      # Frontend completo
└── README.md
```

---

## Deploy su Railway (consigliato, gratuito)

1. Crea account su https://railway.app
2. Clicca **"New Project"** → **"Deploy from GitHub"**
   - Oppure usa **"Deploy from local directory"** e trascina questa cartella
3. Railway rileva automaticamente Python e installa le dipendenze
4. Vai su **Variables** e aggiungi:
   - `ADMIN_PASSWORD` = la tua password admin (es. `CVC2025!`)
   - `PORT` = `5000` (di solito Railway lo imposta da solo)
5. Clicca **Deploy** — in 2 minuti il sito è online
6. Vai su **Settings → Domains** per ottenere il link pubblico (es. `cvc-voti.railway.app`)

---

## Deploy su Render (alternativa gratuita)

1. Crea account su https://render.com
2. Clicca **"New"** → **"Web Service"**
3. Collega il repository GitHub oppure carica i file
4. Configura:
   - **Runtime**: Python 3
   - **Build command**: `pip install -r requirements.txt`
   - **Start command**: `gunicorn app:app --bind 0.0.0.0:$PORT`
5. Aggiungi variabile d'ambiente `ADMIN_PASSWORD`
6. Clicca **Create Web Service**

---

## Variabili d'ambiente

| Variabile        | Default    | Descrizione                        |
|------------------|------------|------------------------------------|
| `ADMIN_PASSWORD` | `admin123` | Password per accesso admin         |
| `PORT`           | `5000`     | Porta del server (auto su Railway) |
| `DB_PATH`        | `cvc.db`   | Percorso del database SQLite       |

⚠️ Cambia sempre `ADMIN_PASSWORD` prima di mettere online!

---

## Uso locale (per test)

```bash
pip install -r requirements.txt
python app.py
# Apri http://localhost:5000
```

---

## Funzionalità

- **Inserimento voti**: form con 7 criteri × 7 giorni, calcolo punteggio in tempo reale
- **Riepilogo**: tabella con filtri per allievo, istruttore e corso (solo admin)
- **Export CSV**: scarica tutti i dati in formato Excel-compatibile
- **Admin**: statistiche generali e per corso

---

## Note tecniche

- Database: SQLite (incluso in Python, zero configurazione)
- Backend: Flask + Gunicorn
- Frontend: HTML/CSS/JS puro, nessun framework esterno
- Autenticazione: token di sessione semplice per l'area admin
