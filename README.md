# Archiva Project Sizing

Tool interno per il sizing parametrico dei progetti in fase di quotazione.

## Stack

- **Backend**: FastAPI + Python 3.12
- **Database**: PostgreSQL (Railway Plugin)
- **Auth**: Utenti in JSON su Railway Volume, sessioni in-memory, SHA-256
- **Frontend**: HTML single-file, nessun framework

## Deploy su Railway

### 1. Crea il progetto
```bash
# Fork/push su GitHub, poi connetti Railway al repo
```

### 2. Aggiungi il plugin PostgreSQL
In Railway → progetto → **+ New** → **Database** → **PostgreSQL**

### 3. Collega DATABASE_URL
Railway → servizio → **Variables** → aggiungi:
```
DATABASE_URL=${{Postgres.DATABASE_URL}}
DEFAULT_PASSWORD=archiva2026
```

### 4. Aggiungi un Volume (per gli utenti)
Railway → servizio → **Volumes** → mount path: `/data`

La variabile `RAILWAY_VOLUME_MOUNT_PATH` viene impostata automaticamente da Railway.

### 5. Deploy
Push su main → Railway fa la build e deploya automaticamente.

## Utenti iniziali

Al primo avvio vengono creati automaticamente:

| Nome | Username | Ruolo |
|------|----------|-------|
| Marco Pastore | marco.pastore | admin |
| Paolo Gandini | paolo.gandini | admin |
| Chiara Pettenuzzo | chiara.pettenuzzo | admin |

Password default: `archiva2026` (cambiare subito dalla pagina Admin)

## Sviluppo locale

```bash
pip install -r requirements.txt

# Serve un PostgreSQL locale oppure usa SQLite per sviluppo
export DATABASE_URL=postgresql://localhost/sizing_dev

uvicorn main:app --reload --port 8000
```

## Struttura

```
project-sizing/
├── main.py           # FastAPI app + tutti gli endpoint
├── auth.py           # Autenticazione (identico alle altre app Archiva)
├── database.py       # Connessione PostgreSQL
├── models.py         # Modello ProjectSizing
├── templates/
│   ├── login.html    # Pagina di accesso
│   ├── index.html    # App principale (questionario + storico)
│   └── admin.html    # Pannello gestione utenti
├── Dockerfile
├── railway.toml
└── requirements.txt
```
