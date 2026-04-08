import os
from fastapi import FastAPI, Depends, Request, HTTPException
from fastapi.responses import JSONResponse, FileResponse, Response
from fastapi.staticfiles import StaticFiles
from typing import Optional

from database import init_db, get_db
from sqlalchemy.orm import Session
from models import ProjectSizing, GovernanceItem
from auth import (
    load_users, save_user, delete_user, delete_user_sessions,
    hash_password, verify_password,
    create_session, get_session, delete_session,
    seed_users, get_current_user, require_admin,
)
from export_docx import generate_sizing_docx

app = FastAPI(title="Archiva Project Sizing", docs_url=None, redoc_url=None)

os.makedirs("static", exist_ok=True)
app.mount("/static", StaticFiles(directory="static", html=True), name="static")


@app.on_event("startup")
def on_startup():
    init_db()
    seed_users()
    seed_governance_items()


# ── PAGES ─────────────────────────────────────────────────────────────────────

@app.get("/")
def root():
    return FileResponse("templates/index.html")

@app.get("/login")
def login_page():
    return FileResponse("templates/login.html")

@app.get("/admin")
def admin_page():
    return FileResponse("templates/admin.html")


# ── AUTH ──────────────────────────────────────────────────────────────────────

@app.post("/api/auth/login")
async def login(request: Request):
    body = await request.json()
    username = (body.get("username") or "").strip().lower()
    password = body.get("password") or ""
    if not username or not password:
        raise HTTPException(400, "Credenziali mancanti")
    users = load_users()
    user = users.get(username)
    if not user or not verify_password(password, user["password"]):
        raise HTTPException(401, "Credenziali non valide")
    token = create_session(user)
    return {"ok": True, "token": token, "user": {
        "username": user["username"], "nome": user["nome"],
        "ruolo": user["ruolo"], "role": user.get("role", "")
    }}

@app.post("/api/auth/logout")
def logout(request: Request):
    auth = request.headers.get("authorization", "")
    if auth.startswith("Bearer "):
        delete_session(auth.removeprefix("Bearer ").strip())
    return {"ok": True}

@app.get("/api/auth/me")
def me(user=Depends(get_current_user)):
    return {"ok": True, "user": user}

@app.post("/api/auth/change-password")
async def change_password(request: Request, user=Depends(get_current_user)):
    body = await request.json()
    current = body.get("currentPassword", "")
    new_pwd = body.get("newPassword", "")
    if not current or not new_pwd:
        raise HTTPException(400, "Compila tutti i campi")
    if len(new_pwd) < 6:
        raise HTTPException(400, "La nuova password deve avere almeno 6 caratteri")
    users = load_users()
    db_user = users.get(user["username"].lower())
    if not db_user or not verify_password(current, db_user["password"]):
        raise HTTPException(401, "Password attuale non corretta")
    db_user["password"] = hash_password(new_pwd)
    save_user(db_user)
    return {"ok": True}


# ── ADMIN: utenti ─────────────────────────────────────────────────────────────

@app.get("/api/admin/users")
def admin_list_users(user=Depends(get_current_user)):
    require_admin(user)
    users = load_users()
    return {"ok": True, "data": [
        {"username": u["username"], "nome": u["nome"], "ruolo": u["ruolo"], "role": u.get("role", "")}
        for u in users.values()
    ]}

@app.post("/api/admin/users")
async def admin_create_user(request: Request, user=Depends(get_current_user)):
    require_admin(user)
    body = await request.json()
    username = (body.get("username") or "").strip().lower()
    password = body.get("password") or ""
    nome     = (body.get("nome") or "").strip()
    if not username or not password or not nome:
        raise HTTPException(400, "Campi obbligatori mancanti")
    users = load_users()
    if username in users:
        raise HTTPException(409, "Username già esistente")
    save_user({
        "username": username, "nome": nome,
        "ruolo": body.get("ruolo", "user"),
        "role":  body.get("role", ""),
        "password": hash_password(password),
    })
    return {"ok": True}

@app.put("/api/admin/users/{username}")
async def admin_update_user(username: str, request: Request, user=Depends(get_current_user)):
    require_admin(user)
    users = load_users()
    db_user = users.get(username.lower())
    if not db_user:
        raise HTTPException(404, "Utente non trovato")
    body = await request.json()
    if "nome"  in body: db_user["nome"]  = body["nome"]
    if "ruolo" in body: db_user["ruolo"] = body["ruolo"]
    if "role"  in body: db_user["role"]  = body["role"]
    if "password" in body and body["password"]:
        db_user["password"] = hash_password(body["password"])
    save_user(db_user)
    return {"ok": True}

@app.delete("/api/admin/users/{username}")
def admin_delete_user(username: str, user=Depends(get_current_user)):
    require_admin(user)
    if username.lower() == user["username"].lower():
        raise HTTPException(400, "Non puoi eliminare il tuo account")
    delete_user_sessions(username)
    delete_user(username)
    return {"ok": True}


# ── SIZINGS ───────────────────────────────────────────────────────────────────

TAGLIE = [
    {"id": "xs",  "label": "XS",  "min": 0,  "max": 2},
    {"id": "s",   "label": "S",   "min": 3,  "max": 5},
    {"id": "m",   "label": "M",   "min": 6,  "max": 9},
    {"id": "l",   "label": "L",   "min": 10, "max": 14},
    {"id": "xl",  "label": "XL",  "min": 15, "max": 19},
    {"id": "xxl", "label": "XXL", "min": 20, "max": 24.5},
]

def compute_taglia(score: float) -> str:
    for t in TAGLIE:
        if t["min"] <= score <= t["max"]:
            return t["id"]
    return "xxl"


@app.post("/api/sizings")
async def create_sizing(request: Request, user=Depends(get_current_user), db: Session = Depends(get_db)):
    body = await request.json()
    cliente = (body.get("cliente") or "").strip()
    titolo  = (body.get("titolo") or "").strip()
    answers = body.get("answers") or {}
    if not cliente or not titolo:
        raise HTTPException(400, "Cliente e titolo sono obbligatori")
    if len(answers) < 11:
        raise HTTPException(400, "Il questionario deve essere completato")

    score = sum(float(v) for v in answers.values())
    taglia = compute_taglia(score)

    sizing = ProjectSizing(
        crm_number=     (body.get("crm_number") or "").strip() or None,
        odl=            (body.get("odl") or "").strip() or None,
        cliente=        cliente,
        titolo=         titolo,
        answers=        answers,
        score=          score,
        taglia=         taglia,
        note=           (body.get("note") or "").strip() or None,
        created_by=     user["username"],
        created_by_nome=user["nome"],
    )
    db.add(sizing)
    db.commit()
    db.refresh(sizing)
    return {"ok": True, "id": sizing.id, "taglia": taglia, "score": score}


@app.get("/api/sizings")
def list_sizings(
    search: Optional[str] = None,
    taglia: Optional[str] = None,
    limit: int = 100,
    offset: int = 0,
    user=Depends(get_current_user),
    db: Session = Depends(get_db)
):
    q = db.query(ProjectSizing)
    if search:
        like = f"%{search}%"
        q = q.filter(
            ProjectSizing.cliente.ilike(like) |
            ProjectSizing.titolo.ilike(like) |
            ProjectSizing.crm_number.ilike(like) |
            ProjectSizing.odl.ilike(like)
        )
    if taglia:
        q = q.filter(ProjectSizing.taglia == taglia)
    total = q.count()
    rows = q.order_by(ProjectSizing.created_at.desc()).offset(offset).limit(limit).all()
    return {
        "ok": True,
        "total": total,
        "data": [_serialize(r) for r in rows],
    }


@app.get("/api/sizings/{sizing_id}")
def get_sizing(sizing_id: int, user=Depends(get_current_user), db: Session = Depends(get_db)):
    s = db.query(ProjectSizing).filter(ProjectSizing.id == sizing_id).first()
    if not s:
        raise HTTPException(404, "Sizing non trovato")
    return {"ok": True, "data": _serialize(s)}


@app.delete("/api/sizings/{sizing_id}")
def delete_sizing(sizing_id: int, user=Depends(get_current_user), db: Session = Depends(get_db)):
    require_admin(user)
    s = db.query(ProjectSizing).filter(ProjectSizing.id == sizing_id).first()
    if not s:
        raise HTTPException(404, "Sizing non trovato")
    db.delete(s)
    db.commit()
    return {"ok": True}


def _serialize(s: ProjectSizing) -> dict:
    return {
        "id":               s.id,
        "crm_number":       s.crm_number,
        "odl":              s.odl,
        "cliente":          s.cliente,
        "titolo":           s.titolo,
        "answers":          s.answers,
        "score":            s.score,
        "taglia":           s.taglia,
        "note":             s.note,
        "created_at":       s.created_at.isoformat() if s.created_at else None,
        "created_by":       s.created_by,
        "created_by_nome":  s.created_by_nome,
    }



# ── GOVERNANCE ITEMS — seed & CRUD ────────────────────────────────────────────

GOV_DEFAULTS = [
    ("verbali_action_log",       "Verbali e action item log",       "D", "Esecuzione",             "xs", "Ad ogni riunione",
     "Registro delle decisioni prese e delle azioni assegnate durante le riunioni di progetto, con responsabile e scadenza. Garantisce tracciabilità e accountability su tutto ciò che viene discusso e concordato. Prodotto dopo ogni riunione significativa per l'intera durata del progetto."),
    ("piano_progetto",           "Piano di progetto (baseline)",    "D", "Pianificazione",          "s",  "Una tantum — aggiornato a ogni variante approvata",
     "Documento che formalizza scope, milestones, timeline, risorse e dipendenze del progetto nella versione approvata. Costituisce la baseline di riferimento per il monitoraggio dell'avanzamento. Viene aggiornato formalmente a ogni variante approvata."),
    ("kickoff",                  "Kick-off",                        "A", "Avvio",                   "s",  "Una tantum",
     "Riunione iniziale con il team di progetto e i referenti cliente per allineare obiettivi, scope, ruoli e modalità operative. Segna il formale avvio del progetto e la condivisione del piano di lavoro. È il momento in cui si stabilisce il tono della collaborazione."),
    ("milestone_plan",           "Milestone plan",                  "D", "Pianificazione",          "m",  "Una tantum — condiviso con Steering Committee",
     "Piano sintetico dei principali punti di controllo del progetto con le date attese di completamento. Fornisce una visione d'insieme dell'avanzamento, pensata per essere condivisa con lo Steering Committee e il management. Aggiornato a ogni revisione della pianificazione."),
    ("raci",                     "RACI",                            "D", "Avvio / Pianificazione",  "m",  "Una tantum — rivisto a ogni cambio scope",
     "Matrice che definisce chi è Responsabile, Approvatore, Consultato e Informato per ciascuna attività di progetto. Elimina ambiguità sui ruoli e riduce i rischi di disallineamento operativo. Viene aggiornata a ogni variazione significativa dello scope o del team."),
    ("risk_register",            "Risk Register",                   "D", "Pianificazione",          "m",  "Continuo (documento vivo)",
     "Documento vivo che censisce i rischi identificati, la loro probabilità, impatto e le azioni di mitigazione associate. Aggiornato continuamente durante l'esecuzione del progetto. Supporta le decisioni proattive prima che i rischi si trasformino in problemi concreti."),
    ("issue_log",                "Issue Log",                       "D", "Esecuzione",              "m",  "Continuo (documento vivo)",
     "Registro continuo dei problemi aperti che impattano l'esecuzione del progetto, con stato, responsabile della risoluzione e data prevista di chiusura. Distinto dal Risk Register perché traccia problemi già materializzati, non rischi potenziali. Aggiornato in modo continuativo."),
    ("meeting_operativi",        "Meeting operativi strutturati",   "A", "Esecuzione",              "m",  "Settimanale",
     "Riunioni periodiche — tipicamente settimanali — con agenda predefinita per monitorare avanzamento, blocchi e azioni in corso. Garantiscono continuità nel coordinamento del team e visibilità costante sullo stato del progetto. Supportate da verbale e action item log."),
    ("organigramma",             "Organigramma di progetto",        "D", "Avvio",                   "m",  "Una tantum — aggiornato a ogni cambio ruolo",
     "Schema visivo delle figure coinvolte nel progetto, sia lato Archiva che lato cliente, con relativi ruoli e linee di riporto. Attivato in taglia M quando sono presenti vendor interni o esterni. Aggiornato a ogni cambio di ruolo rilevante."),
    ("stakeholder_register",     "Stakeholder Register",            "D", "Avvio",                   "l",  "Una tantum — aggiornato se cambiano gli attori",
     "Registro degli stakeholder chiave del progetto con informazioni su ruolo, livello di influenza, aspettative e strategia di coinvolgimento. Supporta la pianificazione delle comunicazioni e la gestione delle relazioni. Aggiornato quando cambiano gli attori coinvolti."),
    ("piano_comunicazione",      "Piano di comunicazione",          "D", "Avvio / Pianificazione",  "l",  "Una tantum",
     "Documento che definisce cosa comunicare, a chi, con quale frequenza e attraverso quale canale per tutta la durata del progetto. Coordina il flusso informativo tra team interno, vendor e cliente. Riduce il rischio di informazioni perse o disallineamenti su decisioni importanti."),
    ("change_management",        "Change management process",       "A", "Esecuzione",              "l",  "On demand",
     "Processo strutturato per valutare, approvare e gestire le richieste di variazione dello scope, del budget o della timeline del progetto. Evita che modifiche non governate alterino gli impegni contrattuali e la pianificazione. Attivato on demand ogni volta che emerge una richiesta di cambio."),
    ("escalation_process",       "Escalation process",              "A", "Esecuzione",              "l",  "On demand",
     "Procedura che definisce come e quando un problema viene portato a un livello decisionale superiore quando non può essere risolto a livello operativo. Garantisce che i blocchi critici vengano trattati tempestivamente senza restare bloccati nel team. Attivato on demand."),
    ("gate_review",              "Gate review / Phase gate",        "A", "Fine di ogni fase",       "l",  "Una tantum per fase",
     "Checkpoint formale al termine di ogni fase del progetto per validare i deliverable prodotti e autorizzare il passaggio alla fase successiva. Coinvolge il PM e i principali responsabili, e in certi casi il cliente. È un momento di controllo qualitativo, non solo di avanzamento temporale."),
    ("steering_committee",       "Steering Committee",              "A", "Esecuzione → Chiusura",   "l",  "Mensile / bimestrale",
     "Incontro periodico con il management di Archiva e del cliente per monitorare l'avanzamento strategico, validare decisioni rilevanti e gestire eventuali escalation. In taglia L è attivato solo se il cliente è tra i top 20. Prevede un'agenda formale e un verbale."),
    ("alignment_report",         "Alignment Report",                "D", "Esecuzione",              "l",  "Settimanale (Confluence)",
     "Report minimalista pubblicato settimanalmente su Confluence, destinato al commerciale di riferimento. Contiene percentuale di avanzamento del progetto, attività in corso, problematiche aperte e prossime attività. Il commerciale viene taggato ad ogni aggiornamento e riceve notifica automatica."),
    ("governance_effort_report", "Governance Effort Report",        "D", "Esecuzione → Chiusura",   "l",  "Mensile / a fine progetto",
     "Report interno che rendiconta le ore e l'effort spesi in attività di governance nel periodo. Serve a quantificare il costo organizzativo della governance stessa e ad alimentare le analisi di efficienza. Prodotto mensilmente e a consuntivo a fine progetto."),
    ("vendor_management_plan",   "Vendor management plan",          "D", "Pianificazione",          "xl", "Una tantum — aggiornato a nuovi contratti",
     "Piano che definisce le modalità di gestione dei fornitori coinvolti nel progetto: obiettivi contrattuali, SLA attesi, interfacce operative e processi di escalation. Garantisce una governance strutturata delle dipendenze esterne. Aggiornato a ogni nuovo contratto o variazione significativa."),
    ("cutover_plan",             "Cutover Plan",                    "D", "Pre go-live",             "xl", "Una tantum — aggiornato fino al go-live",
     "Documento che dettaglia la sequenza precisa di attività, responsabili e timing per il passaggio in produzione del sistema o del nuovo processo. Include finestre temporali, dipendenze tecniche e criteri di go/no-go. In taglia L è attivato solo in presenza di go-live tecnico."),
    ("rollback_plan",            "Rollback Plan",                   "D", "Pre go-live",             "xl", "Una tantum — validato prima del go-live",
     "Procedura di rientro da attivare in caso di blocco critico durante il cutover, con trigger di decisione chiari, responsabili e tempi massimi per tornare alla situazione precedente. Deve essere validato e condiviso prima del go-live, non durante. In taglia L attivato solo con go-live tecnico."),
    ("hypercare",                "Hypercare",                       "A", "Post go-live",            "xl", "Settimanale · 4–8 settimane",
     "Periodo di sorveglianza intensiva successivo al go-live, tipicamente 4-8 settimane, con monitoraggio settimanale dell'operatività e supporto prioritario in caso di anomalie. In taglie M e L è attivato solo se il go-live è tecnico e gli utenti impattati sono più di 5. Da XL in poi è sempre previsto."),
    ("handover_document",        "Handover document",               "D", "Chiusura",                "xl", "Una tantum — con sign-off del ricevente",
     "Documento formale che formalizza il passaggio del progetto in gestione operativa, includendo configurazioni, procedure, contatti e tutto ciò che serve al team di operations per gestire quanto consegnato. Prodotto in unica versione a fine progetto, con sign-off del ricevente."),
    ("handover_operations",      "Handover a operations",           "A", "Go-live / Chiusura",      "xl", "Una tantum",
     "Attività strutturata di trasferimento di conoscenza e responsabilità dal team di progetto al team di operations o supporto. Comprende sessioni di allineamento, consegna della documentazione e firma del verbale di accettazione. Segna il termine operativo del progetto."),
]


def seed_governance_items():
    from database import SessionLocal
    db = SessionLocal()
    try:
        existing = db.query(GovernanceItem).count()
        if existing > 0:
            return
        for row in GOV_DEFAULTS:
            key, name, tipo, fase, from_t, freq, desc = row
            db.add(GovernanceItem(key=key, name=name, tipo=tipo, fase=fase,
                                  from_taglia=from_t, frequenza=freq, descrizione=desc))
        db.commit()
        print(f"[gov] {len(GOV_DEFAULTS)} elementi di governance caricati")
    except Exception as e:
        db.rollback()
        print(f"[gov] errore seed: {e}")
    finally:
        db.close()


@app.get("/api/gov-items")
def list_gov_items(user=Depends(get_current_user), db: Session = Depends(get_db)):
    items = db.query(GovernanceItem).order_by(GovernanceItem.key).all()
    return {"ok": True, "data": [_gov_serialize(i) for i in items]}


@app.put("/api/gov-items/{key}")
async def update_gov_item(key: str, request: Request, user=Depends(get_current_user), db: Session = Depends(get_db)):
    require_admin(user)
    item = db.query(GovernanceItem).filter(GovernanceItem.key == key).first()
    if not item:
        raise HTTPException(404, "Elemento non trovato")
    body = await request.json()
    if "name"        in body: item.name        = body["name"].strip()
    if "frequenza"   in body: item.frequenza   = body["frequenza"].strip()
    if "descrizione" in body: item.descrizione = body["descrizione"].strip()
    item.updated_by = user["username"]
    db.commit()
    return {"ok": True}


def _gov_serialize(i: GovernanceItem) -> dict:
    return {
        "key":         i.key,
        "name":        i.name,
        "tipo":        i.tipo,
        "fase":        i.fase,
        "from_taglia": i.from_taglia,
        "frequenza":   i.frequenza,
        "descrizione": i.descrizione,
        "updated_at":  i.updated_at.isoformat() if i.updated_at else None,
        "updated_by":  i.updated_by,
    }


# ── HEALTH ────────────────────────────────────────────────────────────────────

@app.get("/health")
def health():
    return {"status": "ok"}


# ── EXPORT DOCX ───────────────────────────────────────────────────────────────

@app.get("/api/sizings/{sizing_id}/export")
def export_sizing_docx(sizing_id: int, user=Depends(get_current_user), db: Session = Depends(get_db)):
    s = db.query(ProjectSizing).filter(ProjectSizing.id == sizing_id).first()
    if not s:
        raise HTTPException(404, "Sizing non trovato")

    data = _serialize(s)
    gov_items = db.query(GovernanceItem).all()
    gov_map = {i.key: _gov_serialize(i) for i in gov_items}
    docx_bytes = generate_sizing_docx(data, gov_map)

    # Filename sicuro
    safe_odl    = (s.odl or "NODL").replace("/", "-").replace(" ", "_")
    safe_client = "".join(c for c in (s.cliente or "cliente") if c.isalnum() or c in (" ", "-", "_"))
    safe_client = safe_client.strip().replace(" ", "_")[:30]
    filename    = f"Allegato_Documenti_{safe_odl}_{safe_client}.docx"

    return Response(
        content=docx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
