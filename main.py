import os
from fastapi import FastAPI, Depends, Request, HTTPException
from fastapi.responses import JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from typing import Optional

from database import init_db, get_db
from sqlalchemy.orm import Session
from models import ProjectSizing
from auth import (
    load_users, save_user, delete_user, delete_user_sessions,
    hash_password, verify_password,
    create_session, get_session, delete_session,
    seed_users, get_current_user, require_admin,
)

app = FastAPI(title="Archiva Project Sizing", docs_url=None, redoc_url=None)

os.makedirs("static", exist_ok=True)
app.mount("/static", StaticFiles(directory="static", html=True), name="static")


@app.on_event("startup")
def on_startup():
    init_db()
    seed_users()


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


# ── HEALTH ────────────────────────────────────────────────────────────────────

@app.get("/health")
def health():
    return {"status": "ok"}
