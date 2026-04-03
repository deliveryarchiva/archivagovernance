"""
Auth module — stesso pattern di archiva-v10 (AI Minute Riunioni).
- Utenti in JSON file (Railway Volume o /data locale)
- Password SHA-256 (no bcrypt, no JWT)
- Sessioni in-memory con UUID casuali
- Nessuna SECRET_KEY necessaria
"""
import hashlib
import json
import os
import uuid
from pathlib import Path
from typing import Optional

from fastapi import Header, HTTPException, status

# ── Storage ──────────────────────────────────────────────────────────────────
DATA_DIR   = Path(os.getenv("RAILWAY_VOLUME_MOUNT_PATH", "data"))
USERS_FILE = DATA_DIR / "users.json"
DATA_DIR.mkdir(parents=True, exist_ok=True)

# Sessioni in-memory: token → {username, nome, ruolo, role}
sessions: dict[str, dict] = {}


# ── Hashing ───────────────────────────────────────────────────────────────────
def hash_password(plain: str) -> str:
    return hashlib.sha256(plain.encode()).hexdigest()

def verify_password(plain: str, hashed: str) -> bool:
    return hash_password(plain) == hashed


# ── Users JSON ────────────────────────────────────────────────────────────────
def load_users() -> dict:
    try:
        if USERS_FILE.exists():
            return json.loads(USERS_FILE.read_text("utf-8"))
    except Exception as e:
        print(f"[auth] users load error: {e}")
    return {}

def save_user(user: dict):
    users = load_users()
    users[user["username"].lower()] = user
    tmp = str(USERS_FILE) + ".tmp"
    Path(tmp).write_text(json.dumps(users, indent=2, ensure_ascii=False), "utf-8")
    Path(tmp).rename(USERS_FILE)

def delete_user(username: str):
    users = load_users()
    users.pop(username.lower(), None)
    tmp = str(USERS_FILE) + ".tmp"
    Path(tmp).write_text(json.dumps(users, indent=2, ensure_ascii=False), "utf-8")
    Path(tmp).rename(USERS_FILE)


# ── Seed utenti iniziali ──────────────────────────────────────────────────────
def seed_users():
    users = load_users()
    if users:
        print(f"[auth] {len(users)} utenti già presenti")
        return

    default_pwd = hash_password(os.getenv("DEFAULT_PASSWORD", "archiva2026"))
    initial = [
        {"nome": "Marco Pastore",       "username": "marco.pastore",       "ruolo": "admin", "role": "Head of Project Delivery"},
        {"nome": "Paolo Gandini",        "username": "paolo.gandini",        "ruolo": "admin", "role": "Delivery & Customer Service Director"},
        {"nome": "Chiara Pettenuzzo",    "username": "chiara.pettenuzzo",    "ruolo": "admin", "role": "Service Delivery Manager"},
    ]
    for u in initial:
        save_user({**u, "password": default_pwd})
        print(f"[auth]   ok {u['username']} ({u['ruolo']})")
    print(f"[auth] {len(initial)} utenti creati — password default: {os.getenv('DEFAULT_PASSWORD','archiva2026')}")


# ── Sessioni ──────────────────────────────────────────────────────────────────
def create_session(user: dict) -> str:
    token = str(uuid.uuid4())
    sessions[token] = {
        "username": user["username"],
        "nome":     user["nome"],
        "ruolo":    user["ruolo"],
        "role":     user.get("role", ""),
    }
    return token

def get_session(token: str) -> Optional[dict]:
    return sessions.get(token)

def delete_session(token: str):
    sessions.pop(token, None)

def delete_user_sessions(username: str):
    to_del = [t for t, s in sessions.items() if s["username"].lower() == username.lower()]
    for t in to_del:
        sessions.pop(t, None)


# ── FastAPI dependency ────────────────────────────────────────────────────────
def get_current_user(authorization: Optional[str] = Header(default=None)) -> dict:
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Non autenticato")
    token = authorization.removeprefix("Bearer ").strip()
    user = get_session(token)
    if not user:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Sessione scaduta")
    return user

def require_admin(user: dict = None) -> dict:
    if user and user.get("ruolo") != "admin":
        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="Accesso riservato agli amministratori")
    return user
