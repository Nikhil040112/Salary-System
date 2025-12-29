import os
from fastapi import Request
from fastapi.responses import RedirectResponse
from argon2 import PasswordHasher
from itsdangerous import URLSafeSerializer, BadSignature

ph = PasswordHasher()

ADMIN_USERNAME = os.getenv("ADMIN_USERNAME")
ADMIN_PASSWORD_HASH = os.getenv("ADMIN_PASSWORD_HASH")
SECRET_KEY = os.getenv("SECRET_KEY")

serializer = URLSafeSerializer(SECRET_KEY, salt="login-session")


def verify_login(username: str, password: str) -> bool:
    if not ADMIN_USERNAME or not ADMIN_PASSWORD_HASH:
        return False
    if username != ADMIN_USERNAME:
        return False
    try:
        ph.verify(ADMIN_PASSWORD_HASH, password)
        return True
    except Exception:
        return False
    
def login_required(request: Request):
    session = request.cookies.get("session")

    if not session:
        return RedirectResponse("/login", status_code=303)
    
    try:
        data = serializer.loads(session)
        if data.get("user") != ADMIN_USERNAME:
            raise BadSignature()
    except BadSignature:
        response = RedirectResponse("/login", status_code=303)
        response.delete_cookie("session")
        return response