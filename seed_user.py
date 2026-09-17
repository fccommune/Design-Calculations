"""
Create (or update) a login user.

Usage:
    python seed_user.py <username> <email> <password> [role]

Example:
    python seed_user.py admin admin@example.com Passw0rd! admin
"""
import sys
from dotenv import load_dotenv

load_dotenv()

from app import create_app
from app.extensions import db
from app.models import User
from werkzeug.security import generate_password_hash

if len(sys.argv) not in (4, 5):
    print(__doc__)
    sys.exit(1)

username, email, password = sys.argv[1], sys.argv[2], sys.argv[3]
role = sys.argv[4] if len(sys.argv) == 5 else "user"

app = create_app("development")
with app.app_context():
    user = User.query.filter_by(username=username).first()
    if user:
        user.password = generate_password_hash(password, method="pbkdf2:sha256", salt_length=8)
        user.email = email
        user.role = role
        print(f"Updated existing user '{username}'.")
    else:
        user = User(
            username=username,
            email=email,
            displayname=username,
            password=generate_password_hash(password, method="pbkdf2:sha256", salt_length=8),
            role=role,
            status=True,
        )
        db.session.add(user)
        print(f"Created new user '{username}'.")
    db.session.commit()
