import random
from datetime import datetime, timedelta

from flask import render_template, redirect, url_for, flash, request, session
from flask_login import login_required, current_user, login_user, logout_user
from sqlalchemy import func, or_
from werkzeug.security import generate_password_hash, check_password_hash
from flask_mail import Message

from app.blueprints.auth import auth_bp
from app.extensions import db, mail
from app.models import User, OTP

OTP_VALID_MINUTES = 10


@auth_bp.route("/")
def index():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard.dashboard"))
    return redirect(url_for("auth.login"))


@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard.dashboard"))

    next_url = request.form.get("next", "").strip() or request.args.get("next", "").strip()

    if request.method == "POST":
        username_input = request.form.get("username", "").strip()
        password_input = request.form.get("password", "")

        user = User.query.filter(
            or_(
                func.lower(User.email) == username_input.lower(),
                func.lower(User.username) == username_input.lower(),
            )
        ).first()

        if not user or not check_password_hash(user.password, password_input):
            flash("Invalid username or password.", "danger")
            return render_template("login.html", next_url=next_url)

        if not user.status:
            flash("Your account is disabled. Please contact the administrator.", "danger")
            return render_template("login.html", next_url=next_url)

        remember = request.form.get("remember") == "remember"
        login_user(user, remember=remember)

        from urllib.parse import urlparse
        parsed = urlparse(next_url)
        if next_url and not parsed.netloc and not parsed.scheme:
            return redirect(next_url)

        return redirect(url_for("dashboard.dashboard"))

    return render_template("login.html", next_url=next_url)


@auth_bp.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("auth.login"))


@auth_bp.route("/forgot-password", methods=["GET", "POST"])
def forgotpassword():
    if request.method == "POST":
        email = request.form.get("email", "").strip()
        user = User.query.filter(func.lower(User.email) == email.lower()).first()

        if not user:
            flash("Email is not recognized.", "danger")
            return render_template("forgotpassword.html")

        if send_otp(email):
            session["reset-email"] = email
            return redirect(url_for("auth.resetpassword"))

        flash("Could not send the OTP email. Please try again.", "danger")

    return render_template("forgotpassword.html")


def send_otp(email: str) -> bool:
    code = f"{random.randint(0, 999999):06d}"

    otp = OTP.query.filter_by(email=email).first()
    if not otp:
        otp = OTP(email=email, code=code, created_at=datetime.utcnow())
        db.session.add(otp)
    else:
        otp.code = code
        otp.created_at = datetime.utcnow()
    db.session.commit()

    try:
        message = Message(
            subject="Your password reset OTP",
            recipients=[email],
            body=f"Your OTP to reset your password is: {code}\nIt is valid for {OTP_VALID_MINUTES} minutes.",
        )
        mail.send(message)
        return True
    except Exception as e:
        from flask import current_app
        current_app.logger.warning("Failed to send OTP email: %s", e)
        return False


@auth_bp.route("/reset-password", methods=["GET", "POST"])
def resetpassword():
    email = session.get("reset-email")
    if not email:
        return redirect(url_for("auth.forgotpassword"))

    if request.method == "POST":
        otp_input = request.form.get("otp", "").strip()
        new_password = request.form.get("password", "")

        otp = OTP.query.filter_by(email=email).first()
        expired = otp and datetime.utcnow() - otp.created_at > timedelta(minutes=OTP_VALID_MINUTES)

        if not otp or otp.code != otp_input or expired:
            flash("Incorrect or expired OTP.", "danger")
            return render_template("resetpassword.html", email=email)

        user = User.query.filter(func.lower(User.email) == email.lower()).first()
        user.password = generate_password_hash(new_password, method="pbkdf2:sha256", salt_length=8)
        db.session.delete(otp)
        db.session.commit()

        session.pop("reset-email", None)
        flash("Password reset successfully. Please sign in.", "success")
        return redirect(url_for("auth.login"))

    return render_template("resetpassword.html", email=email)
