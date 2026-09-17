import os
import logging
from flask import Flask
from config import config_map


def create_app(config_name="default"):
    app = Flask(
        __name__,
        template_folder=os.path.join(os.path.dirname(os.path.dirname(__file__)), "templates"),
        static_folder=os.path.join(os.path.dirname(os.path.dirname(__file__)), "static"),
    )
    app.config.from_object(config_map[config_name])

    if config_name == "production":
        for key in ("SECRET_KEY", "SQLALCHEMY_DATABASE_URI"):
            if not app.config.get(key):
                raise RuntimeError(
                    f"Required config '{key}' is not set. "
                    "Copy .env.example to .env and fill in all values."
                )

    log_level = logging.DEBUG if app.debug else logging.INFO
    logging.basicConfig(level=log_level, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    app.logger.setLevel(log_level)

    from app.extensions import db, login_manager, mail, migrate
    db.init_app(app)
    login_manager.init_app(app)
    mail.init_app(app)
    migrate.init_app(app, db)

    login_manager.login_view = "auth.login"
    login_manager.login_message_category = "info"
    login_manager.session_protection = "strong"

    @login_manager.user_loader
    def load_user(user_id):
        from app.models import User
        return User.query.get(int(user_id))

    @app.context_processor
    def inject_app_settings():
        from app.services.app_settings import get_or_create_settings, settings_context
        return {"app_settings": settings_context(get_or_create_settings())}

    from app.blueprints.auth import auth_bp
    from app.blueprints.dashboard import dashboard_bp
    app.register_blueprint(auth_bp)
    app.register_blueprint(dashboard_bp)

    @app.after_request
    def set_security_headers(response):
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "SAMEORIGIN"
        response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
        return response

    return app
