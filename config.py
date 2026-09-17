import os

class BaseConfig:
    APP_NAME = "Flask Auth Template"

    SECRET_KEY = os.getenv("SECRET_KEY", "dev-only-insecure-key")
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = "Lax"
    PERMANENT_SESSION_LIFETIME = 28800

    SQLALCHEMY_TRACK_MODIFICATIONS = False

    MAIL_SERVER = os.getenv("MAIL_SERVER", "smtp.office365.com")
    MAIL_PORT = int(os.getenv("MAIL_PORT", 587))
    MAIL_USE_TLS = True
    MAIL_USE_SSL = False
    MAIL_USERNAME = os.getenv("MAIL_USERNAME", "")
    MAIL_PASSWORD = os.getenv("MAIL_PASSWORD", "")
    MAIL_DEFAULT_SENDER = os.getenv("MAIL_DEFAULT_SENDER", "")

    # GAD Automation (Globe valve) - SolidWorks template/block file locations
    # and where exported PDFs get written. These were hardcoded E:\... paths
    # in the old WinForms app; here they're configurable per machine.
    GAD_TEMPLATE_PATH = os.getenv("GAD_TEMPLATE_PATH", "")
    GAD_BLOCK_DIR = os.getenv("GAD_BLOCK_DIR", "")
    GAD_HOOKUP_BLOCK_DIR = os.getenv("GAD_HOOKUP_BLOCK_DIR", "")
    GAD_CROSSSEC_BLOCK_DIR = os.getenv("GAD_CROSSSEC_BLOCK_DIR", "")
    GAD_NAMEPLATE_BLOCK_PATH = os.getenv("GAD_NAMEPLATE_BLOCK_PATH", "")
    GAD_OUTPUT_DIR = os.getenv("GAD_OUTPUT_DIR", "")


class DevelopmentConfig(BaseConfig):
    DEBUG = True
    SQLALCHEMY_DATABASE_URI = os.getenv(
        "DATABASE_URL", "postgresql://postgres:password@localhost:5432/flaskauthtemplate"
    )


class ProductionConfig(BaseConfig):
    DEBUG = False
    TESTING = False

    SECRET_KEY = os.getenv("SECRET_KEY", "")
    SQLALCHEMY_DATABASE_URI = os.getenv("DATABASE_URL", "")

    SESSION_COOKIE_SECURE = True
    SESSION_COOKIE_SAMESITE = "Strict"


config_map = {
    "development": DevelopmentConfig,
    "production": ProductionConfig,
    "default": DevelopmentConfig,
}
