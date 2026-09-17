from datetime import datetime
from flask_login import UserMixin
from app.extensions import db


class User(db.Model, UserMixin):
    __tablename__ = "user"

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    displayname = db.Column(db.String(80), nullable=True)
    password = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default="user")
    status = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return f"<User {self.username}>"


class AppSetting(db.Model):
    """Site branding - a single row (id=1), fetched/created on demand by
    app.services.app_settings.get_app_settings(). Logo/favicon columns hold
    just the filename under static/uploads/branding/ - None means "use the
    built-in default logo.png" (see that module for the fallback logic)."""
    __tablename__ = "app_setting"

    id = db.Column(db.Integer, primary_key=True)
    app_name = db.Column(db.String(100), nullable=False, default="Design Calculations")
    login_logo = db.Column(db.String(255))
    sidebar_logo = db.Column(db.String(255))
    favicon = db.Column(db.String(255))
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


class OTP(db.Model):
    __tablename__ = "otp"

    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(120), unique=True, nullable=False)
    code = db.Column(db.String(6), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return f"<OTP {self.email}>"


# ---------------------------------------------------------------------------
# GAD Automation - Globe valve (Series 10/11) lookup tables.
#
# These mirror the SQL Server tables the old WinForms app (GAGlobeTable,
# GAGlobeHookUp, GACrossSecGlobe, GADimvalveGlobe, GADimActGlobe) queried to
# turn a valve configuration into drawing/schematic block numbers and
# dimension values. They start empty here - use seed_gad_globe.py to bulk
# load rows exported as CSV from the old database.
# ---------------------------------------------------------------------------

class GAGlobeTable(db.Model):
    """Series/config -> main GA drawing block number (Sheet1).

    valve_size, rating, actuator_type, actuator_size and traval_stop can each
    hold a comma-separated list of values in one cell (e.g. "1,1.5" or
    "Piston with Spring, Piston without Spring"). Matching strips spaces and
    looks for an exact comma-delimited token - see
    app/services/gad_globe_lookup.py.
    """
    __tablename__ = "ga_globe_table"

    id = db.Column(db.Integer, primary_key=True)
    valve_series = db.Column(db.String(10), nullable=False)
    body_style = db.Column(db.String(50), nullable=False)
    valve_size = db.Column(db.String(100), nullable=False)
    rating = db.Column(db.String(100), nullable=False)
    end_connection = db.Column(db.String(50), nullable=False)
    bonnet_type = db.Column(db.String(50), nullable=False)
    flow_direction = db.Column(db.String(50), nullable=False)
    actuator_series = db.Column(db.String(50), nullable=False)
    actuator_type = db.Column(db.String(200), nullable=False)
    actuator_size = db.Column(db.String(100), nullable=False)
    traval_stop = db.Column(db.String(100), nullable=False)
    hw = db.Column(db.String(20), nullable=False)
    drawing_no = db.Column(db.String(50), nullable=False)


class GAGlobeHookUp(db.Model):
    """Actuator/instrumentation config -> hookup schematic block number."""
    __tablename__ = "ga_globe_hookup"

    id = db.Column(db.Integer, primary_key=True)
    actuator = db.Column(db.String(50), nullable=False)
    spring = db.Column(db.String(50), nullable=False)
    air_fail_action = db.Column(db.String(50), nullable=False)
    positioner = db.Column(db.String(50), nullable=False)
    hand_auto = db.Column(db.String(50), nullable=False)
    limit_switch = db.Column(db.String(50), nullable=False)
    position_trans = db.Column(db.String(50), nullable=False)
    volume_booster = db.Column(db.String(50), nullable=False)
    spool_valve = db.Column(db.String(50), nullable=False)
    lock_valve = db.Column(db.String(50), nullable=False)
    solenoid_valve = db.Column(db.String(50), nullable=False)
    schematic_no = db.Column(db.String(50), nullable=False)


class GACrossSecGlobe(db.Model):
    """Body/trim config -> cross-section drawing block number (Sheet2)."""
    __tablename__ = "ga_crosssec_globe"

    id = db.Column(db.Integer, primary_key=True)
    body_style = db.Column(db.String(50), nullable=False)
    size = db.Column(db.String(50), nullable=False)
    bonnet_type = db.Column(db.String(50), nullable=False)
    trim_type = db.Column(db.String(50), nullable=False)
    balancing = db.Column(db.String(50), nullable=False)
    bal_seal_type = db.Column(db.String(50), nullable=False)
    seat_type = db.Column(db.String(50), nullable=False)
    packing_type = db.Column(db.String(50), nullable=False)
    flow_direction = db.Column(db.String(50), nullable=False)
    drawing_no = db.Column(db.String(50), nullable=False)


class GADimValveGlobe(db.Model):
    """Body dimensions (A, B, C, AR) + valve weight, keyed off body config."""
    __tablename__ = "ga_dim_valve_globe"

    id = db.Column(db.Integer, primary_key=True)
    body_style = db.Column(db.String(50), nullable=False)
    end_connection = db.Column(db.String(50), nullable=False)
    bonnet_type = db.Column(db.String(50), nullable=False)
    end_finish = db.Column(db.String(50), nullable=False)
    series = db.Column(db.String(10), nullable=False)
    size = db.Column(db.String(50), nullable=False)
    rating = db.Column(db.String(50), nullable=False)
    stem_dia = db.Column(db.String(50), nullable=False)
    a = db.Column(db.String(50))
    b = db.Column(db.String(50))
    c = db.Column(db.String(50))
    ar = db.Column(db.String(50))
    weight = db.Column(db.Numeric(10, 2), default=0)


class GADimActGlobe(db.Model):
    """Actuator dimensions (D, E, F) + actuator weight, keyed off actuator config."""
    __tablename__ = "ga_dim_act_globe"

    id = db.Column(db.Integer, primary_key=True)
    actuator_type = db.Column(db.String(50), nullable=False)
    actuator_size = db.Column(db.String(50), nullable=False)
    hand_wheel = db.Column(db.String(20), nullable=False)
    travel_stop = db.Column(db.String(50), nullable=False)
    d = db.Column(db.String(50))
    e = db.Column(db.String(50))
    f = db.Column(db.String(50))
    weight = db.Column(db.Numeric(10, 2), default=0)


class GadGenerationJob(db.Model):
    """A queued "Generate GAD" request.

    Decouples the web app (upload, dropdown validation, GAD Masters - none
    of which need SolidWorks) from the actual SolidWorks automation step,
    which can only run on whichever machine physically has SolidWorks
    installed. The web app just writes a 'pending' row here and polls for
    it to flip to 'done'/'error'; a separate worker process (worker.py),
    running on a SolidWorks machine, polls the same table and does the real
    work - see app/services/solidworks_automation.py for that part, which
    is unchanged either way.
    """
    __tablename__ = "gad_generation_job"

    STATUS_PENDING = "pending"
    STATUS_PROCESSING = "processing"
    STATUS_DONE = "done"
    STATUS_ERROR = "error"

    id = db.Column(db.Integer, primary_key=True)
    status = db.Column(db.String(20), nullable=False, default=STATUS_PENDING)
    row_data = db.Column(db.JSON, nullable=False)
    download_name = db.Column(db.String(120))
    pdf_data = db.Column(db.LargeBinary)
    error_message = db.Column(db.Text)
    requested_by_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    worker_hostname = db.Column(db.String(120))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    started_at = db.Column(db.DateTime)
    finished_at = db.Column(db.DateTime)
