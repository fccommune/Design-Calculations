"""
Globe valve (Series 10/11) lookups.

Direct port of the SQL Server queries in the old WinForms app's Form2.cs
(GenerateSheet1/2 and GetGlobeDimensionNoteValues), rewritten against the
Postgres tables in app/models.py (GAGlobeTable, GAGlobeHookUp,
GACrossSecGlobe, GADimValveGlobe, GADimActGlobe).

Each row from the uploaded Excel sheet is a plain dict keyed by the exact
column headers used in the old app (see EXCEL_COLUMNS below) - callers get
that dict from app.services.excel_import.parse_excel_rows().
"""
from sqlalchemy import func

from app.models import (
    GAGlobeTable,
    GAGlobeHookUp,
    GACrossSecGlobe,
    GASheet4Globe,
    GADimValveGlobe,
    GADimActGlobe,
)
from app.services.gad_masters_import import TABLE_LABELS


class GadLookupError(Exception):
    """Raised when a required lookup row can't be found - mirrors the old
    app's MessageBox.Show("No matching ... row ...") + early return.

    table_name (a MODELS_BY_TABLE key, e.g. "ga_crosssec_globe") and
    table_label (its GAD Masters display name, e.g. "Cross-Section") tell
    the caller *which* master table needs a row added/fixed - routes.py
    passes these through so the browser can point the user straight at
    /gad-masters/<table_name> instead of a bare "no match" message."""

    def __init__(self, message: str, table_name: str = None):
        super().__init__(message)
        self.table_name = table_name
        self.table_label = TABLE_LABELS.get(table_name)


def cell(row: dict, key: str) -> str:
    """row.Cells[key].Value?.ToString() ?? "" """
    value = row.get(key)
    return "" if value is None else str(value).strip()


def _no_spaces(value: str) -> str:
    return value.replace(" ", "")


def _token_match(column, value: str):
    """WHERE (',' + REPLACE(column, ' ', '') + ',') LIKE '%,' + value + ',%'

    Matches an exact comma-delimited token regardless of whether the source
    cell used "," or ", " as its separator - same trick the original SQL
    query used, translated to SQLAlchemy (still a bound parameter, not raw
    SQL interpolation).
    """
    wrapped = func.concat(",", func.replace(column, " ", ""), ",")
    return wrapped.like(f"%,{_no_spaces(value)},%")


def find_drawing_no(row: dict) -> str:
    """GAGlobeTable -> the main GA block for Sheet1."""
    series = cell(row, "Series")
    body_style = cell(row, "Body_Style")
    valve_size = cell(row, "Valve_Size ")  # trailing space - matches the Excel header
    rating = cell(row, "Rating")
    end_connection = cell(row, "End _Connection")
    bonnet_type = cell(row, "Bonnet _Type")
    flow_direction = cell(row, "Flow_Direction")
    actuator_series = cell(row, "Act Series")
    actuator_type = cell(row, "Act Type")
    actuator_size = cell(row, "Actuator Size")
    traval_stop = cell(row, "Traval stop")
    hw = cell(row, "H/W")

    match = (
        GAGlobeTable.query.filter(
            GAGlobeTable.valve_series == series,
            GAGlobeTable.body_style == body_style,
            _token_match(GAGlobeTable.valve_size, valve_size),
            _token_match(GAGlobeTable.rating, rating),
            GAGlobeTable.end_connection == end_connection,
            GAGlobeTable.bonnet_type == bonnet_type,
            GAGlobeTable.flow_direction == flow_direction,
            GAGlobeTable.actuator_series == actuator_series,
            _token_match(GAGlobeTable.actuator_type, actuator_type),
            _token_match(GAGlobeTable.actuator_size, actuator_size),
            _token_match(GAGlobeTable.traval_stop, traval_stop),
            GAGlobeTable.hw == hw,
        )
        .order_by(GAGlobeTable.id)
        .first()
    )
    if match is None:
        raise GadLookupError(
            "No matching row found in the Overall Assembly master table for this "
            "configuration - add a row there for this combination.",
            table_name="ga_globe_table",
        )
    return match.drawing_no


def find_hookup_no(row: dict) -> str:
    """GAGlobeHookUp -> the hookup schematic block for Sheet1."""
    actuator = cell(row, "Act Type")
    spring = cell(row, "Spring")
    air_fail_action = cell(row, "AirFailaction")
    positioner = cell(row, "Positioner")
    hand_auto = cell(row, "HandAuto")
    limit_switch = cell(row, "LimitSwitchType")
    position_trans = cell(row, "PositionTrans")
    volume_booster = cell(row, "VolumeBooster")
    spool_valve = cell(row, "SpoolValve")
    lock_valve = cell(row, "Airlockrelay")
    solenoid_valve = cell(row, "Solenoid Valve")

    match = (
        GAGlobeHookUp.query.filter_by(
            actuator=actuator,
            spring=spring,
            air_fail_action=air_fail_action,
            positioner=positioner,
            hand_auto=hand_auto,
            limit_switch=limit_switch,
            position_trans=position_trans,
            volume_booster=volume_booster,
            spool_valve=spool_valve,
            lock_valve=lock_valve,
            solenoid_valve=solenoid_valve,
        )
        .order_by(GAGlobeHookUp.id)
        .first()
    )
    if match is None:
        raise GadLookupError(
            "No matching row found in the Hook-Up Selection master table for this "
            "configuration - add a row there for this combination.",
            table_name="ga_globe_hookup",
        )
    return match.schematic_no


def find_crosssec_drawing_no(row: dict) -> str:
    """GACrossSecGlobe -> the cross-section block for Sheet2."""
    body_style = cell(row, "Body_Style")
    size = cell(row, "Valve_Size ")
    bonnet_type = cell(row, "Bonnet _Type")
    trim_type = cell(row, "Trim Type")
    balancing = cell(row, "Balancing")
    bal_seal_type = cell(row, "Bal Seal Type")
    seat_type = cell(row, "Seat Type")
    packing_type = cell(row, "Packing Type")
    flow_direction = cell(row, "Flow_Direction")

    match = (
        GACrossSecGlobe.query.filter_by(
            body_style=body_style,
            size=size,
            bonnet_type=bonnet_type,
            trim_type=trim_type,
            balancing=balancing,
            bal_seal_type=bal_seal_type,
            seat_type=seat_type,
            packing_type=packing_type,
            flow_direction=flow_direction,
        )
        .order_by(GACrossSecGlobe.id)
        .first()
    )
    if match is None:
        raise GadLookupError(
            "No matching row found in the Cross-Section master table for this "
            "configuration - add a row there for this combination.",
            table_name="ga_crosssec_globe",
        )
    return match.drawing_no


def find_sheet4_drawing_no(row: dict) -> str:
    """GASheet4Globe -> the Sheet4 reference drawing (CS0xx.SLDDRW).

    Same exact-match fields as find_crosssec_drawing_no() (GACrossSecGlobe,
    Sheet2's lookup) plus end_connection, and size/rating are matched the
    same comma-token way find_drawing_no() above matches GAGlobeTable's
    Overall Assembly columns (a master cell of "1,4" matches a BOM size of
    "1" or "4"; "ASME 150,ASME2500" matches a rating of "ASME 150") - see
    _token_match()."""
    body_style = cell(row, "Body_Style")
    end_connection = cell(row, "End _Connection")
    bonnet_type = cell(row, "Bonnet _Type")
    trim_type = cell(row, "Trim Type")
    balancing = cell(row, "Balancing")
    bal_seal_type = cell(row, "Bal Seal Type")
    seat_type = cell(row, "Seat Type")
    packing_type = cell(row, "Packing Type")
    flow_direction = cell(row, "Flow_Direction")
    size = cell(row, "Valve_Size ")
    rating = cell(row, "Rating")

    match = (
        GASheet4Globe.query.filter(
            GASheet4Globe.body_style == body_style,
            GASheet4Globe.end_connection == end_connection,
            GASheet4Globe.bonnet_type == bonnet_type,
            GASheet4Globe.trim_type == trim_type,
            GASheet4Globe.balancing == balancing,
            GASheet4Globe.bal_seal_type == bal_seal_type,
            GASheet4Globe.seat_type == seat_type,
            GASheet4Globe.packing_type == packing_type,
            GASheet4Globe.flow_direction == flow_direction,
            _token_match(GASheet4Globe.size, size),
            _token_match(GASheet4Globe.rating, rating),
        )
        .order_by(GASheet4Globe.id)
        .first()
    )
    if match is None:
        raise GadLookupError(
            "No matching row found in the Sheet4 Selection master table for this "
            "configuration - add a row there for this combination.",
            table_name="ga_sheet4_globe",
        )
    return match.drawing_no


def find_dimension_values(row: dict) -> dict:
    """GADimValveGlobe + GADimActGlobe -> A, B, C, AR, D, E, F, WEIGHT.

    WEIGHT is body weight + actuator weight summed, same as the original.
    """
    body_style = cell(row, "Body_Style")
    end_connection = cell(row, "End _Connection")
    bonnet_type = cell(row, "Bonnet _Type")
    end_finish = cell(row, "End_Finish")
    series = cell(row, "Series")
    size = cell(row, "Valve_Size ")
    rating = cell(row, "Rating")
    stem_dia = cell(row, "Stem_Dia")

    body_match = (
        GADimValveGlobe.query.filter_by(
            body_style=body_style,
            end_connection=end_connection,
            bonnet_type=bonnet_type,
            end_finish=end_finish,
            series=series,
            size=size,
            rating=rating,
            stem_dia=stem_dia,
        )
        .order_by(GADimValveGlobe.id)
        .first()
    )
    if body_match is None:
        raise GadLookupError(
            "No matching row found in the Valve Dimensions master table for this "
            "configuration - add a row there for this combination.",
            table_name="ga_dim_valve_globe",
        )

    actuator_type = cell(row, "Act Type")
    actuator_size = cell(row, "Actuator Size")
    hand_wheel = cell(row, "H/W")
    travel_stop = cell(row, "Traval stop")

    act_match = (
        GADimActGlobe.query.filter_by(
            actuator_type=actuator_type,
            actuator_size=actuator_size,
            hand_wheel=hand_wheel,
            travel_stop=travel_stop,
        )
        .order_by(GADimActGlobe.id)
        .first()
    )
    if act_match is None:
        raise GadLookupError(
            "No matching row found in the Actuator Dimensions master table for this "
            "configuration - add a row there for this combination.",
            table_name="ga_dim_act_globe",
        )

    body_weight = body_match.weight or 0
    act_weight = act_match.weight or 0

    return {
        "A": body_match.a or "",
        "B": body_match.b or "",
        "C": body_match.c or "",
        "AR": body_match.ar or "",
        "D": act_match.d or "",
        "E": act_match.e or "",
        "F": act_match.f or "",
        # "WEIGHT" is a guessed custom-property tag carried over from the
        # original app - confirm the real tag name on the SolidWorks
        # template and adjust here + in solidworks_automation.py if different.
        "WEIGHT": str(body_weight + act_weight),
    }


def resolve_globe_gad(row: dict) -> dict:
    """Does every database lookup GAD generation needs for one BOM row -
    the GA block, hookup schematic and Sheet2 reference drawing numbers,
    plus all ~50 title-block custom properties (dimension-table values and
    plain pass-through fields alike) - and returns them as one flat,
    plain-data dict with no further database dependency.

    This is the split point between the two halves of GAD generation:
    - This function needs the database (master lookup tables) but not
      SolidWorks - runs on the web server.
    - app.services.solidworks_automation.generate_from_resolved(...) needs
      SolidWorks but not the database - runs locally on whichever machine
      actually has SolidWorks, using only what this function already
      resolved (see GadGenerate.exe / local_generate.py).

    Raises GadLookupError if any of the 4 master-table lookups misses.
    """
    series = cell(row, "Series")
    valve_size = cell(row, "Valve_Size ")
    rating = cell(row, "Rating")
    actuator_series = cell(row, "Act Series")
    actuator_size = cell(row, "Actuator Size")
    traval_stop = cell(row, "Traval stop")

    properties = {
        "CUSTOMER": cell(row, "Customer"),
        "PROJECT": cell(row, "Project"),
        "PO_NO": cell(row, "PO.No"),
        "END_USER": cell(row, "End User"),
        "TAG_NO": cell(row, "Tag No"),
        "SERIAL_NO": cell(row, "Serial No"),
        "DWG_NO": cell(row, "Dwg No"),
        "DATE": cell(row, "Date"),
        "DRN": cell(row, "DRN"),
        "CHD": cell(row, "CHD"),
        "APR": cell(row, "APR"),
        "Line1": f"SERIES {series} Valve",
        "Line2": f"SIZE {valve_size} in {rating}",
        "Line3": f"SERIES {actuator_series} ACTUATOR",
        "Line4": actuator_size,
        "Iss": "1",
        "OA_NO.": cell(row, "OANo"),
        "Rating.": rating,
        "Act Series.": actuator_series,
        "AFR Port": cell(row, "AFR Port"),
        "Pos Ele Port": cell(row, "Pos Ele Port"),
        "Valve Model/Make": cell(row, "Valve Model/Make"),
        "Actuator Model/Make": cell(row, "Actuator Model/Make"),
        "AFR Model/Make": cell(row, "AFR Model/Make"),
        "Positioner Model/Make": cell(row, "Positioner Model/Make"),
        "Body Material": cell(row, "Body Material"),
        "Bonnet Material": cell(row, "Bonnet Material"),
        "Plug+Stem Material": cell(row, "PlugStem Material"),
        "Seat Material": cell(row, "Seat Material"),
        "Clamp Material": cell(row, "Seat Material"),
        "Guide bush Material": cell(row, "Seat Material"),
        "Body Gasket Material": cell(row, "Body Gasket"),
        "Seat Gasket Material": cell(row, "Seat Gasket"),
        "Packing Material": cell(row, "Packing"),
        "Trim Material": cell(row, "Trim Material"),
        "Gland Flange Material": cell(row, "Gland Flange Cluster"),
        "Stud Material": cell(row, "Stud Material"),
        "Nut Material": cell(row, "Nut Material"),
        "Stud Qty": cell(row, "Stud Qty"),
        "Nut Qty": cell(row, "Nut Qty"),
        "Fail Action": cell(row, "Fail Action"),
        "Set Pressure": cell(row, "Set Pressure"),
        "Spring": cell(row, "Spring"),
        "Size": valve_size,
        "Act Model No.": cell(row, "Actuator Model"),
        "CV/Char": cell(row, "CVChar"),
        "Port Size": cell(row, "Port Size"),
        "Rating/CWP": f"{rating}/{cell(row, 'CWP')}",
        "Travel": traval_stop,
        "Model No.": cell(row, "Valve Model"),
        "Act Size": actuator_size,
    }
    properties.update(find_dimension_values(row))

    return {
        "drawing_no": find_drawing_no(row),
        "hookup_no": find_hookup_no(row),
        # Not find_crosssec_drawing_no()/GACrossSecGlobe - Sheet2 no longer
        # uses that lookup at all (see solidworks_automation.py's
        # _generate_sheet2), replaced outright by sheet4_no's whole-drawing
        # copy. That table/lookup function are left in place (unused) rather
        # than deleted, in case this ever needs reverting.
        "sheet4_no": find_sheet4_drawing_no(row),
        # Fixed placeholder, same as the original app - not looked up per-row.
        "nameplate": "MSD",
        "properties": properties,
    }
