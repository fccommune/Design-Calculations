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
    GADimValveGlobe,
    GADimActGlobe,
)


class GadLookupError(Exception):
    """Raised when a required lookup row can't be found - mirrors the old
    app's MessageBox.Show("No matching ... row ...") + early return."""


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
        raise GadLookupError("No matching GAGlobeTable row for this configuration.")
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
        raise GadLookupError("No matching GAGlobeHookUp row for this configuration.")
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
        raise GadLookupError("No matching GACrossSecGlobe row for this configuration.")
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
        raise GadLookupError("No matching BODY dimension row found for this configuration.")

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
        raise GadLookupError("No matching ACTUATOR dimension row found for this configuration.")

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
    the GA block, hookup schematic and cross-section drawing numbers, plus
    all ~50 title-block custom properties (dimension-table values and
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
        "crosssec_no": find_crosssec_drawing_no(row),
        # Fixed placeholder, same as the original app - not looked up per-row.
        "nameplate": "MSD",
        "properties": properties,
    }
