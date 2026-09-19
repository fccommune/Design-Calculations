"""
Maps a GAD Automation BOM Excel column to the master lookup table column(s)
that define its valid values, so the Globe page's row editor can render a
dropdown of known values and flag a value that doesn't match any of them -
a cheap proxy for "this configuration will likely fail to find a GA /
hookup / cross-section / dimension match" before the user even clicks
Generate.

Each BOM header maps to one or more (model, column) pairs because the same
real-world field is sometimes stored under a different column name in
different master tables (e.g. the BOM's "Series" feeds both
GAGlobeTable.valve_series and GADimValveGlobe.series) - a field's valid
values are the union across all of them. See app/services/gad_globe_lookup.py
for the exact header strings and how each one is actually used to query.
"""
from app.extensions import db
from app.models import (
    GAGlobeTable,
    GAGlobeHookUp,
    GACrossSecGlobe,
    GASheet4Globe,
    GADimValveGlobe,
    GADimActGlobe,
)

FIELD_SOURCES = {
    "Series": [(GAGlobeTable, "valve_series"), (GADimValveGlobe, "series")],
    "Body_Style": [(GAGlobeTable, "body_style"), (GACrossSecGlobe, "body_style"), (GASheet4Globe, "body_style"), (GADimValveGlobe, "body_style")],
    # Valve_Size / Rating deliberately don't include GASheet4Globe - that
    # table stores them as ranges (e.g. '1"-4"', '150-2500'), not a single
    # value or comma list, so splitting on "," (below) would turn a whole
    # range into one garbage "option" instead of real selectable values.
    "Valve_Size ": [(GAGlobeTable, "valve_size"), (GACrossSecGlobe, "size"), (GADimValveGlobe, "size")],
    "Rating": [(GAGlobeTable, "rating"), (GADimValveGlobe, "rating")],
    "End _Connection": [(GAGlobeTable, "end_connection"), (GASheet4Globe, "end_connection"), (GADimValveGlobe, "end_connection")],
    "Bonnet _Type": [(GAGlobeTable, "bonnet_type"), (GACrossSecGlobe, "bonnet_type"), (GASheet4Globe, "bonnet_type"), (GADimValveGlobe, "bonnet_type")],
    "Flow_Direction": [(GAGlobeTable, "flow_direction"), (GACrossSecGlobe, "flow_direction"), (GASheet4Globe, "flow_direction")],
    "Act Series": [(GAGlobeTable, "actuator_series")],
    "Act Type": [(GAGlobeTable, "actuator_type"), (GAGlobeHookUp, "actuator"), (GADimActGlobe, "actuator_type")],
    "Actuator Size": [(GAGlobeTable, "actuator_size"), (GADimActGlobe, "actuator_size")],
    "Traval stop": [(GAGlobeTable, "traval_stop"), (GADimActGlobe, "travel_stop")],
    "H/W": [(GAGlobeTable, "hw"), (GADimActGlobe, "hand_wheel")],
    "End_Finish": [(GADimValveGlobe, "end_finish")],
    "Stem_Dia": [(GADimValveGlobe, "stem_dia")],
    "Trim Type": [(GACrossSecGlobe, "trim_type"), (GASheet4Globe, "trim_type")],
    "Balancing": [(GACrossSecGlobe, "balancing"), (GASheet4Globe, "balancing")],
    "Bal Seal Type": [(GACrossSecGlobe, "bal_seal_type"), (GASheet4Globe, "bal_seal_type")],
    "Seat Type": [(GACrossSecGlobe, "seat_type"), (GASheet4Globe, "seat_type")],
    "Packing Type": [(GACrossSecGlobe, "packing_type"), (GASheet4Globe, "packing_type")],
    "Spring": [(GAGlobeHookUp, "spring")],
    "AirFailaction": [(GAGlobeHookUp, "air_fail_action")],
    "Positioner": [(GAGlobeHookUp, "positioner")],
    "HandAuto": [(GAGlobeHookUp, "hand_auto")],
    "LimitSwitchType": [(GAGlobeHookUp, "limit_switch")],
    "PositionTrans": [(GAGlobeHookUp, "position_trans")],
    "VolumeBooster": [(GAGlobeHookUp, "volume_booster")],
    "SpoolValve": [(GAGlobeHookUp, "spool_valve")],
    "Airlockrelay": [(GAGlobeHookUp, "lock_valve")],
    "Solenoid Valve": [(GAGlobeHookUp, "solenoid_valve")],
}


def get_field_options(field_name: str):
    """Sorted, de-duplicated valid values for one BOM field, unioned across
    every master column that field feeds. Returns [] for a field with no
    master-table source (free-text fields like Tag No, Customer, ...) or
    one whose master table(s) are still empty."""
    sources = FIELD_SOURCES.get(field_name)
    if not sources:
        return []

    values = set()
    for model, column in sources:
        col_attr = getattr(model, column)
        for (v,) in db.session.query(col_attr).distinct():
            if v is None:
                continue
            if v == "":
                values.add("")
                continue
            # a master cell can itself hold a comma-list (e.g. "1,1.5") that
            # a single BOM value is matched against token-by-token - see
            # _token_match in gad_globe_lookup.py - so split those out into
            # individual options rather than one literal "1,1.5" choice.
            for token in str(v).split(","):
                token = token.strip()
                if token:
                    values.add(token)
    return sorted(values)
