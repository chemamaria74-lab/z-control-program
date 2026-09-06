from datetime import date
from decimal import Decimal
from pathlib import Path

from services.motive_sync import (
    GALLONS_TO_LITERS, normalize_driver_event, normalize_fault, normalize_fuel_purchase,
    normalize_inspection, normalize_speeding_event, normalize_vehicle,
    normalize_vehicle_mileage, normalize_vehicle_utilization, _event_lookback_dates, _inspection_lookback_dates, _lookback_dates,
    _daily_metrics, _merge_motive_events, _official_requester_uuid, normalize_currency,
)
from services.motive import motive_get_all_pages_flexible


def test_incremental_event_window_rechecks_late_motive_changes(monkeypatch):
    monkeypatch.delenv("MOTIVE_EVENT_LOOKBACK_DAYS", raising=False)
    start, end = _event_lookback_dates(False)
    assert (date.fromisoformat(end) - date.fromisoformat(start)).days == 30


def test_event_window_can_be_configured(monkeypatch):
    monkeypatch.setenv("MOTIVE_EVENT_LOOKBACK_DAYS", "21")
    start, end = _event_lookback_dates(False)
    assert (date.fromisoformat(end) - date.fromisoformat(start)).days == 21


def test_incremental_operational_window_covers_manager_report(monkeypatch):
    monkeypatch.delenv("MOTIVE_INCREMENTAL_LOOKBACK_DAYS", raising=False)
    start, end = _lookback_dates(False)
    assert (date.fromisoformat(end) - date.fromisoformat(start)).days == 30


def test_incremental_inspection_window_rechecks_late_repairs(monkeypatch):
    monkeypatch.delenv("MOTIVE_INSPECTION_LOOKBACK_DAYS", raising=False)
    start, end = _inspection_lookback_dates(False)
    assert (date.fromisoformat(end) - date.fromisoformat(start)).days == 365


def test_daily_metrics_seed_confirmed_zero_days_after_complete_trip_sync():
    rows = _daily_metrics(
        integration_id=1, tenant_id="tenant", periods=[], fuels=[], inspections=[], defects=[],
        coverage_vehicle_ids=[10, 11], coverage_start="2026-08-27", coverage_end="2026-08-28",
    )
    assert {(row["vehicle_id"], row["metric_date"]) for row in rows} == {
        (10, "2026-08-27"), (10, "2026-08-28"),
        (11, "2026-08-27"), (11, "2026-08-28"),
    }
    assert all(row["distance_km"] == 0 for row in rows)


def test_sync_requester_rejects_internal_ids_and_keeps_auth_uuid():
    auth_uuid = "2883a5c0-1e8c-416f-a13a-6dc525825374"
    assert _official_requester_uuid("internal:41") is None
    assert _official_requester_uuid(None) is None
    assert _official_requester_uuid(auth_uuid) == auth_uuid


def test_updated_motive_event_wins_when_status_changes_to_dismissed():
    original = {"driver_performance_event": {"id": 10, "coaching_status": "pending_review"}}
    updated = {"driver_performance_event": {"id": 10, "coaching_status": "dismissed"}}
    rows = _merge_motive_events([original], [updated])
    assert len(rows) == 1
    assert rows[0]["driver_performance_event"]["coaching_status"] == "dismissed"


def test_fast_safety_sync_is_bounded_to_recent_event_window():
    source = Path("services/motive_sync.py").read_text()
    start = source.index("def sync_motive_safety")
    end = source.index("def sync_motive_tenant", start)
    fast_sync = source[start:end]
    assert '"/v2/driver_performance_events"' in fast_sync
    assert '"/v1/speeding_events"' in fast_sync
    assert '"/v2/inspection_reports"' in fast_sync
    assert '"inspection_reports"' in fast_sync
    assert '"/v1/fuel_purchases"' in fast_sync
    assert '"/v1/fault_codes"' in fast_sync
    assert '"updated_after": event_start_date' not in fast_sync
    assert "progress=event_progress" in fast_sync
    assert '"/v1/driving_periods"' not in fast_sync


def test_fuel_normalizer_accepts_mexican_peso_currency_labels():
    assert normalize_currency("Mexican Pesos") == "MXN"
    assert normalize_currency("MX$") == "MXN"
    row = normalize_fuel_purchase({"id": 5, "purchased_at": "2026-08-25T10:00:00Z", "fuel": "45", "fuel_unit": "ltr", "currency": "Mexican Peso", "vehicle": {"id": 8}}, integration_id=1, tenant_id="tenant")
    assert row["currency"] == "MXN"


def test_vehicle_normalizer_keeps_only_dashboard_fields():
    row = normalize_vehicle({"vehicle": {"id": 9, "number": "U-09", "vin": "private", "year": "2024", "current_driver": {"first_name": "Ana", "last_name": "Luz"}}}, integration_id=1, tenant_id="tenant")
    assert row["motive_id"] == 9
    assert row["vehicle_number"] == "U-09"
    assert row["model_year"] == 2024
    assert row["current_driver_name"] == "Ana Luz"
    assert "vin" not in row


def test_fuel_normalizer_converts_imperial_units():
    row = normalize_fuel_purchase({"id": 4, "purchased_at": "2026-07-01T12:00:00Z", "fuel": "10", "fuel_unit": "gal", "odometer": "100", "odometer_unit": "mi", "vehicle": {"id": 8}}, integration_id=1, tenant_id="tenant")
    assert row["quantity_liters"] == float(round(Decimal(10) * GALLONS_TO_LITERS, 4))
    assert row["odometer_km"] == 160.934
    assert row["motive_vehicle_id"] == 8


def test_inspection_normalizer_extracts_driver_and_nested_defects():
    inspection, defects = normalize_inspection({"inspection_report": {"id": 3, "time": "2026-07-02T10:00:00Z", "vehicle": {"id": 8}, "driver": {"id": 4, "first_name": "Ana", "last_name": "López"}, "inspected_parts": [{"id": 2, "category": "Frenos", "status": "open", "defects": [{"title": "Presión baja", "severity": "major"}]}]}}, integration_id=1, tenant_id="tenant")
    assert inspection["motive_vehicle_id"] == 8
    assert inspection["motive_driver_id"] == 4
    assert inspection["driver_name"] == "Ana López"
    assert len(defects) == 1
    assert defects[0]["title"] == "Presión baja"
    assert defects[0]["severity"] == "major"


def test_inspection_normalizer_records_repaired_part_resolution():
    inspection, defects = normalize_inspection({"inspection_report": {
        "id": 4, "time": "2026-07-31T18:12:00Z", "status": "resolved",
        "reviewer_signed_at": "2026-08-04T22:18:00Z", "vehicle": {"id": 8},
        "inspected_parts": [{
            "id": 2, "category": "Llantas", "status": "repaired",
            "mechanic_details": {"mechanic_signed_at": "2026-08-04T22:18:00Z"},
            "defects": [{"title": "Profundidad baja", "severity": "minor"}],
        }],
    }}, integration_id=1, tenant_id="tenant")
    assert inspection["status"] == "resolved"
    assert defects[0]["status"] == "repaired"
    assert defects[0]["resolved_at"] == "2026-08-04T22:18:00+00:00"


def test_driver_event_keeps_behaviors_but_not_camera_urls():
    row = normalize_driver_event({"driver_performance_event": {"id": 4, "start_time": "2026-07-20T10:00:00Z",
        "type": "cell_phone", "primary_behavior": ["cell_phone"], "vehicle": {"id": 8},
        "driver": {"id": 2, "first_name": "Ana", "last_name": "López"}, "camera_media": {"url": "private"}}},
        integration_id=3, tenant_id="tenant")
    assert row["primary_behavior"] == "cell_phone"
    assert row["driver_name"] == "Ana López"
    assert "camera_media" not in row["raw_metadata"]


def test_driver_event_marks_motive_dismissal_note_as_discarded():
    row = normalize_driver_event({"driver_performance_event": {
        "id": 5, "start_time": "2026-08-07T19:43:00Z", "type": "near_collision",
        "primary_behavior": ["near_collision"], "vehicle": {"id": 89},
        "driver": {"id": 3, "first_name": "Omar"},
        "coaching_status": "pending_review",
        "notes": "Motive - Event dismissed because the driver is not at fault",
    }}, integration_id=3, tenant_id="tenant")
    assert row["raw_metadata"]["is_discarded"] is True
    assert row["raw_metadata"]["review_texts"]


def test_driver_event_does_not_discard_pending_review_without_note():
    row = normalize_driver_event({"driver_performance_event": {
        "id": 6, "start_time": "2026-08-08T10:00:00Z", "type": "hard_brake",
        "primary_behavior": ["hard_brake"], "vehicle": {"id": 89},
        "driver": {"id": 3, "first_name": "Omar"},
        "coaching_status": "pending_review",
    }}, integration_id=3, tenant_id="tenant")
    assert row["raw_metadata"]["is_discarded"] is False


def test_driver_event_recognizes_dismissed_status_variant_and_nested_tags():
    row = normalize_driver_event({"driver_performance_event": {
        "id": 7, "start_time": "2026-08-05T13:14:00Z", "type": "driver_facing_cam_obstruction",
        "primary_behavior": ["driver_facing_cam_obstruction"], "vehicle": {"id": 97},
        "driver": {"id": 2, "first_name": "Jose"},
        "coaching_status": "dismissed_by_fm",
        "metadata": {"annotation_tags": ["driver_facing_cam_obstruction"]},
    }}, integration_id=3, tenant_id="tenant")
    assert row["raw_metadata"]["is_discarded"] is True
    assert row["raw_metadata"]["annotation_tags"] == ["driver_facing_cam_obstruction"]


def test_driver_event_treats_motive_uncoachable_as_discarded():
    row = normalize_driver_event({"driver_performance_event": {
        "id": 8, "start_time": "2026-08-05T13:14:48Z",
        "type": "driver_facing_cam_obstruction",
        "primary_behavior": ["driver_facing_cam_obstruction"],
        "vehicle": {"id": 97}, "driver": {"id": 2, "first_name": "Jose"},
        "coaching_status": "uncoachable",
    }}, integration_id=3, tenant_id="tenant")
    assert row["coaching_status"] == "uncoachable"
    assert row["raw_metadata"]["is_discarded"] is True


def test_fault_and_speeding_normalizers():
    fault = normalize_fault({"fault_code": {"id": 9, "code": "P0420", "status": "open", "vehicle": {"id": 8}}}, integration_id=3, tenant_id="tenant")
    speeding = normalize_speeding_event({"speeding_event": {"id": 7, "start_time": "2026-07-20T10:00:00Z", "max_over_speed_in_kph": "21", "vehicle": {"id": 8}}}, integration_id=3, tenant_id="tenant")
    assert fault["source_key"] == "9" and fault["status"] == "open"
    assert speeding["max_over_kph"] == 21.0


def test_vehicle_utilization_normalizer_separates_engine_and_consumed_fuel():
    row = normalize_vehicle_utilization(
        {"vehicle_idle_rollup": {
            "vehicle": {"id": 8},
            "utilization": "81.42",
            "driving_time": 7200,
            "idle_time": 1800,
            "driving_fuel": "20.5",
            "idle_fuel": "2.5",
        }},
        integration_id=3,
        tenant_id="tenant",
        period_start="2026-07-01",
        period_end="2026-07-25",
    )
    assert row["motive_vehicle_id"] == 8
    assert row["utilization_pct"] == 0.8142
    assert row["driving_hours"] == 2
    assert row["idle_hours"] == 0.5
    assert row["engine_hours"] == 2.5
    assert row["fuel_consumed_liters"] == 23


def test_vehicle_mileage_normalizer_converts_miles_and_accepts_wrappers():
    motive_id, distance_km = normalize_vehicle_mileage({
        "ifta_summary": {
            "vehicle": {"id": 8, "metric_units": False},
            "distance": "100",
        }
    })
    assert motive_id == 8
    assert distance_km == 160.934


def test_flexible_mileage_pages_accept_unknown_nested_collection(monkeypatch):
    monkeypatch.setattr(
        "services.motive.motive_get",
        lambda *_args, **_kwargs: {
            "result": {"vehicle_mileage": [{
                "ifta_summary": {"vehicle": {"id": 8}, "distance": "25"}
            }]},
            "pagination": {"page_no": 1},
        },
    )
    rows = motive_get_all_pages_flexible(
        "/v1/ifta/summary",
        collection_keys=("summaries",),
    )
    assert rows[0]["ifta_summary"]["distance"] == "25"
