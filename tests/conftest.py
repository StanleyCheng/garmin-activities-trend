from datetime import datetime
import pytest


def _make(*, activity_id=1, type_="running", distance_m=10_000, duration_s=3_000, avg_speed_mps=3.33,
          avg_hr=150, start_time_local="2024-03-15T06:30:00", avg_pace_s_per_km=None,
          calories=800, elevation=50, vo2max=52):
    raw = {
        "activityId": activity_id,
        "activityType": {"typeKey": type_},
        "startTimeLocal": start_time_local,
        "distance": distance_m,
        "duration": duration_s,
        "averageSpeed": avg_speed_mps,
        "averageHR": avg_hr,
        "calories": calories,
        "elevationGain": elevation,
        "vO2MaxValue": vo2max,
    }
    if avg_pace_s_per_km is not None:
        raw["averagePace"] = 1.0 / (avg_pace_s_per_km * avg_speed_mps) if avg_speed_mps else None
    return raw


@pytest.fixture
def synthetic_activity():
    return _make()
