import time
from garminconnect import (
    Garmin,
    GarminConnectConnectionError,
    GarminConnectTooManyRequestsError,
)


DEFAULT_MAX_ACTIVITIES = 5000
DEFAULT_BATCH_SIZE = 100
MAX_BATCH_SIZE = 100
DEFAULT_RETRIES = 3


def get_client(username, password):
    client = Garmin(username, password)
    client.login()
    return client


def fetch_activities(client, *, max_activities=DEFAULT_MAX_ACTIVITIES,
                    batch_size=DEFAULT_BATCH_SIZE, retries=DEFAULT_RETRIES):
    """Fetch up to max_activities from Garmin Connect, paginated.

    Retries on transient errors. Raises TooManyRequestsError after final retry.
    """
    if batch_size > MAX_BATCH_SIZE:
        batch_size = MAX_BATCH_SIZE
    activities, start = [], 0
    while start < max_activities:
        limit = min(batch_size, max_activities - start)
        last_exc = None
        for attempt in range(1, retries + 1):
            try:
                batch = client.get_activities(start, limit)
                break
            except (GarminConnectTooManyRequestsError,
                    GarminConnectConnectionError) as exc:
                last_exc = exc
                if attempt == retries:
                    raise
                delay = 2 ** (attempt - 1)
                print(f"Garmin fetch retry {attempt}/{retries} after {delay}s: {exc}")
                time.sleep(delay)
        if not batch:
            break
        activities.extend(batch)
        fetched = len(batch)
        start += fetched
        if fetched < limit:
            break
    return activities
