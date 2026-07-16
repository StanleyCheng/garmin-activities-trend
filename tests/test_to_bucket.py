import pytest
from transform import to_bucket


@pytest.mark.parametrize("km,bucket", [
    (0.0, "<3"),
    (2.99, "<3"),
    (2.9999, "<3"),
    (3.0, "3-5"),
    (4.99, "3-5"),
    (5.0, "5-10"),
    (9.99, "5-10"),
    (10.0, "10-15"),
    (14.99, "10-15"),
    (15.0, "15-25"),
    (24.99, "15-25"),
    (25.0, "25-40"),
    (39.99, "25-40"),
    (40.0, "40+"),
    (100.0, "40+"),
    (None, "all"),
])
def test_to_bucket(km, bucket):
    assert to_bucket(km) == bucket
