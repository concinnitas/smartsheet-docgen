from datetime import datetime, timedelta

from docgen.photos import PhotoInfo, classify

ME = "bottchers@starktech.com"
T0 = datetime(2026, 9, 1)


def p(i, who, days, name=None):
    return PhotoInfo(i, name or f"p{i}.jpg", T0 + timedelta(days=days), who, "image/jpeg")


def test_uploader_split():
    b, a, _ = classify([p(1, ME, 0), p(2, "gc@contractor.com", 5)], [ME])
    assert (b.id, a.id) == (1, 2)


def test_uploader_case_insensitive_config():
    b, a, _ = classify([p(1, ME, 0), p(2, "gc@x.com", 1)], ["BottcherS@StarkTech.com"])
    assert (b.id, a.id) == (1, 2)


def test_multiple_per_side_pick_latest_and_note():
    b, a, notes = classify([p(1, ME, 0), p(3, ME, 2), p(2, "gc@x.com", 5)], [ME])
    assert b.id == 3 and a.id == 2 and notes


def test_pick_earliest():
    b, _, _ = classify([p(1, ME, 0), p(3, ME, 2), p(2, "gc@x.com", 5)], [ME], pick="earliest")
    assert b.id == 1


def test_single_side_falls_back_to_date():
    b, a, notes = classify([p(5, ME, 9), p(4, ME, 1), p(6, ME, 4)], [ME])
    assert (b.id, a.id) == (4, 5) and "split by date" in notes[0]


def test_single_photo_assignment():
    assert classify([p(1, ME, 0)], [ME])[:2] == (p(1, ME, 0), None)
    assert classify([p(2, "gc@x.com", 0)], [ME])[:2] == (None, p(2, "gc@x.com", 0))


def test_no_photos():
    assert classify([], [ME]) == (None, None, [])
