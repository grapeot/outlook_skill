from datetime import UTC, datetime

from outlook_skill.exporter import (
    build_md_filename,
    parse_datetime,
    short_graph_id,
    slugify,
    yaml_escape,
)


def test_slugify_preserves_ascii_and_collapses_spaces():
    assert slugify("Hello, World!   Test", limit=50) == "Hello-World-Test"


def test_slugify_truncates_to_limit():
    result = slugify("a" * 100, limit=10)
    assert len(result) == 10


def test_short_graph_id_returns_last_eight_alnum():
    gid = "AAkALgAAAAAAHYQDEapmEc2byACqAC-EWg0=="
    short = short_graph_id(gid)
    assert len(short) == 8
    assert short.isalnum()


def test_build_md_filename_shape():
    dt = datetime(2026, 4, 16, tzinfo=UTC)
    name = build_md_filename(received_at=dt, folder="Inbox", subject="Hello World", graph_id="ABC123XYZ987")
    assert name.startswith("2026-04-16_Inbox_Hello-World_")
    assert name.endswith(".md")


def test_parse_datetime_iso():
    dt = parse_datetime("2026-04-16T12:34:56Z")
    assert dt is not None and dt.tzinfo is not None


def test_parse_datetime_rfc2822():
    dt = parse_datetime("Thu, 16 Apr 2026 23:38:23 +0000")
    assert dt is not None and dt.year == 2026


def test_yaml_escape_quotes_colon():
    assert yaml_escape("Re: hello") == '"Re: hello"'


def test_yaml_escape_passes_plain_text():
    assert yaml_escape("plain text") == "plain text"
