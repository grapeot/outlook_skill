from email.message import EmailMessage
from pathlib import Path

from outlook_skill.markdown_renderer import cleanup_markdown_noise, render_eml_file, render_message_to_markdown, simplify_tracking_url


def test_render_message_prefers_plain_text_when_present():
    message = EmailMessage()
    message["Subject"] = "Hello"
    message["From"] = "Duck <duck@example.com>"
    message["To"] = "a@example.com"
    message.set_content("This is the plain body.\n\n> older quote")
    message.add_alternative("<html><body><p>This is <b>html</b>.</p></body></html>", subtype="html")

    result = render_message_to_markdown(message, source_name="sample.eml")

    markdown = result["markdown"]
    assert "This is the plain body." in markdown
    assert "older quote" not in markdown


def test_render_message_falls_back_to_html_and_lists_attachments():
    message = EmailMessage()
    message["Subject"] = "HTML only"
    message["From"] = "Duck <duck@example.com>"
    message["To"] = "a@example.com"
    message.make_mixed()

    html_part = EmailMessage()
    html_part.set_content("<html><body><h1>Title</h1><p>Hello <b>world</b></p></body></html>", subtype="html")
    message.attach(html_part)

    attachment = EmailMessage()
    attachment.add_attachment(b"hello", maintype="text", subtype="plain", filename="notes.txt")
    for part in attachment.iter_attachments():
        message.attach(part)

    result = render_message_to_markdown(message, source_name="html-only.eml")

    markdown = result["markdown"]
    assert "# HTML only" in markdown
    assert "Hello **world**" in markdown or "Hello world" in markdown
    assert "## Attachments" in markdown
    assert "notes.txt" in markdown


def test_render_eml_file_writes_markdown_file(tmp_path: Path):
    message = EmailMessage()
    message["Subject"] = "Write file"
    message["From"] = "Duck <duck@example.com>"
    message["To"] = "a@example.com"
    message.set_content("Body")

    source = tmp_path / "mail.eml"
    source.write_bytes(message.as_bytes())
    output_dir = tmp_path / "markdowns"

    rendered = render_eml_file(source, output_dir)

    assert rendered is not None
    assert rendered.output_path.exists()
    assert "Body" in rendered.output_path.read_text(encoding="utf-8")


def test_cleanup_markdown_noise_removes_zero_width_and_tracking_image():
    noisy = "hello\u200b\n![](https://x.emltrk.com/v2/open)\n-->\nworld"

    cleaned = cleanup_markdown_noise(noisy)

    assert "\u200b" not in cleaned
    assert "emltrk" not in cleaned
    assert "-->" not in cleaned
    assert "hello" in cleaned and "world" in cleaned


def test_simplify_tracking_url_prefers_embedded_destination():
    original = "https://us3-cdn.eng.bloomreach.com/path?url=https%3A%2F%2Fexample.com%2Fhello%3Fa%3D1&token=abc"

    simplified = simplify_tracking_url(original)

    assert simplified == "https://example.com/hello?a=1"
