from email.message import EmailMessage

from outlook_skill.downloader import extract_metadata


def test_extract_metadata_reads_headers():
    message = EmailMessage()
    message["Message-ID"] = "<abc@example.com>"
    message["Subject"] = "Hello"
    message["From"] = "Duck <duck@example.com>"
    message["Date"] = "Tue, 14 Apr 2026 10:00:00 +0000"

    metadata = extract_metadata(message.as_bytes())

    assert metadata["message_id"] == "<abc@example.com>"
    assert metadata["subject"] == "Hello"
    assert metadata["from_addr"] == "duck@example.com"
    assert metadata["received_at"] == "Tue, 14 Apr 2026 10:00:00 +0000"
