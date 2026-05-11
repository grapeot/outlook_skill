from outlook_skill.graph_client import WELL_KNOWN_FOLDERS


def test_well_known_folders_maps_inbox():
    assert WELL_KNOWN_FOLDERS["INBOX"] == "inbox"
