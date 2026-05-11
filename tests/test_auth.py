from outlook_skill.auth import filter_reserved_scopes


def test_filter_reserved_scopes_removes_msal_reserved_values():
    scopes = ("Mail.Read", "offline_access", "openid", "profile")

    assert filter_reserved_scopes(scopes) == ("Mail.Read",)
