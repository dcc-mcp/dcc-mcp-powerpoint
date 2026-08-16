"""Package-level smoke tests (no Office installation required)."""

from dcc_mcp_powerpoint import __version__
from dcc_mcp_powerpoint.sidecar.office_host import APP_NAME, HOST_EXE, OfficeHostConfig


def test_version_is_pep440() -> None:
    parts = __version__.split(".")
    assert len(parts) == 3
    assert all(p.isdigit() for p in parts)


def test_office_host_defaults_to_powerpoint() -> None:
    cfg = OfficeHostConfig()
    assert cfg.app == APP_NAME == "powerpoint"
    resolved = cfg.resolve_exe()
    assert resolved is None or resolved.endswith(HOST_EXE)


def test_missing_exe_raises_with_hint() -> None:
    cfg = OfficeHostConfig()
    try:
        cfg.build_args()
    except FileNotFoundError as exc:
        assert HOST_EXE in str(exc)
    else:
        # An office-host was found on this machine; args must still be shaped.
        args = cfg.build_args()
        assert args[1] == "--app=powerpoint"


def test_pipe_name_matches_protocol() -> None:
    cfg = OfficeHostConfig()
    name = cfg.pipe_name("S-1-5-21-42", 3)
    assert name == r"\\.\pipe\dcc-mcp-office-powerpoint-S-1-5-21-42-3"
