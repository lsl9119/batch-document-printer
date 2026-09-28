from src.backends.office_backend import (ComponentStatus, OfficeKind, _parse_server_path, select_backends)
from src.backends.ms_office_backend import MsOfficeBackend
from src.backends.wps_backend import WpsBackend


def _statuses(wps_ok=(True, True, True), ms_ok=(False, False, False), wps_registered=True):
    result = {}
    for backend, oks in ((WpsBackend(), wps_ok), (MsOfficeBackend(), ms_ok)):
        result[backend.key] = {}
        for kind, ok in zip(OfficeKind, oks):
            registered = ok or (backend.key == "wps" and wps_registered)
            status = ComponentStatus(backend.key, kind, backend.progid(kind), backend.component_name(kind),
                                     registered=registered, clsid="{X}" if registered else "",
                                     server_path="C:/x.exe" if registered else "", server_exists=ok)
            if not ok:
                status.error = "COM 注册指向的程序不存在"
            result[backend.key][kind] = status
    return result


def test_progids_are_wps_first_class():
    wps = WpsBackend()
    assert wps.progid(OfficeKind.WRITER) == "KWPS.Application"
    assert wps.progid(OfficeKind.SPREADSHEET) == "KET.Application"
    assert wps.progid(OfficeKind.PRESENTATION) == "KWPP.Application"
    assert {"wps.exe", "et.exe", "wpp.exe"} <= wps.all_process_names


def test_wps_only_environment_selects_wps_for_all():
    sel = select_backends([WpsBackend(), MsOfficeBackend()], "auto", _statuses())
    for kind in OfficeKind:
        backend, _ = sel.backend_for(kind)
        assert backend.key == "wps"
    assert sel.primary_display == "WPS Office"


def test_wps_spreadsheet_broken_reports_progid():
    sel = select_backends([WpsBackend(), MsOfficeBackend()], "auto", _statuses(wps_ok=(True, False, True)))
    backend, message = sel.backend_for(OfficeKind.SPREADSHEET)
    assert backend is None
    assert "检测到 WPS Office，但WPS 表格 COM 接口不可用。ProgID: KET.Application" in message
    assert sel.backend_for(OfficeKind.WRITER)[0].key == "wps"


def test_auto_falls_back_to_ms_office_per_component():
    sel = select_backends([WpsBackend(), MsOfficeBackend()], "auto",
                          _statuses(wps_ok=(True, False, True), ms_ok=(True, True, True)))
    assert sel.backend_for(OfficeKind.SPREADSHEET)[0].key == "msoffice"
    assert sel.backend_for(OfficeKind.WRITER)[0].key == "wps"


def test_wps_only_preference_never_uses_ms_office():
    sel = select_backends([WpsBackend(), MsOfficeBackend()], "wps",
                          _statuses(wps_ok=(False, False, False), ms_ok=(True, True, True)))
    assert not sel.any_available


def test_nothing_installed_message():
    sel = select_backends([WpsBackend(), MsOfficeBackend()], "auto",
                          _statuses(wps_ok=(False,) * 3, wps_registered=False))
    backend, message = sel.backend_for(OfficeKind.WRITER)
    assert backend is None and "请安装 WPS Office" in message


def test_parse_server_path():
    assert _parse_server_path('"C:\\Program Files\\Kingsoft\\WPS Office\\office6\\wps.exe" /Automation') == \
        "C:\\Program Files\\Kingsoft\\WPS Office\\office6\\wps.exe"
    assert _parse_server_path("C:\\WPS\\office6\\et.exe /automation -Embedding") == "C:\\WPS\\office6\\et.exe"
    assert _parse_server_path("") == ""


def test_check_component_on_non_windows_is_unavailable(monkeypatch):
    import src.backends.office_backend as ob
    monkeypatch.setattr(ob, "lookup_progid_registration", lambda progid: ("", "", f"注册表中未找到 {progid}"))
    status = WpsBackend().check_component(OfficeKind.PRESENTATION)
    assert not status.available and "KWPP.Application" in status.error
