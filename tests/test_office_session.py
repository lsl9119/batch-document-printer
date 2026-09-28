import threading
import time

import pytest

from src.backends import process_utils
from src.backends.office_backend import (BackendSelection, OfficeKind, OfficeSession, OfficeSessionPool,
                                         OfficeStartError, Watchdog)
from src.backends.wps_backend import WpsBackend
from tests.fakes import FakeComError, FakeOfficeApp


class ProcessTable:
    """模拟系统进程表"""

    def __init__(self, preexisting=None):
        self.procs = dict(preexisting or {})
        self.next_pid = 1000
        self.killed = []

    def spawn(self, name):
        self.next_pid += 1
        self.procs[self.next_pid] = name
        return self.next_pid

    def install(self, monkeypatch):
        monkeypatch.setattr(process_utils, "PSUTIL_AVAILABLE", True)
        monkeypatch.setattr(process_utils, "snapshot",
                            lambda names: {p: n for p, n in self.procs.items() if n in {x.lower() for x in names}})
        monkeypatch.setattr(process_utils, "pid_exists", lambda pid: pid in self.procs)

        def kill(pids, reason=""):
            done = [p for p in pids if p in self.procs]
            for p in done:
                del self.procs[p]
                self.killed.append(p)
            return done

        monkeypatch.setattr(process_utils, "kill_pids", kill)
        monkeypatch.setattr(process_utils, "wait_for_exit",
                            lambda pids, timeout: {p for p in pids if p in self.procs})


def make_factory(table, process_name="wps.exe", quit_exits=True, apps=None):
    def factory(progid):
        pid = table.spawn(process_name)
        app = FakeOfficeApp(progid)
        original_quit = app.Quit

        def quit_():
            original_quit()
            if quit_exits:
                table.procs.pop(pid, None)

        app.Quit = quit_
        if apps is not None:
            apps.append(app)
        return app
    return factory


def test_session_tracks_spawned_process_and_quits(monkeypatch):
    table = ProcessTable()
    table.install(monkeypatch)
    apps = []
    session = OfficeSession(WpsBackend(), OfficeKind.WRITER, dispatch_factory=make_factory(table, apps=apps))
    session.start()
    assert session.owned and len(session.spawned_pids) == 1
    app = apps[0]
    assert app.Visible is False and app.DisplayAlerts == 0
    session.close()
    assert app.quit_called
    assert not table.procs and not table.killed  # 正常退出，无需强制结束


def test_session_kills_process_that_does_not_exit_after_quit(monkeypatch):
    table = ProcessTable()
    table.install(monkeypatch)
    session = OfficeSession(WpsBackend(), OfficeKind.SPREADSHEET,
                            dispatch_factory=make_factory(table, "et.exe", quit_exits=False))
    session.start()
    pid = next(iter(session.spawned_pids))
    session.close(quit_timeout=0.1)
    assert pid in table.killed and not table.procs


def test_attached_to_user_instance_never_quits_or_kills(monkeypatch):
    table = ProcessTable({42: "wps.exe"})  # 用户自己打开的 WPS
    table.install(monkeypatch)
    apps = []

    def attach_factory(progid):
        app = FakeOfficeApp(progid)
        apps.append(app)
        return app  # 不产生新进程

    session = OfficeSession(WpsBackend(), OfficeKind.WRITER, dispatch_factory=attach_factory)
    session.start()
    assert not session.owned
    assert session.kill("test") == []
    session.close()
    assert not apps[0].quit_called
    assert 42 in table.procs


def test_start_failure_reports_progid(monkeypatch):
    ProcessTable().install(monkeypatch)

    def failing(progid):
        raise FakeComError(-2147221005, "无效的类字符串")

    session = OfficeSession(WpsBackend(), OfficeKind.SPREADSHEET, dispatch_factory=failing)
    with pytest.raises(OfficeStartError) as info:
        session.start()
    assert "KET.Application" in str(info.value) and "WPS 表格" in str(info.value)


def _pool(table, monkeypatch, recycle_after=30, apps=None):
    table.install(monkeypatch)
    backend = WpsBackend()
    selection = BackendSelection(chosen={k: backend for k in OfficeKind})
    factory = make_factory(table, apps=apps)
    return OfficeSessionPool(selection, recycle_after=recycle_after,
                             session_factory=lambda b, k: OfficeSession(b, k, dispatch_factory=factory))


def test_pool_reuses_and_recycles_sessions(monkeypatch):
    table = ProcessTable()
    apps = []
    pool = _pool(table, monkeypatch, recycle_after=2, apps=apps)
    s1 = pool.get(OfficeKind.WRITER)
    s1.jobs_done = 1
    assert pool.get(OfficeKind.WRITER) is s1
    s1.jobs_done = 2
    s2 = pool.get(OfficeKind.WRITER)
    assert s2 is not s1 and apps[0].quit_called
    pool.close_all()
    assert not table.procs


def test_pool_restarts_dead_session_and_close_all_leaves_no_process(monkeypatch):
    table = ProcessTable()
    pool = _pool(table, monkeypatch)
    writer = pool.get(OfficeKind.WRITER)
    pool.get(OfficeKind.SPREADSHEET)
    pool.get(OfficeKind.PRESENTATION)
    assert len(table.procs) == 3
    killed = pool.kill_kind(OfficeKind.WRITER, "timeout")
    assert killed and not writer.is_alive
    writer2 = pool.get(OfficeKind.WRITER)
    assert writer2 is not writer
    pool.close_all()
    assert table.procs == {}


def test_watchdog_fires_and_unblocks():
    unblock = threading.Event()
    with Watchdog(0.05, unblock.set, name="t") as wd:
        assert unblock.wait(2)
    assert wd.fired


def test_watchdog_not_fired_for_fast_operation():
    fired = []
    with Watchdog(5, lambda: fired.append(1)) as wd:
        time.sleep(0.01)
    assert not wd.fired and not fired
