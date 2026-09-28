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


def test_attached_to_user_instance_is_left_untouched(monkeypatch):
    table = ProcessTable({42: "wps.exe"})  # 用户自己打开的 WPS
    table.install(monkeypatch)
    apps = []

    def attach_factory(progid):
        app = FakeOfficeApp(progid)
        app.open_docs.append("用户未保存的文档")
        apps.append(app)
        return app  # 不产生新进程

    session = OfficeSession(WpsBackend(), OfficeKind.WRITER, dispatch_factory=attach_factory)
    session.start()
    app = apps[0]
    assert session.mode == OfficeSession.ATTACHED and not session.owned
    assert app.Visible is True            # 不隐藏用户的窗口
    assert app.DisplayAlerts == 0         # 临时关闭提示框
    session.close()
    assert not app.quit_called
    assert app.open_docs == ["用户未保存的文档"]  # 不关闭用户文档
    assert app.DisplayAlerts is True      # 恢复原值
    assert session.kill("test") == []     # 超时保护也不会结束用户的进程
    assert 42 in table.procs


def test_second_component_attaching_to_user_process_is_not_owned(monkeypatch):
    """我们的 wps.exe 存活时，表格组件附着到用户自己的 et.exe：不能被当作自己的实例"""
    table = ProcessTable({77: "et.exe"})
    table.install(monkeypatch)
    ours = table.spawn("wps.exe")
    session = OfficeSession(WpsBackend(), OfficeKind.SPREADSHEET,
                            dispatch_factory=lambda progid: FakeOfficeApp(progid))
    session.start(owned_pids={ours})
    assert session.mode == OfficeSession.ATTACHED
    session.close()
    assert 77 in table.procs


def test_pool_shared_process_in_integrated_mode(monkeypatch):
    table = ProcessTable()
    table.install(monkeypatch)
    ours = table.spawn("wps.exe")
    session = OfficeSession(WpsBackend(), OfficeKind.PRESENTATION,
                            dispatch_factory=lambda progid: FakeOfficeApp(progid))
    session.start(owned_pids={ours})
    assert session.mode == OfficeSession.POOL_SHARED and session.spawned_pids == {ours}
    assert session.app.Visible is False
    assert session.kill("timeout") == [ours]


def test_failed_start_kills_spawned_process(monkeypatch):
    table = ProcessTable()
    table.install(monkeypatch)

    def spawn_then_fail(progid):
        table.spawn("wps.exe")  # 例如首次启动弹窗后返回拒绝访问
        raise FakeComError(-2147024891, "拒绝访问")

    session = OfficeSession(WpsBackend(), OfficeKind.WRITER, dispatch_factory=spawn_then_fail)
    with pytest.raises(OfficeStartError):
        session.start()
    assert table.procs == {} and table.killed


def test_hung_quit_is_killed_by_close_watchdog(monkeypatch):
    table = ProcessTable()
    table.install(monkeypatch)
    release = threading.Event()
    apps = []
    factory = make_factory(table, apps=apps)

    def hanging_factory(progid):
        app = factory(progid)

        def hang():
            release.wait(10)
            raise FakeComError(-2147023174, "RPC 服务器不可用")

        app.Quit = hang
        return app

    session = OfficeSession(WpsBackend(), OfficeKind.WRITER, dispatch_factory=hanging_factory)
    session.start()
    original_kill = session.kill

    def kill_and_release(reason=""):
        killed = original_kill(reason)
        release.set()
        return killed

    session.kill = kill_and_release
    start = time.monotonic()
    session.close(quit_timeout=-14.8)  # 看门狗 = quit_timeout + 15 秒 → 0.2 秒
    assert time.monotonic() - start < 5
    assert table.procs == {}


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
