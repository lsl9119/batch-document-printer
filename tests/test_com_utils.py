import pytest

from src.backends import com_utils
from tests.fakes import FakeComError, dead_server_error, password_error, signature_error


def test_hresult_parsing_and_classification():
    assert com_utils.get_hresult(signature_error()) == com_utils.DISP_E_UNKNOWNNAME
    assert com_utils.is_signature_error(signature_error())
    assert com_utils.is_signature_error(TypeError("bad kwarg"))
    assert com_utils.is_dead_server_error(dead_server_error())
    assert not com_utils.is_signature_error(dead_server_error())
    # 无符号 HRESULT 也能正确识别
    assert com_utils.get_hresult(FakeComError(0x800706BA)) == com_utils.RPC_S_SERVER_UNAVAILABLE


def test_describe_com_error_contains_hresult_and_description():
    text = com_utils.describe_com_error(password_error())
    assert "密码" in text and "HRESULT=0x80020009" in text
    assert com_utils.looks_like_password_error(password_error())


def test_call_with_fallbacks_uses_next_on_signature_error():
    calls = []

    def named():
        calls.append("named")
        raise signature_error()

    def positional():
        calls.append("positional")
        return 42

    value, label = com_utils.call_with_fallbacks([("named", named), ("positional", positional)])
    assert value == 42 and label == "positional" and calls == ["named", "positional"]


def test_call_with_fallbacks_does_not_mask_real_errors():
    def fails():
        raise dead_server_error()

    with pytest.raises(FakeComError):
        com_utils.call_with_fallbacks([("a", fails), ("b", lambda: 1)])


def test_com_retry_on_busy(monkeypatch):
    monkeypatch.setattr(com_utils.time, "sleep", lambda s: None)
    attempts = {"n": 0}

    def flaky():
        attempts["n"] += 1
        if attempts["n"] < 3:
            raise FakeComError(com_utils.RPC_E_CALL_REJECTED)
        return "ok"

    assert com_utils.com_retry(flaky) == "ok" and attempts["n"] == 3


def test_safe_get_set():
    class Obj:
        @property
        def Bad(self):
            raise RuntimeError("x")

    o = Obj()
    assert com_utils.safe_get(o, "Bad", "d") == "d"
    assert com_utils.safe_set(o, "Good", 1) and o.Good == 1
