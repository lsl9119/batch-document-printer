"""
COM 调用辅助工具

- 统一的 COM 错误解析（HRESULT → 中文说明）
- "签名不兼容"时的降级调用（WPS 的 IDispatch 对命名参数/可选参数的支持与 MS Office 不完全一致）
- 服务器忙（RPC_E_CALL_REJECTED）时的自动重试
"""
import time
from typing import Callable, Optional, Sequence, Tuple, TypeVar

T = TypeVar("T")

# ---- HRESULT 常量（有符号 32 位） ----
RPC_E_CALL_REJECTED = -2147418111        # 0x80010001 服务器忙
RPC_E_SERVERCALL_RETRYLATER = -2147417846  # 0x8001010A
RPC_E_DISCONNECTED = -2147417848         # 0x80010108 对象已与客户端断开
RPC_E_SERVERFAULT = -2147417851          # 0x80010105
RPC_S_SERVER_UNAVAILABLE = -2147023174   # 0x800706BA RPC 服务器不可用
RPC_S_CALL_FAILED = -2147023170          # 0x800706BE 远程过程调用失败
CO_E_SERVER_EXEC_FAILURE = -2146959355   # 0x80080005 服务器执行失败
REGDB_E_CLASSNOTREG = -2147221164        # 0x80040154 类未注册
CO_E_CLASSSTRING = -2147221005           # 0x800401F3 无效的类字符串（ProgID 未注册）
DISP_E_MEMBERNOTFOUND = -2147352573      # 0x80020003
DISP_E_PARAMNOTFOUND = -2147352572       # 0x80020004
DISP_E_TYPEMISMATCH = -2147352571        # 0x80020005
DISP_E_UNKNOWNNAME = -2147352570         # 0x80020006
DISP_E_EXCEPTION = -2147352567           # 0x80020009 应用程序内部抛出的异常
DISP_E_BADPARAMCOUNT = -2147352562       # 0x8002000E
DISP_E_PARAMNOTOPTIONAL = -2147352561    # 0x8002000F

BUSY_HRESULTS = {RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER}
DEAD_SERVER_HRESULTS = {RPC_E_DISCONNECTED, RPC_E_SERVERFAULT, RPC_S_SERVER_UNAVAILABLE,
                        RPC_S_CALL_FAILED, CO_E_SERVER_EXEC_FAILURE}
NOT_REGISTERED_HRESULTS = {REGDB_E_CLASSNOTREG, CO_E_CLASSSTRING}
SIGNATURE_HRESULTS = {DISP_E_MEMBERNOTFOUND, DISP_E_PARAMNOTFOUND, DISP_E_TYPEMISMATCH,
                      DISP_E_UNKNOWNNAME, DISP_E_BADPARAMCOUNT, DISP_E_PARAMNOTOPTIONAL}

_HRESULT_TEXT = {
    RPC_E_CALL_REJECTED: "应用程序忙，拒绝了调用",
    RPC_E_SERVERCALL_RETRYLATER: "应用程序忙，请稍后重试",
    RPC_E_DISCONNECTED: "COM 对象已断开（应用程序可能已退出或崩溃）",
    RPC_E_SERVERFAULT: "COM 服务器内部错误",
    RPC_S_SERVER_UNAVAILABLE: "RPC 服务器不可用（应用程序进程已结束）",
    RPC_S_CALL_FAILED: "远程过程调用失败（应用程序进程已结束）",
    CO_E_SERVER_EXEC_FAILURE: "COM 服务器启动失败",
    REGDB_E_CLASSNOTREG: "COM 类未注册",
    CO_E_CLASSSTRING: "COM ProgID 未注册",
    DISP_E_MEMBERNOTFOUND: "方法或属性不存在",
    DISP_E_PARAMNOTFOUND: "参数不存在",
    DISP_E_TYPEMISMATCH: "参数类型不匹配",
    DISP_E_UNKNOWNNAME: "未知的名称",
    DISP_E_BADPARAMCOUNT: "参数个数错误",
    DISP_E_PARAMNOTOPTIONAL: "缺少必需参数",
}


def get_hresult(error: BaseException) -> Optional[int]:
    """从 pywintypes.com_error（或任意带 hresult 的异常）中提取 HRESULT"""
    hresult = getattr(error, "hresult", None)
    if hresult is None:
        args = getattr(error, "args", ())
        if args and isinstance(args[0], int):
            hresult = args[0]
    if hresult is None:
        return None
    # 规范化为有符号 32 位
    hresult = int(hresult) & 0xFFFFFFFF
    if hresult >= 0x80000000:
        hresult -= 0x100000000
    return hresult


def get_inner_scode(error: BaseException) -> Optional[int]:
    """DISP_E_EXCEPTION 时，excepinfo 中的 scode 才是应用程序返回的真实错误码"""
    args = getattr(error, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 6:
        scode = args[2][5]
        if isinstance(scode, int) and scode:
            return get_hresult(Exception(scode))
    return None


def describe_com_error(error: BaseException) -> str:
    """把 COM 异常转换为可读的中文描述（保留原始 HRESULT 以便诊断）"""
    hresult = get_hresult(error)
    if hresult is None:
        return str(error) or error.__class__.__name__
    args = getattr(error, "args", ())
    parts = []
    base = _HRESULT_TEXT.get(hresult)
    if base:
        parts.append(base)
    elif len(args) >= 2 and args[1]:
        parts.append(str(args[1]))
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 3:
        source, description = args[2][1], args[2][2]
        if description:
            parts.append(f"{source + ': ' if source else ''}{str(description).strip()}")
    parts.append(f"HRESULT=0x{hresult & 0xFFFFFFFF:08X}")
    return "；".join(parts)


def is_busy_error(error: BaseException) -> bool:
    return get_hresult(error) in BUSY_HRESULTS or get_inner_scode(error) in BUSY_HRESULTS


def is_dead_server_error(error: BaseException) -> bool:
    return get_hresult(error) in DEAD_SERVER_HRESULTS


def is_not_registered_error(error: BaseException) -> bool:
    return get_hresult(error) in NOT_REGISTERED_HRESULTS


def is_signature_error(error: BaseException) -> bool:
    """调用方式不被支持（命名参数、参数个数、类型不匹配等），可以换一种调用方式重试"""
    if isinstance(error, TypeError):
        return True
    return get_hresult(error) in SIGNATURE_HRESULTS or get_inner_scode(error) in SIGNATURE_HRESULTS


def looks_like_password_error(error: BaseException) -> bool:
    text = describe_com_error(error).lower()
    keywords = ("password", "密码", "加密", "encrypted", "protected", "保护")
    return any(k in text for k in keywords)


def com_retry(func: Callable[[], T], attempts: int = 6, delay: float = 0.5) -> T:
    """应用程序忙时自动重试（指数退避）"""
    last_error: Optional[BaseException] = None
    for i in range(attempts):
        try:
            return func()
        except Exception as e:  # noqa: BLE001 - 需要检查 COM 异常类型
            if not is_busy_error(e):
                raise
            last_error = e
            time.sleep(delay * (2 ** i))
    assert last_error is not None
    raise last_error


def call_with_fallbacks(attempts: Sequence[Tuple[str, Callable[[], T]]],
                        log: Optional[Callable[[str], None]] = None) -> Tuple[T, str]:
    """
    依次尝试多种调用方式，仅在"签名不兼容"错误时换下一种。

    Args:
        attempts: [(方式描述, 可调用对象), ...]
        log: 日志回调

    Returns:
        (返回值, 成功的方式描述)

    Raises:
        最后一次（或第一个非签名类）异常
    """
    last_error: Optional[BaseException] = None
    for label, func in attempts:
        try:
            return com_retry(func), label
        except Exception as e:  # noqa: BLE001
            if not is_signature_error(e):
                raise
            last_error = e
            if log:
                log(f"调用方式 [{label}] 不被支持，尝试下一种: {describe_com_error(e)}")
    assert last_error is not None
    raise last_error


def safe_get(obj, attr: str, default=None):
    """读取 COM 属性，失败时返回默认值"""
    try:
        return getattr(obj, attr)
    except Exception:  # noqa: BLE001 - COM 属性读取失败不影响主流程
        return default


def safe_set(obj, attr: str, value, log: Optional[Callable[[str], None]] = None) -> bool:
    """设置 COM 属性，返回是否成功"""
    try:
        setattr(obj, attr, value)
        return True
    except Exception as e:  # noqa: BLE001
        if log:
            log(f"设置 {attr}={value!r} 失败: {describe_com_error(e)}")
        return False
