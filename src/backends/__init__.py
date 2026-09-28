"""
Office 引擎后端（WPS Office 优先，Microsoft Office 可选回退）
"""
from .office_backend import (BackendSelection, BackendUnavailableError, ComponentStatus,
                             KIND_BY_FILE_TYPE, OfficeBackend, OfficeKind, OfficeSession,
                             OfficeSessionPool, OfficeStartError, Watchdog, com_initialized,
                             default_backends, detect_backend, select_backends)
from .wps_backend import WpsBackend
from .ms_office_backend import MsOfficeBackend

__all__ = [
    "BackendSelection", "BackendUnavailableError", "ComponentStatus", "KIND_BY_FILE_TYPE",
    "OfficeBackend", "OfficeKind", "OfficeSession", "OfficeSessionPool", "OfficeStartError",
    "Watchdog", "com_initialized", "default_backends", "detect_backend", "select_backends",
    "WpsBackend", "MsOfficeBackend",
]
