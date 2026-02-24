import logging
import os

logger = logging.getLogger(__name__)


if os.name != "nt":
    def prompt_admin_credentials(parent_hwnd=None, caption=None, message=None):
        return True, ""


else:
    import ctypes
    from ctypes import wintypes

    ERROR_CANCELLED = 1223
    SEE_MASK_NOCLOSEPROCESS = 0x00000040
    INFINITE = 0xFFFFFFFF
    SW_HIDE = 0

    class SHELLEXECUTEINFO(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("fMask", ctypes.c_ulong),
            ("hwnd", wintypes.HWND),
            ("lpVerb", wintypes.LPCWSTR),
            ("lpFile", wintypes.LPCWSTR),
            ("lpParameters", wintypes.LPCWSTR),
            ("lpDirectory", wintypes.LPCWSTR),
            ("nShow", ctypes.c_int),
            ("hInstApp", wintypes.HINSTANCE),
            ("lpIDList", ctypes.c_void_p),
            ("lpClass", wintypes.LPCWSTR),
            ("hkeyClass", wintypes.HKEY),
            ("dwHotKey", wintypes.DWORD),
            ("hIconOrMonitor", wintypes.HANDLE),
            ("hProcess", wintypes.HANDLE),
        ]

    _shell32 = ctypes.WinDLL("Shell32", use_last_error=True)
    _kernel32 = ctypes.WinDLL("Kernel32", use_last_error=True)

    _shell32.ShellExecuteExW.argtypes = [ctypes.POINTER(SHELLEXECUTEINFO)]
    _shell32.ShellExecuteExW.restype = wintypes.BOOL

    _kernel32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    _kernel32.WaitForSingleObject.restype = wintypes.DWORD

    _kernel32.GetExitCodeProcess.argtypes = [
        wintypes.HANDLE,
        ctypes.POINTER(wintypes.DWORD),
    ]
    _kernel32.GetExitCodeProcess.restype = wintypes.BOOL

    _kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    _kernel32.CloseHandle.restype = wintypes.BOOL

    def _admin_probe_command():
        system_root = str(os.getenv("SystemRoot", r"C:\Windows") or r"C:\Windows")
        cmd_path = os.path.join(system_root, "System32", "cmd.exe")
        if not os.path.exists(cmd_path):
            cmd_path = "cmd.exe"
        return cmd_path, "/c exit 0"

    def _run_uac_probe(parent_hwnd):
        exe, params = _admin_probe_command()
        sei = SHELLEXECUTEINFO()
        sei.cbSize = ctypes.sizeof(SHELLEXECUTEINFO)
        sei.fMask = SEE_MASK_NOCLOSEPROCESS
        sei.hwnd = wintypes.HWND(parent_hwnd or 0)
        sei.lpVerb = "runas"
        sei.lpFile = exe
        sei.lpParameters = params
        sei.lpDirectory = None
        sei.nShow = SW_HIDE

        ok = _shell32.ShellExecuteExW(ctypes.byref(sei))
        if not ok:
            return False, ctypes.get_last_error()
        if not sei.hProcess:
            return False, 0

        try:
            _kernel32.WaitForSingleObject(sei.hProcess, INFINITE)
            exit_code = wintypes.DWORD(1)
            if not _kernel32.GetExitCodeProcess(
                sei.hProcess, ctypes.byref(exit_code)
            ):
                return False, ctypes.get_last_error()
            return exit_code.value == 0, 0
        finally:
            _kernel32.CloseHandle(sei.hProcess)

    def prompt_admin_credentials(parent_hwnd=None, caption=None, message=None):
        ok, err_code = _run_uac_probe(parent_hwnd)
        if ok:
            return True, ""
        if err_code == ERROR_CANCELLED:
            return False, "cancelled"
        logger.warning("UAC admin prompt failed with code %s", err_code)
        return (
            False,
            "Nie udało się potwierdzić uprawnień administratora "
            f"(kod systemu: {err_code}).",
        )
