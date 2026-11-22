"""Serialize native Word/MathType access across local CLI, GUI and MCP workers."""
from __future__ import annotations

from contextlib import contextmanager
import ctypes
from ctypes import wintypes
import os
import threading
import time


_THREAD_LOCK = threading.RLock()
_NAME = "Local\\PaperFlow.NativeWord.MathType"


@contextmanager
def native_word_lock(*, timeout=None):
    """Own the native engine until cleanup completes; waiting can be cancelled.

    A Windows kernel mutex is released when a crashed worker exits. An
    abandoned mutex is acquired normally, without restarting another worker.
    """
    if os.name != "nt":
        with _THREAD_LOCK:
            yield
        return
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.CreateMutexW.argtypes = [ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR]
    kernel.CreateMutexW.restype = wintypes.HANDLE
    kernel.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    kernel.WaitForSingleObject.restype = wintypes.DWORD
    kernel.ReleaseMutex.argtypes = [wintypes.HANDLE]
    kernel.ReleaseMutex.restype = wintypes.BOOL
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    handle = kernel.CreateMutexW(None, False, _NAME)
    if not handle:
        raise OSError(ctypes.get_last_error(), "Cannot create the PaperFlow native-engine mutex")
    acquired = False
    start = time.monotonic()
    try:
        while not acquired:
            if timeout is not None and time.monotonic() - start >= timeout:
                raise TimeoutError("Another PaperFlow worker is using Word/MathType")
            wait_ms = 250 if timeout is None else max(1, min(250, int((timeout - (time.monotonic() - start)) * 1000)))
            result = kernel.WaitForSingleObject(handle, wait_ms)
            if result in (0, 0x80):
                acquired = True
            elif result != 0x102:
                raise OSError(ctypes.get_last_error(), "Cannot wait for the PaperFlow native engine")
        yield
    finally:
        if acquired:
            kernel.ReleaseMutex(handle)
        kernel.CloseHandle(handle)
