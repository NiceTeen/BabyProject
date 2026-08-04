from __future__ import annotations

import base64
import ctypes
import os
from ctypes import wintypes


_CREDENTIAL_PREFIX = "dpapi-v1:"
_DPAPI_ENTROPY = b"BabyProject/ImageAI/v1"
_CRYPTPROTECT_UI_FORBIDDEN = 0x1


class _DataBlob(ctypes.Structure):
    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_ubyte)),
    ]


def _data_blob(data: bytes) -> tuple[_DataBlob, object]:
    buffer = (ctypes.c_ubyte * len(data)).from_buffer_copy(data)
    blob = _DataBlob(
        len(data),
        ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)),
    )
    return blob, buffer


def _windows_libraries() -> tuple[object, object]:
    if os.name != "nt":
        raise OSError("API Key 加密仅支持 Windows。")

    crypt32 = ctypes.WinDLL("Crypt32.dll", use_last_error=True)
    kernel32 = ctypes.WinDLL("Kernel32.dll", use_last_error=True)
    crypt32.CryptProtectData.argtypes = [
        ctypes.POINTER(_DataBlob),
        wintypes.LPCWSTR,
        ctypes.POINTER(_DataBlob),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(_DataBlob),
    ]
    crypt32.CryptProtectData.restype = wintypes.BOOL
    crypt32.CryptUnprotectData.argtypes = [
        ctypes.POINTER(_DataBlob),
        ctypes.c_void_p,
        ctypes.POINTER(_DataBlob),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(_DataBlob),
    ]
    crypt32.CryptUnprotectData.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [wintypes.HLOCAL]
    kernel32.LocalFree.restype = wintypes.HLOCAL
    return crypt32, kernel32


def protect_secret(secret: str) -> str:
    if not secret:
        return ""

    crypt32, kernel32 = _windows_libraries()
    input_blob, input_buffer = _data_blob(secret.encode("utf-8"))
    entropy_blob, entropy_buffer = _data_blob(_DPAPI_ENTROPY)
    output_blob = _DataBlob()
    result = crypt32.CryptProtectData(
        ctypes.byref(input_blob),
        "BabyProject 图片 AI API Key",
        ctypes.byref(entropy_blob),
        None,
        None,
        _CRYPTPROTECT_UI_FORBIDDEN,
        ctypes.byref(output_blob),
    )
    del input_buffer, entropy_buffer
    if not result:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        protected = ctypes.string_at(output_blob.pbData, output_blob.cbData)
    finally:
        kernel32.LocalFree(ctypes.cast(output_blob.pbData, ctypes.c_void_p))
    return _CREDENTIAL_PREFIX + base64.b64encode(protected).decode("ascii")


def unprotect_secret(value: str) -> str:
    if not value:
        return ""
    if not value.startswith(_CREDENTIAL_PREFIX):
        raise ValueError("无法识别保存的 API Key 格式。")

    try:
        protected = base64.b64decode(
            value.removeprefix(_CREDENTIAL_PREFIX).encode("ascii"),
            validate=True,
        )
    except (UnicodeEncodeError, ValueError) as exc:
        raise ValueError("保存的 API Key 数据已损坏。") from exc

    crypt32, kernel32 = _windows_libraries()
    input_blob, input_buffer = _data_blob(protected)
    entropy_blob, entropy_buffer = _data_blob(_DPAPI_ENTROPY)
    output_blob = _DataBlob()
    result = crypt32.CryptUnprotectData(
        ctypes.byref(input_blob),
        None,
        ctypes.byref(entropy_blob),
        None,
        None,
        _CRYPTPROTECT_UI_FORBIDDEN,
        ctypes.byref(output_blob),
    )
    del input_buffer, entropy_buffer
    if not result:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        secret = ctypes.string_at(output_blob.pbData, output_blob.cbData)
    finally:
        kernel32.LocalFree(ctypes.cast(output_blob.pbData, ctypes.c_void_p))
    return secret.decode("utf-8")
