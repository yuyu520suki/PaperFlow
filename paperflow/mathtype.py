"""Native MathType 7 adapter and DOCX embedding audit.

Import and discovery are read only. SDK connection/transforms are explicit
operations and may launch the installed MathType server. Clipboard output is
separate and requires the caller to own the desktop clipboard. Nothing here
substitutes OMML or a rendered picture for a native MathType OLE equation.
"""
from __future__ import annotations

import ctypes
import hashlib
import os
import posixpath
import struct
import uuid
import xml.etree.ElementTree as ET
import zipfile
from contextlib import contextmanager
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

MATHTYPE_PROGID = "Equation.DSMT4"
MATHTYPE_CLSID = "0002ce03-0000-0000-c000-000000000046"
LEGACY_EQUATION_PROGID = "Equation.3"
LEGACY_EQUATION_CLSID = "0002ce02-0000-0000-c000-000000000046"
_CFB_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_END = 0xFFFFFFFE
_FREE = 0xFFFFFFFF
_NS = {
    "o": "urn:schemas-microsoft-com:office:office",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


class MathTypeError(RuntimeError):
    pass


@dataclass(frozen=True)
class MathTypeInstallation:
    directory: str | None
    executable: str | None
    sdk_dll: str | None
    process_bits: int
    progid_registered: bool
    clsid: str | None
    word_templates: tuple[str, ...]

    @property
    def native_sdk_available(self) -> bool:
        return bool(self.executable and self.sdk_dll and self.progid_registered)

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["native_sdk_available"] = self.native_sdk_available
        return data


def discover_mathtype(directory: str | Path | None = None) -> MathTypeInstallation:
    """Inspect bounded install locations and COM registration; do not launch."""
    bits = struct.calcsize("P") * 8
    candidates = [Path(directory)] if directory else []
    for env in ("ProgramFiles(x86)", "ProgramFiles"):
        if os.environ.get(env):
            candidates.append(Path(os.environ[env]) / "MathType")
    clsid = None
    registered = False
    if os.name == "nt":
        import winreg
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, MATHTYPE_PROGID + r"\CLSID") as key:
                clsid = str(winreg.QueryValueEx(key, "")[0]).strip("{}").lower()
                registered = clsid == MATHTYPE_CLSID
        except OSError:
            pass
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,
                                r"CLSID\{" + MATHTYPE_CLSID + r"}\LocalServer32") as key:
                value = str(winreg.QueryValueEx(key, "")[0])
                # MathType registration normally holds an unquoted exe path.
                exe_index = value.lower().find(".exe")
                if exe_index >= 0:
                    candidates.insert(0, Path(value[:exe_index + 4].strip('"')).parent)
        except OSError:
            pass
    root = next((p for p in candidates if (p / "MathType.exe").is_file()), None)
    sdk = None
    templates: tuple[str, ...] = ()
    if root:
        possibilities = [root / "System" / str(bits) / "MT6.dll"]
        if bits == 32:
            possibilities.append(root / "System" / "MT6.dll")
        sdk = next((str(p.resolve()) for p in possibilities if p.is_file()), None)
        office = root / "Office Support" / str(bits)
        templates = tuple(str(p.resolve()) for p in sorted(office.glob("*.dotm")))
    return MathTypeInstallation(str(root.resolve()) if root else None,
                                str((root / "MathType.exe").resolve()) if root else None,
                                sdk, bits, registered, clsid, templates)


class _Rect(ctypes.Structure):
    _fields_ = [(name, ctypes.c_int32) for name in ("left", "top", "right", "bottom")]


class _Dimensions(ctypes.Structure):
    _fields_ = [("baseline", ctypes.c_int32), ("bounds", _Rect)]


class MathTypeSDK:
    """Thin native binding. The caller serializes all MathType operations.

    The installed DLL remains the authority for supported LaTeX/MathML. A
    conversion error is propagated; a caller must not silently fall back to a
    picture or OMML when native MathType is required.
    """
    LOCAL = -3
    CLIPBOARD = -2
    MTEF = 4
    TEXT = 7

    def __init__(self, installation: MathTypeInstallation | None = None):
        self.installation = installation or discover_mathtype()
        self._dll = None
        self._connected = False
        self._dll_dir = None

    def connect(self, timeout: int = 20) -> "MathTypeSDK":
        if os.name != "nt":
            raise MathTypeError("MathType native SDK requires Windows")
        if not self.installation.native_sdk_available:
            raise MathTypeError("Installed MathType server, matching-bitness MT6.dll or COM registration missing")
        if not 1 <= timeout <= 32767:
            raise ValueError("timeout must be 1..32767 seconds")
        if self._connected:
            return self
        dll_path = Path(self.installation.sdk_dll)
        if hasattr(os, "add_dll_directory"):
            self._dll_dir = os.add_dll_directory(str(dll_path.parent))
        self._dll = ctypes.WinDLL(str(dll_path))
        api = self._dll
        api.MTAPIConnect.argtypes = [ctypes.c_int16, ctypes.c_int16]
        api.MTAPIConnect.restype = ctypes.c_int32
        api.MTAPIDisconnect.argtypes = []
        api.MTAPIDisconnect.restype = ctypes.c_int32
        api.MTXFormReset.argtypes = []
        api.MTXFormReset.restype = ctypes.c_int32
        api.MTXFormGetStatus.argtypes = [ctypes.c_int16]
        api.MTXFormGetStatus.restype = ctypes.c_int32
        api.MTXFormEqn.argtypes = [
            ctypes.c_int16, ctypes.c_int16, ctypes.c_void_p, ctypes.c_int32,
            ctypes.c_int16, ctypes.c_int16, ctypes.c_void_p, ctypes.c_int32,
            ctypes.c_char_p, ctypes.POINTER(_Dimensions),
        ]
        api.MTXFormEqn.restype = ctypes.c_int32
        api.MTEquationOnClipboard.argtypes = []
        api.MTEquationOnClipboard.restype = ctypes.c_int32
        api.MTSetEqnFromLangStr.argtypes = [ctypes.c_void_p, ctypes.c_int16,
                                         ctypes.c_void_p, ctypes.c_int32]
        api.MTSetEqnFromLangStr.restype = ctypes.c_int32
        api.MTGetLangBytesFromEqn.argtypes = [ctypes.c_void_p, ctypes.c_int16,
                                            ctypes.c_void_p, ctypes.POINTER(ctypes.c_int32)]
        api.MTGetLangBytesFromEqn.restype = ctypes.c_int32
        api.MTCloseOleObject.argtypes = [ctypes.c_int32, ctypes.c_void_p]
        api.MTCloseOleObject.restype = ctypes.c_int32
        # Launch-as-needed, rather than opening an editor window proactively.
        self._check("MTAPIConnect", api.MTAPIConnect(0, timeout))
        self._connected = True
        return self

    def close(self) -> None:
        if self._connected:
            self._dll.MTAPIDisconnect()
            self._connected = False
        if self._dll_dir is not None:
            self._dll_dir.close()
            self._dll_dir = None

    def __enter__(self) -> "MathTypeSDK":
        return self.connect()

    def __exit__(self, *exc: Any) -> None:
        self.close()

    @staticmethod
    def _check(operation: str, status: int) -> None:
        if status != 0:
            raise MathTypeError(f"{operation} failed with native MathType status {status}")

    def _require_connection(self) -> Any:
        if not self._connected:
            raise MathTypeError("Call connect() explicitly before converting equations")
        return self._dll

    def to_mtef(self, source: str, source_format: str = "mathml",
                max_output_bytes: int = 1024 * 1024) -> bytes:
        """Unsupported without an OLE host; retained to fail explicitly.

        MTXFormEqn TEXT accepts MathType's translated/MTEF text, not arbitrary
        LaTeX or MathML. Passing bare MathML caused a native access violation on
        the installed 7.12.2 DLL. Import through set_ole_equation instead.
        """
        raise MathTypeError("Bare source to MTEF requires a native OLE host; use set_ole_equation then get_ole_mtef")

    def set_ole_equation(self, ole_object: Any, source: str,
                         source_format: str = "mathml", save_close: bool = True) -> None:
        """Insert source at the active MathType selection without using the clipboard.

        The caller creates/activates the native Equation.DSMT4 inline shape in
        its own Word document. Use a vendor blank equation to create a new
        formula. Existing formulas can keep their content before the insertion.
        """
        if source_format not in {"mathml", "latex"}:
            raise ValueError("source_format must be mathml or latex")
        if not source.strip() or "\0" in source:
            raise ValueError("Equation source is empty or contains a NUL")
        if source_format == "mathml":
            root = ET.fromstring(source)
            if root.tag.rsplit("}", 1)[-1] != "math":
                raise ValueError("MathML source must have a math root element")
        payload = source.encode("utf-16le")
        text = ctypes.create_string_buffer(payload + b"\0\0")
        api = self._require_connection()
        with _idispatch_pointer(ole_object) as pointer:
            self._check("MTSetEqnFromLangStr", api.MTSetEqnFromLangStr(
                pointer, 2 if source_format == "mathml" else 1,
                text, len(payload) // 2))
            if save_close:
                self._check("MTCloseOleObject", api.MTCloseOleObject(1, pointer))

    def get_ole_mtef(self, ole_object: Any) -> bytes:
        """Read native MTEF from an active native equation, without clipboard."""
        api = self._require_connection()
        with _idispatch_pointer(ole_object) as pointer:
            count = ctypes.c_int32(0)
            scratch = ctypes.create_string_buffer(16)
            self._check("MTGetLangBytesFromEqn", api.MTGetLangBytesFromEqn(
                pointer, 3, scratch, ctypes.byref(count)))
            if not 0 < count.value <= 16 * 1024 * 1024:
                raise MathTypeError(f"Invalid native MTEF length {count.value}")
            dst = ctypes.create_string_buffer(count.value)
            self._check("MTGetLangBytesFromEqn", api.MTGetLangBytesFromEqn(
                pointer, 3, dst, ctypes.byref(count)))
            return dst.raw[:count.value]

    def close_ole_equation(self, ole_object: Any, save_changes: bool = True) -> None:
        """Save/close only the caller's active native equation object."""
        api = self._require_connection()
        with _idispatch_pointer(ole_object) as pointer:
            self._check("MTCloseOleObject", api.MTCloseOleObject(1 if save_changes else 0, pointer))

    def put_mtef_on_clipboard(self, mtef: bytes) -> int:
        """Publish native OLE clipboard data; caller must own clipboard/Word.

        After this succeeds, paste as wdPasteOLEObject (0) using the caller's
        Word instance. Then check the pasted OLE ProgID and the saved DOCX with
        audit_docx_math. The returned clipboard flags alone do not verify Word.
        """
        if not mtef or mtef[0] not in {3, 4, 5}:
            raise ValueError("Expected native MTEF data")
        api = self._require_connection()
        self._check("MTXFormReset", api.MTXFormReset())
        src = ctypes.create_string_buffer(mtef)
        dims = _Dimensions()
        self._check("MTXFormEqn", api.MTXFormEqn(
            self.LOCAL, self.MTEF, src, len(mtef), self.CLIPBOARD,
            self.MTEF, None, 0, None, ctypes.byref(dims)))
        flags = api.MTEquationOnClipboard()
        if flags < 0 or not flags & (1 | 8):
            raise MathTypeError(f"Clipboard lacks a native equation OLE object (flags {flags})")
        return flags


def _release_interface(pointer: ctypes.c_void_p) -> None:
    if pointer.value:
        vtable = ctypes.cast(pointer, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
        ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)(vtable[2])(pointer)


@contextmanager
def _idispatch_pointer(ole_object: Any):
    """Bridge pywin32 to a native IDispatch* using documented COM marshaling.

    No CPython structure offsets, repr parsing, or third party COM package is
    used. Both sides remain in the caller's initialized COM apartment.
    """
    if os.name != "nt":
        raise MathTypeError("Native COM marshaling requires Windows")
    import pythoncom
    interface = getattr(ole_object, "_oleobj_", ole_object)
    py_stream = pythoncom.CreateStreamOnHGlobal()
    pythoncom.CoMarshalInterface(py_stream, pythoncom.IID_IDispatch, interface,
                                 pythoncom.MSHCTX_INPROC, pythoncom.MSHLFLAGS_NORMAL)
    length = py_stream.Seek(0, 2)
    py_stream.Seek(0, 0)
    data = py_stream.Read(length)
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
    kernel.GlobalAlloc.restype = ctypes.c_void_p
    kernel.GlobalLock.argtypes = [ctypes.c_void_p]
    kernel.GlobalLock.restype = ctypes.c_void_p
    kernel.GlobalUnlock.argtypes = [ctypes.c_void_p]
    kernel.GlobalFree.argtypes = [ctypes.c_void_p]
    kernel.GlobalFree.restype = ctypes.c_void_p
    ole32 = ctypes.OleDLL("ole32")
    ole32.CreateStreamOnHGlobal.argtypes = [ctypes.c_void_p, ctypes.c_int,
                                          ctypes.POINTER(ctypes.c_void_p)]
    ole32.CreateStreamOnHGlobal.restype = ctypes.c_long
    ole32.CoUnmarshalInterface.argtypes = [ctypes.c_void_p, ctypes.c_void_p,
                                         ctypes.POINTER(ctypes.c_void_p)]
    ole32.CoUnmarshalInterface.restype = ctypes.c_long
    handle = kernel.GlobalAlloc(2, len(data))  # GMEM_MOVEABLE
    if not handle:
        py_stream.Seek(0, 0)
        pythoncom.CoReleaseMarshalData(py_stream)
        raise MemoryError("Cannot allocate a COM marshal buffer")
    pointer = kernel.GlobalLock(handle)
    if not pointer:
        kernel.GlobalFree(handle)
        py_stream.Seek(0, 0)
        pythoncom.CoReleaseMarshalData(py_stream)
        raise MemoryError("Cannot lock a COM marshal buffer")
    ctypes.memmove(pointer, data, len(data))
    kernel.GlobalUnlock(handle)
    stream = ctypes.c_void_p()
    dispatch = ctypes.c_void_p()
    unmarshalled = False
    try:
        status = ole32.CreateStreamOnHGlobal(handle, True, ctypes.byref(stream))
        if status < 0:
            raise MathTypeError(f"CreateStreamOnHGlobal failed: HRESULT {status:#x}")
        handle = None  # stream owns HGLOBAL after successful creation
        iid = (ctypes.c_ubyte * 16).from_buffer_copy(uuid.UUID(str(pythoncom.IID_IDispatch)).bytes_le)
        status = ole32.CoUnmarshalInterface(stream, ctypes.byref(iid), ctypes.byref(dispatch))
        if status < 0 or not dispatch.value:
            raise MathTypeError(f"CoUnmarshalInterface failed: HRESULT {status:#x}")
        unmarshalled = True
        yield dispatch
    finally:
        if dispatch.value:
            _release_interface(dispatch)
        if stream.value:
            _release_interface(stream)
        elif handle:
            kernel.GlobalFree(handle)
        if not unmarshalled:
            py_stream.Seek(0, 0)
            pythoncom.CoReleaseMarshalData(py_stream)


class _CompoundFile:
    """Read only minimal CFB reader for Equation Native stream verification."""
    def __init__(self, data: bytes):
        if len(data) < 512 or data[:8] != _CFB_SIGNATURE:
            raise ValueError("Embedding is not a Compound File Binary object")
        self.data = data
        if data[28:30] != b"\xfe\xff":
            raise ValueError("Unsupported CFB byte order")
        shift = struct.unpack_from("<H", data, 30)[0]
        mini_shift = struct.unpack_from("<H", data, 32)[0]
        if shift not in {9, 12} or mini_shift != 6:
            raise ValueError("Unsupported CFB sector size")
        self.sector_size = 1 << shift
        self.mini_size = 1 << mini_shift
        self.sector_count = len(data) // self.sector_size - 1
        self.cutoff = struct.unpack_from("<I", data, 56)[0]
        difat = list(struct.unpack_from("<109I", data, 76))
        next_difat, difat_count = struct.unpack_from("<II", data, 68)
        if difat_count > self.sector_count:
            raise ValueError("Invalid CFB DIFAT sector count")
        seen_difat = set()
        for _ in range(difat_count):
            if next_difat in seen_difat:
                raise ValueError("Cyclic CFB DIFAT chain")
            seen_difat.add(next_difat)
            sector = self._sector(next_difat)
            entries = struct.unpack("<" + "I" * (self.sector_size // 4), sector)
            difat.extend(entries[:-1])
            next_difat = entries[-1]
        fat_ids = [s for s in difat if s != _FREE]
        self.fat: list[int] = []
        for sid in fat_ids:
            self.fat.extend(struct.unpack("<" + "I" * (self.sector_size // 4), self._sector(sid)))
        dir_start = struct.unpack_from("<I", data, 48)[0]
        directory = self._chain(dir_start, self.fat, self._sector)
        self.entries: dict[str, dict[str, Any]] = {}
        self.root = None
        for offset in range(0, len(directory), 128):
            entry = directory[offset:offset + 128]
            if len(entry) < 128 or entry[66] == 0:
                continue
            nlen = struct.unpack_from("<H", entry, 64)[0]
            if not 2 <= nlen <= 64 or nlen % 2:
                raise ValueError("Invalid CFB directory name")
            name = entry[:nlen - 2].decode("utf-16le")
            record = {"name": name, "type": entry[66],
                      "clsid": str(uuid.UUID(bytes_le=entry[80:96])),
                      "start": struct.unpack_from("<I", entry, 116)[0],
                      "size": struct.unpack_from("<Q", entry, 120)[0]}
            self.entries[name] = record
            if record["type"] == 5:
                self.root = record
        if not self.root:
            raise ValueError("CFB has no root storage")
        self.mini_stream = self._chain(self.root["start"], self.fat, self._sector)[:self.root["size"]]
        mini_start, mini_count = struct.unpack_from("<II", data, 60)
        mini_fat_bytes = self._chain(mini_start, self.fat, self._sector) if mini_count else b""
        self.mini_fat = list(struct.unpack("<" + "I" * (len(mini_fat_bytes) // 4), mini_fat_bytes))

    def _sector(self, sid: int) -> bytes:
        if not 0 <= sid < self.sector_count:
            raise ValueError("CFB sector index out of range")
        start = (sid + 1) * self.sector_size
        return self.data[start:start + self.sector_size]

    @staticmethod
    def _chain(start: int, fat: list[int], reader: Any) -> bytes:
        parts = []
        visited = set()
        while start not in {_END, _FREE}:
            if start in visited or not 0 <= start < len(fat):
                raise ValueError("Invalid or cyclic CFB allocation chain")
            visited.add(start)
            parts.append(reader(start))
            start = fat[start]
        return b"".join(parts)

    def stream(self, name: str) -> bytes:
        item = self.entries.get(name)
        if not item or item["type"] != 2:
            raise ValueError(f"CFB missing {name!r} stream")
        if item["size"] < self.cutoff:
            def read_mini(sid: int) -> bytes:
                start = sid * self.mini_size
                chunk = self.mini_stream[start:start + self.mini_size]
                if len(chunk) != self.mini_size:
                    raise ValueError("CFB mini sector is truncated")
                return chunk
            result = self._chain(item["start"], self.mini_fat, read_mini)
        else:
            result = self._chain(item["start"], self.fat, self._sector)
        if len(result) < item["size"]:
            raise ValueError("CFB stream is truncated")
        return result[:item["size"]]


def inspect_native_embedding(data: bytes) -> dict[str, Any]:
    """Verify native storage structure; does not prove visual/editing fidelity."""
    result: dict[str, Any] = {"valid_native_storage": False, "errors": []}
    try:
        ole = _CompoundFile(data)
        result["storage_clsid"] = ole.root["clsid"]
        result["streams"] = sorted(k for k, v in ole.entries.items() if v["type"] == 2)
        native = ole.stream("Equation Native")
        result["native_stream_bytes"] = len(native)
        if len(native) < 29:
            raise ValueError("Equation Native stream too short")
        header_size = struct.unpack_from("<I", native, 0)[0]
        result["native_header_bytes"] = header_size
        if not 8 <= header_size < len(native):
            raise ValueError("Invalid Equation Native header length")
        result["mtef_version"] = native[header_size]
        if native[header_size] not in {3, 4, 5}:
            raise ValueError("Unrecognized MTEF format version")
        if ole.root["clsid"].lower() != MATHTYPE_CLSID:
            raise ValueError("Embedding CLSID is not MathType Equation.DSMT4")
        # Word may rewrite the OLE stream header while saving unchanged MTEF.
        result["mtef_sha256"] = hashlib.sha256(native[header_size:]).hexdigest()
        result["valid_native_storage"] = True
    except (ValueError, struct.error, UnicodeError) as exc:
        result["errors"].append(str(exc))
    return result


def legacy_mtef(data: bytes) -> bytes:
    """Read the declared MTEF v3 bytes from a Microsoft Equation 3 storage."""
    ole = _CompoundFile(data)
    if ole.root["clsid"].casefold() != LEGACY_EQUATION_CLSID:
        raise ValueError("Embedding CLSID is not Microsoft Equation 3")
    native = ole.stream("Equation Native")
    if len(native) < 28:
        raise ValueError("Equation Native stream too short")
    header, version, _, length = struct.unpack_from("<HIHI", native)
    if header != 28 or version != 0x00020000 or length < 5 or header + length > len(native):
        raise ValueError("Invalid Microsoft Equation 3 native stream header or payload length")
    payload = native[header:header + length]
    if payload[0] != 3:
        raise ValueError("Microsoft Equation 3 conversion requires MTEF version 3")
    return payload


def audit_docx_math(path: str | Path) -> dict[str, Any]:
    """Count OMML and distinguish native MathType and legacy Equation 3 objects."""
    objects: list[dict[str, Any]] = []
    legacy: list[dict[str, Any]] = []
    omml_count = 0
    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())
        for part in sorted(names):
            if not part.startswith("word/") or not part.endswith(".xml"):
                continue
            root = ET.fromstring(package.read(part))
            omml_count += len(root.findall(".//m:oMath", _NS))
            rels_path = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
            rels = {}
            if rels_path in names:
                for rel in ET.fromstring(package.read(rels_path)):
                    rels[rel.attrib["Id"]] = rel.attrib
            for element in root.findall(".//o:OLEObject", _NS):
                progid = element.get("ProgID", "")
                is_legacy = progid.casefold() == LEGACY_EQUATION_PROGID.casefold()
                if not is_legacy and progid.casefold() != MATHTYPE_PROGID.casefold():
                    continue
                valid_key = "valid_legacy_storage" if is_legacy else "valid_native_storage"
                item: dict[str, Any] = {"part": part, "progid": progid,
                                        "relationship_id": element.get("{" + _NS["r"] + "}id"),
                                        valid_key: False, "errors": []}
                relation = rels.get(item["relationship_id"], {})
                if element.get("Type", "Embed") != "Embed" or relation.get("TargetMode", "Internal") != "Internal":
                    item["errors"].append("Equation object is linked rather than embedded")
                else:
                    target = relation.get("Target", "")
                    resolved = posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
                    if target.startswith("/"):
                        resolved = target.lstrip("/")
                    item["embedding_part"] = resolved
                    if not target or resolved not in names:
                        item["errors"].append("Missing embedded OLE package part")
                    elif not relation.get("Type", "").endswith("/oleObject"):
                        item["errors"].append("Relationship is not an oleObject relationship")
                    elif is_legacy:
                        try:
                            payload = legacy_mtef(package.read(resolved))
                        except (ValueError, struct.error, UnicodeError) as exc:
                            item["errors"].append(str(exc))
                        else:
                            item.update(valid_legacy_storage=True, mtef_version=3,
                                        mtef_sha256=hashlib.sha256(payload).hexdigest())
                    else:
                        item.update(inspect_native_embedding(package.read(resolved)))
                (legacy if is_legacy else objects).append(item)
    valid = sum(bool(o["valid_native_storage"]) for o in objects)
    return {"omml_equations": omml_count, "mathtype_objects": len(objects),
            "legacy_equation_objects": len(legacy), "legacy_objects": legacy,
            "native_mathtype_objects": valid, "objects": objects,
            "all_mathtype_objects_valid": valid == len(objects),
            "native_equations_present": valid > 0,
            "visual_fidelity_verified": False,
            "edit_roundtrip_verified": False}
