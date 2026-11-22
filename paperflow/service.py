"""Shared command/job service for the local GUI and MCP server.

Business commands use the same argparse contract as the CLI. Jobs run in
separate Python workers, never through a shell. Only job-owned state files are
updated here; CLI operations retain their own source/output protections.
"""
from __future__ import annotations

import argparse
from contextlib import redirect_stderr, redirect_stdout
from datetime import datetime, timezone
import hashlib
import io
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import threading
from typing import Any
import uuid

from . import __version__


PROJECT_ROOT = Path(__file__).resolve().parents[1]
TERMINAL_STATUSES = {"succeeded", "failed", "cancelled"}
AI_ENV_KEYS = frozenset({"PAPERFLOW_AI_BASE_URL", "PAPERFLOW_AI_MODEL", "PAPERFLOW_AI_API_KEY", "PAPERFLOW_AI_TIMEOUT"})
_JOB_ID = re.compile(r"^[0-9a-f]{32}$")
_EXCLUDED_COMMANDS = {"gui", "mcp"}


def console_python():
    executable = Path(sys.executable)
    if os.name == "nt" and executable.name.lower() == "pythonw.exe":
        return str(executable.with_name("python.exe"))
    return str(executable)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def atomic_json(path: Path, value: dict[str, Any]) -> None:
    """Replace only an owned metadata file, keeping readers off partial JSON."""
    temporary = path.with_name("." + path.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        with temporary.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(value, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def _read_json(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise ValueError(f"Invalid job metadata object: {path}")
    return value


def process_identity(pid: int) -> dict[str, Any]:
    """Inspect the live process, including its birth token where available."""
    if not isinstance(pid, int) or isinstance(pid, bool) or pid <= 0:
        return {"alive": False, "creation_token": None, "verified": True}
    if os.name == "nt":
        import ctypes
        from ctypes import wintypes
        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        kernel.OpenProcess.restype = wintypes.HANDLE
        kernel.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
        kernel.WaitForSingleObject.restype = wintypes.DWORD
        kernel.GetProcessTimes.argtypes = [wintypes.HANDLE, *([ctypes.POINTER(wintypes.FILETIME)] * 4)]
        kernel.GetProcessTimes.restype = wintypes.BOOL
        kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        kernel.CloseHandle.restype = wintypes.BOOL
        handle = kernel.OpenProcess(0x00100000 | 0x1000, False, pid)
        if not handle:
            error = ctypes.get_last_error()
            return {"alive": None if error == 5 else False, "creation_token": None,
                    "verified": error != 5, "error": error}
        try:
            wait = kernel.WaitForSingleObject(handle, 0)
            alive = True if wait == 258 else False if wait == 0 else None
            times = [wintypes.FILETIME() for _ in range(4)]
            token = None
            if kernel.GetProcessTimes(handle, *(ctypes.byref(value) for value in times)):
                token = f"{times[0].dwHighDateTime:08x}{times[0].dwLowDateTime:08x}"
            return {"alive": alive, "creation_token": token, "verified": alive is not None}
        finally:
            kernel.CloseHandle(handle)
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return {"alive": False, "creation_token": None, "verified": True}
    except PermissionError:
        return {"alive": None, "creation_token": None, "verified": False}
    token = None
    try:
        stat = Path(f"/proc/{pid}/stat").read_text(encoding="utf-8")
        token = stat[stat.rfind(")") + 2:].split()[19]
    except (OSError, IndexError):
        pass
    return {"alive": True, "creation_token": token, "verified": True}


def _parser():
    from . import cli
    builder = getattr(cli, "build_parser", None)
    if builder is None:
        raise RuntimeError("The CLI must expose build_parser() before starting the shared service")
    return builder()


def _command_parsers(parser):
    for action in parser._actions:
        if isinstance(action, argparse._SubParsersAction):
            return action, action.choices
    raise RuntimeError("The CLI parser does not expose business subcommands")


def validate_cli_args(args: list[str]):
    if not isinstance(args, list) or not args or any(not isinstance(value, str) or "\0" in value for value in args):
        raise ValueError("args must be a nonempty list of strings without NUL characters")
    parser = _parser()
    output = io.StringIO()
    try:
        with redirect_stdout(output), redirect_stderr(output):
            namespace = parser.parse_args(args)
    except SystemExit as error:
        detail = output.getvalue().strip()
        raise ValueError(detail or f"CLI argument parsing ended with code {error.code}") from None
    command = getattr(namespace, "command", None)
    if not command or command in _EXCLUDED_COMMANDS:
        raise ValueError("Start a business command; GUI/MCP hosting cannot be launched as a job")
    return namespace


def _json_safe(value):
    if value is argparse.SUPPRESS:
        return None
    try:
        json.dumps(value)
        return value
    except (TypeError, ValueError):
        return str(value)


def _tail(path: Path, chars: int) -> str:
    if chars == 0 or not path.is_file():
        return ""
    with path.open("rb") as stream:
        stream.seek(max(0, path.stat().st_size - chars * 4))
        return stream.read().decode("utf-8", errors="replace")[-chars:]


class Service:
    def __init__(self, workspace: Path | None = None, state_dir: Path | None = None):
        self.workspace = Path(workspace or PROJECT_ROOT).expanduser().resolve(strict=True)
        if not self.workspace.is_dir():
            raise ValueError("workspace must be an existing directory")
        self.state_dir = Path(state_dir).expanduser().resolve() if state_dir is not None else self.workspace / ".paperflow" / "jobs"
        self.state_dir.mkdir(parents=True, exist_ok=True)
        self.state_dir = self.state_dir.resolve(strict=True)
        self._processes: dict[str, subprocess.Popen] = {}
        self._lock = threading.RLock()

    def capabilities(self) -> dict[str, Any]:
        parser = _parser()
        subparsers, commands = _command_parsers(parser)
        descriptions = {action.dest: action.help for action in subparsers._choices_actions}
        result = []
        for name, command in commands.items():
            if name in _EXCLUDED_COMMANDS:
                continue
            arguments = []
            for action in command._actions:
                if isinstance(action, argparse._HelpAction):
                    continue
                arguments.append({"name": action.dest, "flags": list(action.option_strings),
                                  "positional": not bool(action.option_strings), "required": bool(action.required),
                                  "nargs": action.nargs, "default": _json_safe(action.default),
                                  "choices": list(action.choices) if action.choices is not None else None,
                                  "type": getattr(action.type, "__name__", "str"),
                                  "action": action.__class__.__name__, "help": action.help or ""})
            result.append({"name": name, "help": descriptions.get(name, "") or "",
                           "description": command.description or "", "usage": command.format_usage().strip(),
                           "arguments": arguments})
        return {"version": __version__, "workspace": str(self.workspace), "commands": result}

    def get_command_help(self, command: str) -> str:
        _, commands = _command_parsers(_parser())
        if command not in commands or command in _EXCLUDED_COMMANDS:
            raise ValueError(f"Unknown business command: {command}")
        return commands[command].format_help()

    def _job_dir(self, job_id: str) -> Path:
        if not isinstance(job_id, str) or not _JOB_ID.fullmatch(job_id):
            raise ValueError("job_id must be the ID returned by start_cli")
        directory = (self.state_dir / job_id).resolve()
        if directory.parent != self.state_dir:
            raise ValueError("Job directory resolves outside the job store")
        if not directory.is_dir():
            raise FileNotFoundError(f"Unknown job: {job_id}")
        return directory

    def start_cli(self, args: list[str], env_overrides: dict | None = None) -> dict[str, Any]:
        namespace = validate_cli_args(args)
        overrides = dict(env_overrides or {})
        if any(key not in AI_ENV_KEYS for key in overrides):
            raise ValueError("Only PAPERFLOW_AI_BASE_URL/MODEL/API_KEY/TIMEOUT environment overrides are accepted")
        if any(not isinstance(value, str) or "\0" in value for value in overrides.values()):
            raise ValueError("Environment override values must be strings without NUL characters")
        job_id = uuid.uuid4().hex
        directory = self.state_dir / job_id
        directory.mkdir(exist_ok=False)
        created = utc_now()
        state = {"id": job_id, "job_id": job_id, "status": "starting", "command": namespace.command,
                 "args": list(args), "workspace": str(self.workspace), "created_at": created, "updated_at": created,
                 "pid": None, "creation_token": None, "exit_code": None, "result": None,
                 "env_override_names": sorted(overrides)}
        atomic_json(directory / "job.json", state)
        environment = os.environ.copy()
        environment.update(overrides)
        environment["PYTHONNOUSERSITE"] = "1"
        environment["PYTHONUNBUFFERED"] = "1"
        environment["PYTHONPATH"] = str(PROJECT_ROOT) + (os.pathsep + environment["PYTHONPATH"] if environment.get("PYTHONPATH") else "")
        command = [console_python(), "-s", "-X", "utf8", "-m", "paperflow.job_worker", "--job-dir", str(directory)]
        try:
            with (directory / "stdout.log").open("ab", buffering=0) as stdout, (directory / "stderr.log").open("ab", buffering=0) as stderr:
                process = subprocess.Popen(command, cwd=PROJECT_ROOT, env=environment,
                                           stdin=subprocess.DEVNULL, stdout=stdout, stderr=stderr, shell=False,
                                           creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0)
            identity = process_identity(process.pid)
            atomic_json(directory / "process.json", {"pid": process.pid, "creation_token": identity["creation_token"],
                                                      "launched_at": utc_now()})
            with self._lock:
                self._processes[job_id] = process
        except Exception as error:
            state.update(status="failed", updated_at=utc_now(), finished_at=utc_now(), exit_code=-1,
                         error=f"Could not start worker: {error}")
            atomic_json(directory / "job.json", state)
            raise
        return {"id": job_id, "job_id": job_id, "status": "starting", "pid": process.pid}

    def job_status(self, job_id: str, tail_chars: int = 12000) -> dict[str, Any]:
        if not isinstance(tail_chars, int) or isinstance(tail_chars, bool) or not 0 <= tail_chars <= 2_000_000:
            raise ValueError("tail_chars must be an integer from 0 through 2000000")
        directory = self._job_dir(job_id)
        state = _read_json(directory / "job.json")
        launch = _read_json(directory / "process.json") if (directory / "process.json").is_file() else {}
        pid = state.get("pid") or launch.get("pid")
        expected_token = state.get("creation_token") or launch.get("creation_token")
        identity = process_identity(pid) if pid else {"alive": None, "verified": False, "creation_token": None}
        reused = bool(expected_token and identity.get("creation_token") and expected_token != identity["creation_token"])
        alive = False if reused else identity["alive"]
        with self._lock:
            process = self._processes.get(job_id)
            if process is not None and process.poll() is not None:
                self._processes.pop(job_id, None)
                alive = False
        if state.get("status") not in TERMINAL_STATUSES:
            stale_launch = not pid and (datetime.now(timezone.utc) - datetime.fromisoformat(state["created_at"])).total_seconds() > 10
            if alive is False or stale_launch:
                # A worker can publish its terminal state between the first
                # read and process exit; reread before recording a crash.
                state = _read_json(directory / "job.json")
                if state.get("status") not in TERMINAL_STATUSES:
                    state.update(status="failed", updated_at=utc_now(), finished_at=utc_now(),
                                 error="Worker PID was reused" if reused else "Worker exited without a terminal status",
                                 exit_code=state.get("exit_code") if state.get("exit_code") is not None else -1)
                    atomic_json(directory / "job.json", state)
        state = dict(state)
        cancellation = (directory / "cancel.requested").exists()
        state.update(pid=pid, worker_alive=alive, liveness_verified=bool(identity["verified"]),
                     pid_identity_matches=bool(expected_token and expected_token == identity.get("creation_token")),
                     cancel_requested=cancellation, job_dir=str(directory),
                     stdout=_tail(directory / "stdout.log", tail_chars), stderr=_tail(directory / "stderr.log", tail_chars))
        if cancellation and state.get("status") not in TERMINAL_STATUSES:
            state["status"] = "cancel_requested"
        return state

    def list_jobs(self, limit: int = 50) -> list[dict[str, Any]]:
        if not isinstance(limit, int) or isinstance(limit, bool) or not 1 <= limit <= 1000:
            raise ValueError("limit must be an integer from 1 through 1000")
        directories = sorted((path for path in self.state_dir.iterdir() if path.is_dir() and _JOB_ID.fullmatch(path.name)),
                             key=lambda path: path.stat().st_mtime, reverse=True)
        result = []
        for directory in directories:
            if not (directory / "job.json").is_file():
                continue
            try:
                job = self.job_status(directory.name, tail_chars=0)
                result.append({key: value for key, value in job.items() if key not in {"result", "stdout", "stderr"}})
            except (ValueError, OSError) as error:
                result.append({"id": directory.name, "job_id": directory.name, "status": "failed", "error": str(error)})
            if len(result) == limit:
                break
        return result

    def cancel_job(self, job_id: str) -> dict[str, Any]:
        state = self.job_status(job_id)
        if state["status"] in TERMINAL_STATUSES:
            return state
        flag = self._job_dir(job_id) / "cancel.requested"
        try:
            with flag.open("x", encoding="utf-8") as stream:
                stream.write(utc_now() + "\n")
        except FileExistsError:
            pass
        return self.job_status(job_id)

    def read_text(self, path: str, max_chars: int = 200000) -> dict[str, Any]:
        if not isinstance(max_chars, int) or isinstance(max_chars, bool) or max_chars < 0:
            raise ValueError("max_chars must be a nonnegative integer")
        target = Path(path).expanduser()
        target = (self.workspace / target).resolve() if not target.is_absolute() else target.resolve()
        contents = target.read_text(encoding="utf-8-sig")
        return {"path": str(target), "text": contents[:max_chars], "total_chars": len(contents),
                "truncated": len(contents) > max_chars, "sha256": hashlib.sha256(target.read_bytes()).hexdigest()}

    def write_text(self, path: str, text: str, overwrite: bool = False) -> dict[str, Any]:
        if not isinstance(text, str):
            raise ValueError("text must be a string")
        target = Path(path).expanduser()
        target = (self.workspace / target).resolve() if not target.is_absolute() else target.resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        existed = target.exists()
        with target.open("w" if overwrite else "x", encoding="utf-8", newline="") as stream:
            stream.write(text)
        data = target.read_bytes()
        return {"path": str(target), "bytes_written": len(data), "sha256": hashlib.sha256(data).hexdigest(),
                "overwritten": bool(existed and overwrite)}
