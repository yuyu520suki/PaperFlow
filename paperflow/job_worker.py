"""One background CLI job; stdout is its UTF-8 log, never an MCP transport."""
from __future__ import annotations

import argparse
import _thread
from contextlib import redirect_stderr, redirect_stdout
import io
import json
import os
from pathlib import Path
import sys
import threading
import traceback

from .service import atomic_json, process_identity, utc_now, validate_cli_args


class RedactingStream(io.TextIOBase):
    """Line-buffer secrets so fragmented writes do not expose an API key."""
    def __init__(self, stream, secrets=()):
        self.stream = stream
        self.secrets = tuple(sorted({value for value in secrets if value}, key=len, reverse=True))
        self.pending = ""

    @property
    def encoding(self):
        return "utf-8"

    def _redact(self, value):
        for secret in self.secrets:
            value = value.replace(secret, "[REDACTED]")
        return value

    def write(self, value):
        self.pending += value
        if "\n" in self.pending:
            boundary = self.pending.rfind("\n") + 1
            self.stream.write(self._redact(self.pending[:boundary]))
            self.pending = self.pending[boundary:]
            self.stream.flush()
        return len(value)

    def flush(self):
        self.stream.flush()

    def finish(self):
        if self.pending:
            self.stream.write(self._redact(self.pending))
            self.pending = ""
        self.stream.flush()

    def isatty(self):
        return False


def _watch_cancel(flag, stopped):
    while not stopped.wait(0.2):
        if flag.exists():
            # Python delivers this cooperatively when the main thread next
            # handles signals. The CLI's owned Word context can clean up.
            _thread.interrupt_main()
            return


def run_job(directory: Path) -> int:
    directory = directory.resolve(strict=True)
    path = directory / "job.json"
    state = json.loads(path.read_text(encoding="utf-8"))
    if state.get("status") not in {"starting", "queued"}:
        raise ValueError("A worker can only start a new job once")
    identity = process_identity(os.getpid())
    state.update(status="running", pid=os.getpid(), creation_token=identity["creation_token"],
                 started_at=utc_now(), updated_at=utc_now())
    atomic_json(path, state)
    stopped = threading.Event()
    flag = directory / "cancel.requested"
    secrets = (os.environ.get("PAPERFLOW_AI_API_KEY", ""),)
    stdout, stderr = RedactingStream(sys.stdout, secrets), RedactingStream(sys.stderr, secrets)
    code, error = 0, None
    try:
        if flag.exists():
            raise KeyboardInterrupt
        validate_cli_args(state["args"])
        os.chdir(Path(state["workspace"]).resolve(strict=True))
        monitor = threading.Thread(target=_watch_cancel, args=(flag, stopped), daemon=True)
        monitor.start()
        with redirect_stdout(stdout), redirect_stderr(stderr):
            try:
                from .cli import main
                result = main(state["args"])
                code = result if isinstance(result, int) else 0
            except SystemExit as exited:
                code = exited.code if isinstance(exited.code, int) else 0 if exited.code is None else 1
                if exited.code is not None and not isinstance(exited.code, int):
                    print(exited.code, file=stderr)
    except KeyboardInterrupt:
        code, error = 130, "Cancellation requested; the worker returned through normal cleanup"
    except Exception as failure:
        code, error = 1, stderr._redact(str(failure))
        traceback.print_exc(file=stderr)
    finally:
        stopped.set()
        stdout.finish()
        stderr.finish()
    text = (directory / "stdout.log").read_text(encoding="utf-8") if (directory / "stdout.log").exists() else ""
    try:
        result = json.loads(text)
    except json.JSONDecodeError:
        result = None
    state.update(status="cancelled" if code == 130 else "succeeded" if code == 0 else "failed",
                 updated_at=utc_now(), finished_at=utc_now(), exit_code=code, result=result)
    if error:
        state["error"] = error
    atomic_json(path, state)
    return code


def main(argv=None):
    parser = argparse.ArgumentParser(prog="paperflow.job_worker")
    parser.add_argument("--job-dir", type=Path, required=True)
    args = parser.parse_args(argv)
    return run_job(args.job_dir)


if __name__ == "__main__":
    raise SystemExit(main())
