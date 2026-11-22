"""Official SDK MCP transport over stdio for the shared local service."""
from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

from .service import Service


def create_server(service: Service | None = None):
    from mcp.server.fastmcp import FastMCP
    backend = service or Service()
    server = FastMCP("PaperFlow", instructions=(
        "Use capabilities and command_help to discover the current business CLI. "
        "Start long operations with start_cli and poll job_status. Native Word operations are serialized by the backend. "
        "Cancellation is cooperative; keep polling cancel_requested until cleanup reaches a terminal status. "
        "read_text/write_text edit reviewable text plans; write_text refuses an existing file unless overwrite is explicit."
    ))

    @server.tool(name="capabilities")
    def capabilities() -> dict[str, Any]:
        """List every current business command and its argument metadata."""
        return backend.capabilities()

    @server.tool(name="command_help")
    def command_help(command: str) -> str:
        """Return full help for one business command."""
        return backend.get_command_help(command)

    @server.tool(name="start_cli")
    def start_cli(args: list[str], env_overrides: dict[str, str] | None = None) -> dict[str, Any]:
        """Validate CLI arguments and launch a background job without a shell.

        Pass AI endpoint/model/key/timeout through whitelisted environment
        overrides. API keys are not stored in job metadata or command lines.
        """
        return backend.start_cli(args, env_overrides)

    @server.tool(name="job_status")
    def job_status(job_id: str, tail_chars: int = 12000) -> dict[str, Any]:
        """Read persisted status, verified worker liveness, result and log tails."""
        return backend.job_status(job_id, tail_chars)

    @server.tool(name="list_jobs")
    def list_jobs(limit: int = 50) -> list[dict[str, Any]]:
        """List recent jobs shared with the local GUI and other MCP sessions."""
        return backend.list_jobs(limit)

    @server.tool(name="cancel_job")
    def cancel_job(job_id: str) -> dict[str, Any]:
        """Request cooperative cancellation; this never force-kills Word."""
        return backend.cancel_job(job_id)

    @server.tool(name="read_text")
    def read_text(path: str, max_chars: int = 200000) -> dict[str, Any]:
        """Read a UTF-8 text/JSON plan with a hash, relative to the workspace or by full path."""
        return backend.read_text(path, max_chars)

    @server.tool(name="write_text")
    def write_text(path: str, text: str, overwrite: bool = False) -> dict[str, Any]:
        """Write a UTF-8 text/JSON plan; existing files are protected by default."""
        return backend.write_text(path, text, overwrite)

    return server


def run_server(workspace: Path | None = None, state_dir: Path | None = None):
    create_server(Service(workspace, state_dir)).run(transport="stdio")


def main(argv=None):
    parser = argparse.ArgumentParser(prog="paperflow.mcp_server")
    parser.add_argument("--workspace", type=Path)
    parser.add_argument("--state-dir", type=Path)
    args = parser.parse_args(argv)
    run_server(args.workspace, args.state_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
