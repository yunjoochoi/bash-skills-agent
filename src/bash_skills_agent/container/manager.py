"""Docker container lifecycle management via CLI."""

from __future__ import annotations

import asyncio
import logging
import os
from dataclasses import dataclass

logger = logging.getLogger(__name__)


@dataclass
class ExecResult:
    exit_code: int
    stdout: str
    stderr: str


class ContainerManager:
    """Manages a per-session Docker container for code execution."""

    def __init__(
        self,
        image: str,
        workspace_dir: str,
        skills_dir: str,
        memory: str = "512m",
        network: str = "none",
    ):
        self._image = image
        self._workspace_dir = workspace_dir
        self._skills_dir = skills_dir
        self._memory = memory
        self._network = network
        self._container_id: str | None = None

    async def _ensure_started(self) -> str:
        """Lazy-start the container on first use."""
        if self._container_id:
            return self._container_id

        cmd = [
            "docker", "run", "-d",
            "--user", f"{os.getuid()}:{os.getgid()}",
            "--memory", self._memory,
            "--network", self._network,
            "-v", f"{self._workspace_dir}:/workspace:rw",
            "-v", f"{self._skills_dir}:/skills:ro",
            self._image,
        ]
        proc = await asyncio.create_subprocess_exec(
            *cmd, stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE,
        )
        stdout, stderr = await proc.communicate()

        if proc.returncode != 0:
            raise RuntimeError(f"Failed to start container: {stderr.decode().strip()}")

        self._container_id = stdout.decode().strip()[:12]
        logger.info("Container started: %s", self._container_id)
        return self._container_id

    async def exec(self, command: str, timeout: int = 120) -> ExecResult:
        """Execute a command inside the container."""
        cid = await self._ensure_started()

        cmd = ["docker", "exec", cid, "bash", "-c", command]
        proc = await asyncio.create_subprocess_exec(
            *cmd, stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE,
        )

        try:
            stdout, stderr = await asyncio.wait_for(proc.communicate(), timeout=timeout)
        except asyncio.TimeoutError:
            proc.kill()
            return ExecResult(exit_code=124, stdout="", stderr=f"Timed out after {timeout}s")

        return ExecResult(
            exit_code=proc.returncode or 0,
            stdout=stdout.decode(errors="replace"),
            stderr=stderr.decode(errors="replace"),
        )

    async def stop(self) -> None:
        """Remove the container."""
        if not self._container_id:
            return

        proc = await asyncio.create_subprocess_exec(
            "docker", "rm", "-f", self._container_id,
            stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE,
        )
        await proc.communicate()
        logger.info("Container stopped: %s", self._container_id)
        self._container_id = None

    @property
    def is_running(self) -> bool:
        return self._container_id is not None
