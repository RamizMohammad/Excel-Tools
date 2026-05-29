"""
Thin async wrapper around jupyter_client.

One :class:`KernelSession` == one live IPython kernel == one notebook's
worth of persistent state. The FastAPI layer creates one per WebSocket
connection so variables survive across cells the way they do in Jupyter.
"""

from __future__ import annotations

import asyncio
from typing import AsyncIterator, Optional

from jupyter_client import AsyncKernelManager

from .kernel_bootstrap import BOOTSTRAP


class KernelSession:
    def __init__(self, kernel_name: str = "python3"):
        self.kernel_name = kernel_name
        self._km: Optional[AsyncKernelManager] = None
        self._kc = None

    # ---- lifecycle -------------------------------------------------------

    async def start(self) -> None:
        self._km = AsyncKernelManager(kernel_name=self.kernel_name)
        await self._km.start_kernel()
        self._kc = self._km.client()
        self._kc.start_channels()
        await self._kc.wait_for_ready(timeout=60)
        await self._run_silent(BOOTSTRAP)

    async def shutdown(self) -> None:
        if self._kc is not None:
            try:
                self._kc.stop_channels()
            except Exception:
                pass
        if self._km is not None:
            try:
                await self._km.shutdown_kernel(now=True)
            except Exception:
                pass

    async def interrupt(self) -> None:
        if self._km is not None:
            await self._km.interrupt_kernel()

    async def restart(self) -> None:
        if self._km is not None:
            await self._km.restart_kernel(now=True)
            await self._kc.wait_for_ready(timeout=60)
            await self._run_silent(BOOTSTRAP)

    # ---- execution -------------------------------------------------------

    async def execute(self, code: str, timeout: float = 600.0) -> AsyncIterator[dict]:
        """Run ``code`` and yield raw iopub messages belonging to it.

        The caller is responsible for translating these into the wire
        protocol the task pane expects. Iteration ends after the kernel
        reports it has gone back to ``idle`` for this request.
        """
        if self._kc is None:
            raise RuntimeError("Kernel not started")

        msg_id = self._kc.execute(code, store_history=True)

        while True:
            try:
                msg = await self._kc.get_iopub_msg(timeout=timeout)
            except asyncio.TimeoutError:
                break
            except Exception:
                break

            if msg.get("parent_header", {}).get("msg_id") != msg_id:
                continue  # belongs to some other request (e.g. a comm)

            yield msg

            if (
                msg["header"]["msg_type"] == "status"
                and msg["content"].get("execution_state") == "idle"
            ):
                break

    async def _run_silent(self, code: str, timeout: float = 60.0) -> None:
        """Execute setup code, discarding all output except surfacing errors."""
        async for msg in self.execute(code, timeout=timeout):
            if msg["header"]["msg_type"] == "error":
                tb = "\n".join(msg["content"].get("traceback", []))
                # Don't crash the session over a soft bootstrap failure; just log.
                print("[jupyxl] bootstrap warning:\n", tb)
