"""StartupTask wrapper for the "Open at Login" toggle (Windows / MSIX).

Mirrors the macOS ``login_item`` module. The MSIX manifest declares a
``windows.startupTask`` extension (``TaskId="TabledownStartup"``,
``Enabled="false"``); this module flips it on/off through the WinRT
``StartupTask`` API so the tray menu can offer the same toggle as macOS, and the
OS keeps the state in Settings ▸ Apps ▸ Startup — no JSON shadow.

Degrades gracefully: when ``winsdk`` is missing, or the app runs without package
identity (source/dev run, or the bare non-MSIX exe), :func:`current_state`
returns ``None`` and the menu simply omits the toggle — exactly like
``login_item.is_supported()`` returning ``False`` on older macOS.

Threading: each public operation runs its *entire* WinRT exchange — fetch the
task, toggle it, read the state back — inside one coroutine on one private,
short-lived thread (:func:`_run`). The caller, including the pystray
message-pump thread that is running a Windows message loop, only ever blocks on
``join()``. That keeps winsdk's COM apartment and asyncio setup off the caller's
thread, and — because the task object is created and used within a single
apartment lifetime — avoids sharing a WinRT proxy across threads. An unpackaged
run fails cleanly with ``OSError`` (ERROR_NOT_FOUND), which we treat as
"unsupported". (``asyncio.run`` needs a coroutine; the WinRT ``IAsyncOperation``
is awaitable but not one, so the work lives in ``async def`` helpers.)
"""
from __future__ import annotations

import asyncio
import threading

from .logger import log

# Must match Applications/Application/Extensions/.../StartupTask@TaskId in
# windows/packaging/AppxManifest.xml.
TASK_ID = "TabledownStartup"

# State names from current_state()/set_enabled() that mean the app *will* launch
# at login.
ENABLED_STATES = ("enabled", "enabled_by_policy")

try:
    from winsdk.windows.applicationmodel import StartupTask, StartupTaskState
except Exception:  # noqa: BLE001 - any import failure means the toggle is hidden
    StartupTask = None
    StartupTaskState = None

if StartupTaskState is not None:
    _STATE_NAMES = {
        StartupTaskState.DISABLED: "disabled",
        StartupTaskState.DISABLED_BY_USER: "disabled_by_user",
        StartupTaskState.DISABLED_BY_POLICY: "disabled_by_policy",
        StartupTaskState.ENABLED: "enabled",
        StartupTaskState.ENABLED_BY_POLICY: "enabled_by_policy",
    }
else:
    _STATE_NAMES = {}


def _run(fn):
    """Run ``fn()`` on a private daemon thread and return its result.

    Exceptions propagate to the caller. The thread is where winsdk gets to
    initialise its COM apartment and asyncio loop, kept off the caller.
    """
    box: dict = {}

    def runner() -> None:
        try:
            box["result"] = fn()
        except BaseException as exc:  # noqa: BLE001 - re-raised on the caller
            box["error"] = exc

    thread = threading.Thread(target=runner, name="TabledownWinRT", daemon=True)
    thread.start()
    thread.join()
    if "error" in box:
        raise box["error"]
    return box.get("result")


def _state_name(task) -> str:
    return _STATE_NAMES.get(task.state, "disabled")


async def _read_state_coro() -> str:
    task = await StartupTask.get_async(TASK_ID)
    return _state_name(task)


async def _set_enabled_coro(enabled: bool) -> str:
    task = await StartupTask.get_async(TASK_ID)
    try:
        if enabled:
            await task.request_enable_async()
        else:
            task.disable()  # synchronous WinRT call; returns void
    except Exception as exc:  # noqa: BLE001 - a refused toggle must still read back
        log(f"StartupTask {'enable' if enabled else 'disable'} failed: {exc}")
    # Read the live state back on this same thread/apartment so the menu
    # checkmark reflects what actually took (Windows can refuse an enable).
    task = await StartupTask.get_async(TASK_ID)
    return _state_name(task)


def current_state() -> str | None:
    """Current launch-at-login state name, or ``None`` when unsupported.

    ``None`` means winsdk is absent or there is no package identity (dev/source
    run, bare exe). Otherwise one of the names in :data:`_STATE_NAMES`.
    """
    if StartupTask is None:
        return None
    try:
        return _run(lambda: asyncio.run(_read_state_coro()))
    except Exception as exc:  # noqa: BLE001 - unpackaged run raises OSError
        log(f"StartupTask read failed: {exc}")
        return None


def is_supported() -> bool:
    """True when the StartupTask API is usable (winsdk present + packaged)."""
    return current_state() is not None


def is_enabled() -> bool:
    """True when the app is currently registered to launch at login."""
    return current_state() in ENABLED_STATES


def set_enabled(enabled: bool) -> str:
    """Enable/disable launch-at-login. Returns the resulting state name.

    One of ``"enabled"``, ``"enabled_by_policy"``, ``"disabled"``,
    ``"disabled_by_user"``, ``"disabled_by_policy"``, or ``"unavailable"``.

    Windows can refuse an *enable* the user turned off in Task Manager
    (``disabled_by_user``) or that an admin policy controls
    (``disabled_by_policy``). The read-back state therefore tells the caller
    whether the toggle actually took — so the menu checkmark stays truthful and
    the app can point the user at Task Manager.
    """
    if StartupTask is None:
        return "unavailable"
    try:
        return _run(lambda: asyncio.run(_set_enabled_coro(enabled)))
    except Exception as exc:  # noqa: BLE001 - a toggle must never crash the tray
        log(f"StartupTask toggle failed: {exc}")
        return "unavailable"
