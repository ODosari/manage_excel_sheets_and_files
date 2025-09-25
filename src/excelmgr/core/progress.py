"""Utilities for reporting progress events from core operations."""

from __future__ import annotations

from collections.abc import Sequence
from typing import Callable


ProgressHook = Callable[[str, dict[str, object]], None]


def emit_progress(hooks: Sequence[ProgressHook], event: str, **payload: object) -> None:
    """Notify all hooks about a progress event.

    Parameters
    ----------
    hooks:
        A sequence of callback functions to be invoked. Each callback receives
        the event name followed by a payload dictionary describing the event.
    event:
        The name of the progress event being emitted.
    payload:
        Arbitrary keyword arguments describing the event context.
    """

    if not hooks:
        return
    for hook in hooks:
        hook(event, dict(payload))
