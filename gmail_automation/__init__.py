"""Gmail confirmation automation package."""

import platform
import sys


def _patch_windows_wmi_query() -> None:
    # SQLAlchemy imports platform.machine(), which may block indefinitely on some
    # Windows setups when platform._wmi_query hangs. Returning empty metadata keeps
    # startup deterministic and avoids modern UI launch deadlocks.
    if not sys.platform.startswith("win"):
        return
    if hasattr(platform, "_wmi_query"):
        def _disabled_wmi_query(*args, **kwargs):
            raise OSError("disabled")

        platform._wmi_query = _disabled_wmi_query  # type: ignore[attr-defined]


_patch_windows_wmi_query()

__version__ = "0.1.0"
