"""
Windows power actions.
Each function calls the appropriate system command and does not return.
"""

import subprocess
import sys


def shutdown_windows() -> None:
    """Shutdown Windows immediately."""
    subprocess.run(["shutdown", "/s", "/t", "0"], check=True)


def restart_windows() -> None:
    """Restart Windows immediately."""
    subprocess.run(["shutdown", "/r", "/t", "0"], check=True)


def sleep_windows() -> None:
    """
    Put Windows to sleep.
    Uses rundll32 to invoke SetSuspendState.
    Parameters: (hibernate=0, forceCritical=1, disableWakeEvent=0)
    
    Note: If hibernate is enabled on the system, this may hibernate instead of sleep.
    To ensure sleep, run `powercfg /h off` as admin once.
    """
    subprocess.run(
        ["rundll32.exe", "powrprof.dll,SetSuspendState", "0,1,0"],
        check=True,
    )
