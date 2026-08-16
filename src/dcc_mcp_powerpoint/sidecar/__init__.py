"""Sidecar launcher for the shared C# office-host runtime."""

from .office_host import OfficeHostConfig, launch

__all__ = ["OfficeHostConfig", "launch"]
