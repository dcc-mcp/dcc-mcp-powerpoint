"""Build the standalone dcc-mcp-powerpoint.exe with PyOxidizer.

Prereqs (dev-time tooling, not runtime dependencies):
- dotnet SDK 8+ (publish the C# host, self-contained win-x64)
- pyoxidizer (pip install pyoxidizer)
- Rust toolchain (PyOxidizer build)

Usage:
  python tools/build_binary.py [--host-source ../dcc-mcp-office]

Layout produced under dist/binary/:
  dcc-mcp-powerpoint.exe + lib/ (python resources + dcc-office-host.exe)
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
HOST_PROJECT = "dotnet/Office.Automation.Host/Office.Automation.Host.csproj"
VENDOR_HOST = ROOT / "vendor/lib/dcc-office-host.exe"


def publish_host(host_source: Path) -> None:
    project = host_source / HOST_PROJECT
    if not project.is_file():
        raise SystemExit(f"host project not found: {project} (pass --host-source)")
    VENDOR_HOST.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [
            "dotnet", "publish", str(project),
            "-c", "Release", "-r", "win-x64", "--self-contained",
            # Single-file host: FileManifest.add_path bundles exactly one
            # binary — no apphost/dll siblings to lose.
            "-p:PublishSingleFile=true",
            "-p:IncludeNativeLibrariesForSelfExtract=true",
            "-o", str(VENDOR_HOST.parent),
        ],
        check=True,
    )
    print(f"host published: {VENDOR_HOST}")


def build(host_source: Path | None) -> None:
    if host_source is not None:
        publish_host(host_source)
    if not VENDOR_HOST.is_file():
        raise SystemExit(f"host binary missing: {VENDOR_HOST} — run with --host-source")
    subprocess.run(["pyoxidizer", "build", "install"], cwd=str(ROOT), check=True)
    exe_path = next(ROOT.glob("build/**/dcc-mcp-powerpoint.exe"), None)
    if exe_path is None:
        raise SystemExit("pyoxidizer completed but the exe was not found under build/")
    install_dir = exe_path.parent
    dist_dir = ROOT / "dist/binary"
    dist_dir.mkdir(parents=True, exist_ok=True)
    for entry in install_dir.iterdir():
        target = dist_dir / entry.name
        if entry.is_dir():
            shutil.copytree(entry, target, dirs_exist_ok=True)
        else:
            shutil.copy2(entry, target)
    print(f"standalone build complete: {dist_dir / 'dcc-mcp-powerpoint.exe'}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Build the standalone PowerPoint adapter")
    parser.add_argument("--host-source", type=Path, default=None, help="path to the dcc-mcp-office checkout")
    args = parser.parse_args()
    build(args.host_source)


if __name__ == "__main__":
    main()
