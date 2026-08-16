# This file defines how PyOxidizer application building and packaging is
# performed (pattern learned from dcc-mcp-photoshop). See
# https://pyoxidizer.readthedocs.io/en/stable/ for details.
#
# Output layout (dist/binary/):
#   dcc-mcp-powerpoint.exe   standalone Python app (stdlib-only runtime)
#   lib/                     Python resources (filesystem-relative) + the
#                            bundled dcc-office-host.exe (our C# Open XML host)

def make_exe():
    dist = default_python_distribution()

    policy = dist.make_python_packaging_policy()
    # Filesystem-relative mode sets __file__, which dcc-mcp-core and our
    # skill packages depend on. In-memory mode breaks __file__ usage.
    policy.resources_location = "filesystem-relative:lib"

    python_config = dist.make_python_interpreter_config()
    # Standard filesystem importer: handles init_fs_encoding and __file__.
    python_config.oxidized_importer = False
    python_config.filesystem_importer = True
    python_config.module_search_paths = ["$ORIGIN/lib"]

    # Dual-purpose entry: CLI commands + skill-script dispatch.
    python_config.run_module = "dcc_mcp_powerpoint._standalone_entry"
    python_config.parse_argv = False

    exe = dist.to_python_executable(
        name="dcc-mcp-powerpoint",
        packaging_policy=policy,
        config=python_config,
    )

    # Install our package and all its dependencies via pip. Runtime stays
    # stdlib-only by policy; the heavy Office surfaces live in the bundled
    # C# host (dcc-office-host.exe).
    exe.add_python_resources(exe.pip_install(["."]))

    return exe


def make_install(exe):
    files = FileManifest()
    files.add_python_resource(".", exe)

    # Bundle the C# host next to the Python resources. tools/build_binary.py
    # publishes the host to vendor/lib/dcc-office-host.exe first; the
    # manifest strips the vendor/ prefix so the file lands in lib/.
    files.add_path("vendor/lib/dcc-office-host.exe", "vendor")

    return files


register_target("exe", make_exe)
register_target("install", make_install, depends=["exe"], default=True)
resolve_targets()
