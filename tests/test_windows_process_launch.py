"""Exercise adapter launch sites from a real Windows GUI-subsystem parent."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import textwrap
from pathlib import Path

import pytest


@pytest.mark.skipif(
    sys.platform != "win32", reason="requires Windows pythonw and console process APIs"
)
@pytest.mark.parametrize("route", ["serve", "host", "plugin"])
@pytest.mark.parametrize("fail", [False, True])
def test_pythonw_launch_preserves_protocol_errors_and_has_no_console(
    tmp_path, route, fail
):
    pythonw = Path(sys._base_executable).with_name("pythonw.exe")
    if not pythonw.is_file():
        pytest.skip("pythonw.exe is unavailable")
    receipt = tmp_path / "child.json"
    result_path = tmp_path / "result.json"
    server_stderr = tmp_path / "server-stderr.txt"
    script = tmp_path / "worker.py"
    script.write_text(
        textwrap.dedent(f"""
            import ctypes,json,os,sys
            from pathlib import Path
            info={{'pid':os.getpid(),'console':bool(ctypes.windll.kernel32.GetConsoleWindow())}}
            Path({str(receipt)!r}).write_text(json.dumps(info),encoding='utf-8')
            print('worker diagnostic',file=sys.stderr)
            if {fail!r}: raise SystemExit(7)
            if {route!r} == 'host':
                request=json.load(sys.stdin)
                print(json.dumps({{'jsonrpc':'2.0','id':request['id'],'result':{{'ready':True}}}}))
            else:
                print(json.dumps({{'success':True}}))
        """),
        encoding="utf-8",
    )
    plugin = tmp_path / "plugin"
    plugin.mkdir()
    (plugin / "plugin.json").write_text(
        json.dumps(
            {
                "name": "probe",
                "version": "1",
                "description": "Console launch probe",
                "script": "worker.py",
                "interpreter": sys.executable,
            }
        ),
        encoding="utf-8",
    )
    (plugin / "worker.py").write_text(
        script.read_text(encoding="utf-8"), encoding="utf-8"
    )
    parent = textwrap.dedent(f"""
        import json
        from pathlib import Path
        from dcc_mcp_powerpoint import host_client,plugins,server_launcher
        if {route!r} == 'serve':
            server_launcher.ServeConfig.command=lambda self:[{sys.executable!r},{str(script)!r}]
            original_call=server_launcher.subprocess.call
            errors=open({str(server_stderr)!r},'w')
            server_launcher.subprocess.call=lambda argv,**kwargs:original_call(argv,stderr=errors,**kwargs)
            result=server_launcher.serve(server_launcher.ServeConfig())
        elif {route!r} == 'host':
            host_client.find_host_binary=lambda:{sys.executable!r}
            original_run=host_client.subprocess.run
            def invoke(argv,**kwargs):
                return original_run([argv[0],{str(script)!r},*argv[1:]],**kwargs)
            host_client.subprocess.run=invoke
            result=host_client.rpc('office.host.handshake',{{}})
        else:
            result=plugins.run_plugin({str(plugin)!r})
        Path({str(result_path)!r}).write_text(json.dumps(result),encoding='utf-8')
    """)
    environment = os.environ.copy()
    environment["PYTHONPATH"] = os.pathsep.join(str(path) for path in sys.path if path)
    parent_result = subprocess.run(
        [str(pythonw), "-c", parent],
        env=environment,
        capture_output=True,
        text=True,
        timeout=30,
        check=False,
    )
    assert parent_result.returncode == 0, parent_result.stderr
    assert json.loads(receipt.read_text(encoding="utf-8"))["console"] is False
    result = json.loads(result_path.read_text(encoding="utf-8"))
    if route == "serve":
        assert result == (7 if fail else 0)
        assert "worker diagnostic" in server_stderr.read_text(encoding="utf-8")
    else:
        assert result["success"] is (not fail)
        if fail:
            assert "worker diagnostic" in (
                result.get("reason", "") + result.get("stderr", "")
            )
        elif route == "host":
            assert result["result"]["ready"] is True
