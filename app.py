"""Evolution edition bootstrap.

Phase 1 boots the locally vendored, proven legacy UI. The directory is therefore
portable and no longer depends on a sibling checkout.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path


HERE = Path(__file__).resolve().parent
LEGACY_APP = HERE / "legacy_app.py"


def load_legacy_app():
    if not LEGACY_APP.is_file():
        raise RuntimeError(f"找不到本地评分核心：{LEGACY_APP}")
    spec = importlib.util.spec_from_file_location("legacy_lab_report_grader", LEGACY_APP)
    if spec is None or spec.loader is None:
        raise RuntimeError("无法加载原版评分工具")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


if __name__ == "__main__":
    legacy = load_legacy_app()
    app, theme, css = legacy.create_ui()
    app.launch(
        server_name="127.0.0.1",
        server_port=7870,
        share=False,
        inbrowser=True,
        theme=theme,
        css=css,
        show_error=True,
    )
