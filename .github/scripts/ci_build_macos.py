import argparse
import hashlib
import os
from pathlib import Path
import platform
import plistlib
import py_compile
import re
import shutil
import subprocess
import sys

APP_VERSION = "1.9.8"
TOOL_DIR_NAME = "\u96f6\u6563\u6539\u70b9\u5de5\u5177"  # 零散改点工具
APP_BUNDLE_FINAL_NAME = "\u96f6\u6563\u6539\u70b9\u6574\u7406\u5de5\u5177.app"  # 零散改点整理工具.app
DOC_MANUAL_NAME = "\u4f7f\u7528\u8bf4\u660e.md"  # 使用说明.md
BUNDLE_ID = "com.translationtools.scatteredrevision"

SECRET_RX = re.compile(
    r"(AIza[0-9A-Za-z_-]{16,}|AQ\.[0-9A-Za-z._-]{16,}|ya29\.[0-9A-Za-z._-]{16,}|"
    r"dpapi:v1:[A-Za-z0-9+/=]{12,}|posix:v1:[A-Za-z0-9+/=:]{12,})"
)


def sha256_file(p: Path) -> str:
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().lower()


def _detect_default_arch() -> str:
    m = (platform.machine() or "").strip().lower()
    return "arm64" if m in ("arm64", "aarch64") else "x64"


def _patch_info_plist(plist_path: Path) -> None:
    if not plist_path.is_file():
        raise FileNotFoundError(f"Info.plist not found in built .app bundle: {plist_path}")
    with open(plist_path, "rb") as fp:
        info = plistlib.load(fp)
    info["CFBundleDisplayName"] = "零散改点整理工具"
    info["CFBundleName"] = "零散改点整理工具"
    info["CFBundleIdentifier"] = BUNDLE_ID
    info["CFBundleShortVersionString"] = APP_VERSION
    info["CFBundleVersion"] = APP_VERSION
    info["NSHighResolutionCapable"] = True
    info["LSMinimumSystemVersion"] = "11.0"
    with open(plist_path, "wb") as fp:
        plistlib.dump(info, fp)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build standalone macOS .app bundle for Scattered Revision Tool")
    parser.add_argument("--arch", choices=("arm64", "x64"), default=_detect_default_arch())
    parser.add_argument("--check-only", action="store_true")
    args = parser.parse_args()

    arch = args.arch
    repo_root = Path(__file__).resolve().parents[2]
    tool_dir = repo_root / TOOL_DIR_NAME
    skill_dir = repo_root / ".agents" / "skills" / "scattered-revision"
    skill_scripts_dir = skill_dir / "scripts"
    skill_resources_dir = skill_dir / "resources"
    req_file = tool_dir / "requirements-build.txt"

    if not tool_dir.is_dir():
        raise FileNotFoundError(f"Tool directory not found: {tool_dir}")
    if not skill_scripts_dir.is_dir():
        raise FileNotFoundError(f"Skill scripts directory not found: {skill_scripts_dir}")

    print("[1/5] Running static syntax, eager-annotation & secret-leak checks...", flush=True)
    py_files = sorted(list(tool_dir.rglob("*.py")) + list(skill_scripts_dir.rglob("*.py")))
    for pf in py_files:
        py_compile.compile(str(pf), doraise=True)
        txt = pf.read_text(encoding="utf-8")
        if SECRET_RX.search(txt):
            raise RuntimeError(f"Hardcoded secret pattern detected in {pf.name}!")

    import builtins
    import inspect
    import symtable

    builtin_set = set(dir(builtins))
    sys.path.insert(0, str(tool_dir))
    sys.path.insert(0, str(skill_scripts_dir))
    ann_lib = None
    try:
        import annotationlib as _al  # Python 3.14+

        ann_lib = _al
    except ImportError:
        pass
    for pf in py_files:
        mod = __import__(pf.stem)
        mod_globals = set(mod.__dict__.keys()) | builtin_set
        st_top = symtable.symtable(pf.read_text(encoding="utf-8"), str(pf), "exec")

        def _check_symtable(tbl, enclosing_defs):
            local_defs = set(enclosing_defs)
            for sym in tbl.get_symbols():
                if sym.is_local() or sym.is_parameter() or sym.is_imported() or sym.is_declared_global():
                    if not sym.is_declared_global():
                        local_defs.add(sym.get_name())
            if tbl.get_type() == "function":
                for sym in tbl.get_symbols():
                    nm = sym.get_name()
                    if sym.is_global() and sym.is_referenced() and nm not in mod_globals and nm not in local_defs:
                        raise RuntimeError(
                            f"Undefined global symbol {nm!r} in {pf.name}:{tbl.get_name()} (line {tbl.get_lineno()})"
                        )
            next_defs = local_defs if tbl.get_type() == "function" else enclosing_defs
            for child in tbl.get_children():
                _check_symtable(child, next_defs)

        _check_symtable(st_top, set())
        if ann_lib is not None:
            for _, obj in inspect.getmembers(mod):
                if inspect.isfunction(obj) or inspect.isclass(obj):
                    ann_lib.get_annotations(obj, format=ann_lib.Format.VALUE)
                if inspect.isclass(obj):
                    for _, m in inspect.getmembers(obj, inspect.isfunction):
                        ann_lib.get_annotations(m, format=ann_lib.Format.VALUE)
    print(f"      Passed ({len(py_files)} Python files verified clean).", flush=True)

    if args.check_only:
        return

    if sys.platform != "darwin":
        raise RuntimeError("Full macOS .app packaging must run on a macOS runner (use --check-only on other OSes).")

    print(f"[2/5] Installing pinned PyInstaller & optional Finder DnD dependencies (macOS {arch})...", flush=True)
    env = dict(os.environ)
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"], env=env)
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", str(req_file)], env=env)

    has_tkdnd = False
    rc_dnd = subprocess.call([sys.executable, "-m", "pip", "install", "tkinterdnd2==0.4.3"], env=env)
    if rc_dnd == 0:
        chk = subprocess.call([sys.executable, "-c", "import tkinterdnd2"], env=env)
        has_tkdnd = chk == 0
    print(f"      tkinterdnd2 native Finder drop bundled: {has_tkdnd}", flush=True)

    print(f"[3/5] Building standalone macOS {arch} .app bundle with PyInstaller...", flush=True)
    work_root = repo_root / "ci_build_tmp"
    if work_root.exists():
        shutil.rmtree(work_root)
    dist_dir = work_root / "dist"
    work_dir = work_root / "work"
    spec_dir = work_root / "spec"

    add_scripts_arg = f"{skill_scripts_dir}{os.pathsep}skill_scripts"
    add_resources_arg = f"{skill_resources_dir}{os.pathsep}resources"
    pyi_cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--noupx",
        "--windowed",
        "--name",
        "ScatteredRevisionTool",
        "--osx-bundle-identifier",
        BUNDLE_ID,
        "--distpath",
        str(dist_dir),
        "--workpath",
        str(work_dir),
        "--specpath",
        str(spec_dir),
        "--paths",
        str(tool_dir),
        "--paths",
        str(skill_scripts_dir),
        "--add-data",
        add_scripts_arg,
        "--add-data",
        add_resources_arg,
        "--hidden-import",
        "docx_xml_utils",
        "--hidden-import",
        "library_index",
        "--hidden-import",
        "step1_extract_and_diff",
        "--hidden-import",
        "step1_render_draft",
        "--hidden-import",
        "step2_parse_reviewed",
        "--hidden-import",
        "step2_render_final",
        "--hidden-import",
        "ai_client",
        "--hidden-import",
        "lang_registry",
        "--hidden-import",
        "security_utils",
        "--hidden-import",
        "word_counter",
        "--hidden-import",
        "win_dnd",
        "--hidden-import",
        "fellowship_processor",
        "--hidden-import",
        "pipeline",
        "--hidden-import",
        "master_sync",
        "--hidden-import",
        "full_article_processor",
        "--hidden-import",
        "dedup_checker",
        "--hidden-import",
        "sheets_reader",
        "--hidden-import",
        "library_ingester",
        "--hidden-import",
        "proofread_ingester",
        "--hidden-import",
        "ingest_tab_ui",
        "--hidden-import",
        "updater",
    ]
    if has_tkdnd:
        pyi_cmd.extend(["--collect-all", "tkinterdnd2"])
    pyi_cmd.append(str(tool_dir / "app_gui.py"))

    subprocess.check_call(pyi_cmd, env=env)

    raw_app = dist_dir / "ScatteredRevisionTool.app"
    if not raw_app.is_dir():
        raise RuntimeError(f"Expected built macOS .app bundle not found: {raw_app}")

    final_app = dist_dir / APP_BUNDLE_FINAL_NAME
    if final_app.exists():
        shutil.rmtree(final_app)
    raw_app.rename(final_app)

    _patch_info_plist(final_app / "Contents" / "Info.plist")
    subprocess.run(["/usr/bin/xattr", "-cr", str(final_app)], check=False)
    subprocess.check_call(["/usr/bin/codesign", "--force", "--deep", "--sign", "-", str(final_app)])

    built_bin = final_app / "Contents" / "MacOS" / "ScatteredRevisionTool"
    if not built_bin.is_file():
        raise RuntimeError(f"Expected Mach-O executable missing inside .app: {built_bin}")

    print("[3.5/5] Verifying frozen macOS .app self-test (bundled config.json & modules)...", flush=True)
    self_test_log = work_root / "self_test_error.log"
    env_st = dict(env)
    env_st["SCATTERED_SELF_TEST_LOG"] = str(self_test_log)
    try:
        subprocess.run([str(built_bin), "--self-test"], env=env_st, check=True, timeout=60)
    except Exception:
        if self_test_log.is_file():
            print("=== Frozen macOS .app self-test traceback ===", flush=True)
            print(self_test_log.read_text(encoding="utf-8", errors="replace"), flush=True)
        raise
    print("      Frozen macOS .app self-test passed!", flush=True)

    print(f"[4/5] Packaging macOS {arch} portable ZIP with ditto (preserving symlinks & UTF-8 names)...", flush=True)
    art_dir = repo_root / "artifact_out"
    if art_dir.exists():
        shutil.rmtree(art_dir)
    art_dir.mkdir(parents=True, exist_ok=True)

    stage_pkg = work_root / "stage_pkg"
    stage_pkg.mkdir(parents=True, exist_ok=True)
    shutil.copytree(final_app, stage_pkg / APP_BUNDLE_FINAL_NAME, symlinks=True)
    shutil.copy2(tool_dir / "languages.json", stage_pkg / "languages.json")
    shutil.copy2(tool_dir / DOC_MANUAL_NAME, stage_pkg / DOC_MANUAL_NAME)
    shutil.copy2(tool_dir / ".env.example", stage_pkg / ".env.example")

    zip_name = f"ScatteredRevisionTool-macOS-{arch}.zip"
    zip_path = art_dir / zip_name
    subprocess.check_call(
        ["/usr/bin/ditto", "-c", "-k", "--sequesterRsrc", str(stage_pkg), str(zip_path)],
        env=env,
    )

    print(f"[5/5] Generating SHA256SUMS.mac-{arch}.txt...", flush=True)
    zip_sha = sha256_file(zip_path)
    sums_line = f"{zip_sha}  {zip_name}\n"
    (art_dir / f"SHA256SUMS.mac-{arch}.txt").write_text(sums_line, encoding="utf-8")
    print(sums_line.strip(), flush=True)


if __name__ == "__main__":
    main()
