import ast
import atexit
import contextlib
import importlib.util
import io
import logging
import os
import queue
import shutil
import time
import subprocess
import shlex
import sys
import sysconfig
import tempfile
import threading
import tkinter as tk
from dataclasses import dataclass, field, replace
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple
import urllib.request
import zipfile

BASE_DIR = Path(__file__).resolve().parent
SELF_PATH = Path(__file__).resolve()

def _guess_default_target(base_dir: Path) -> str:
    """嘗試尋找預設要打包的腳本，避免 GUI 初始畫面為空字串"""
    for pattern in ("*.pyw", "*.py"):
        matches = sorted(base_dir.glob(pattern))
        for match in matches:
            resolved = match.resolve()
            if resolved.exists() and resolved != SELF_PATH:
                return str(resolved)
    return ""

DEFAULT_TARGET = _guess_default_target(BASE_DIR)
MAIN_ICON = BASE_DIR / "icon.png"
TOOLS_DIR = BASE_DIR / "tools"
UPX_DIR = TOOLS_DIR / "upx"
UPX_EXE = UPX_DIR / "upx.exe"
UPX_TEMP_DIR_CREATED = False
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
STARTUPINFO_HIDE = None
if os.name == "nt":
    startup = subprocess.STARTUPINFO()
    startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startup.wShowWindow = 0
    STARTUPINFO_HIDE = startup
EXCLUDED_DIRS = {"build", "Build", "dist", "__pycache__", ".git", ".venv", "env", "venv", "ENV", "node_modules"}
STDLIB_ROOT = Path(sysconfig.get_paths()["stdlib"]).resolve()
DEFAULT_EXCLUDE_MODULES = {
    "unittest",
    "test",
    "tkinter.test",
    "idlelib",
    "lib2to3",
    "pydoc",
    "pydoc_data",
    "doctest",
    "pdb",
    "turtledemo",
    "venv",
    "ensurepip",
    "wsgiref",
    "ctypes.test",
    "multiprocessing.resource_sharer",
}

PATH_GUARD_HINT = (
    "偵測到主程式會透過 __file__ / sys.argv 等路徑機制讀取附近檔案，但未發現凍結環境下的路徑判斷機制；"
    "請加入如 resolve_base_dir() 或使用 sys.executable / getattr(sys, 'frozen', False) 等保護邏輯。"
)

DEPENDENCY_ALIAS = {
    "PIL": "Pillow",
    "PIL.Image": "Pillow",
    "PIL.ImageTk": "Pillow",
    "PIL.ImageFont": "Pillow",
    "PIL.ImageFilter": "Pillow",
    "PIL._imaging": "Pillow",
    "PIL._imagingft": "Pillow",
    "cv2": "opencv-python",
    "skimage": "scikit-image",
    "ImageQt": "Pillow",
    "yaml": "PyYAML",
    "Crypto": "pycryptodome",
    "Crypto.Cipher": "pycryptodome",
}


@dataclass
class SpecOptions:
    scripts: List[str] = field(default_factory=list)
    pathex: List[str] = field(default_factory=list)
    datas: List[Tuple[str, str]] = field(default_factory=list)
    binaries: List[Tuple[str, str]] = field(default_factory=list)
    hiddenimports: List[str] = field(default_factory=list)
    hookspath: List[str] = field(default_factory=list)
    runtime_hooks: List[str] = field(default_factory=list)
    excludes: List[str] = field(default_factory=list)
    noarchive: bool = False
    console: bool = False
    upx: bool = False
    name: Optional[str] = None
    icon: Optional[str] = None


@dataclass
class SignatureConfig:
    enabled: bool = False
    signtool_path: str = ""
    cert_path: str = ""
    password: str = ""
    timestamp_url: str = ""
    description: str = ""
    digest_alg: str = "sha256"


def resolve_dependency_package(module: str) -> str:
    """依據常見別名回推實際套件名稱，避免像 PIL 這類模組安裝失敗"""
    normalized = module.strip()
    if not normalized:
        return normalized
    if normalized in DEPENDENCY_ALIAS:
        return DEPENDENCY_ALIAS[normalized]
    parts = normalized.split(".")
    while parts:
        candidate = ".".join(parts)
        if candidate in DEPENDENCY_ALIAS:
            return DEPENDENCY_ALIAS[candidate]
        parts.pop()
    return normalized


def _unique_preserve(values: Iterable[str]) -> List[str]:
    seen: Set[str] = set()
    ordered: List[str] = []
    for value in values:
        text = str(value).strip()
        if not text or text in seen:
            continue
        seen.add(text)
        ordered.append(text)
    return ordered


def _ensure_absolute_path(value: str, root: Path) -> str:
    path_obj = Path(value)
    if not path_obj.is_absolute():
        path_obj = (root / path_obj).resolve()
    return str(path_obj)


def _normalize_path_list(values: Iterable[str], root: Path) -> List[str]:
    return _unique_preserve(_ensure_absolute_path(value, root) for value in values)


def _normalize_tuple_pairs(values: Iterable[Tuple[str, str]], root: Path) -> List[Tuple[str, str]]:
    normalized: List[Tuple[str, str]] = []
    for src, dest in values:
        src_text = str(src).strip()
        if not src_text:
            continue
        dest_text = str(dest).strip() or "."
        normalized.append((_ensure_absolute_path(src_text, root), dest_text))
    return normalized


def _normalize_optional_path(value: Optional[str], root: Path) -> Optional[str]:
    if not value:
        return None
    return _ensure_absolute_path(value, root)


def _get_call_name(call: ast.Call) -> Optional[str]:
    func = call.func
    if isinstance(func, ast.Name):
        return func.id
    if isinstance(func, ast.Attribute):
        return func.attr
    return None


def _safe_literal_eval(node: ast.AST):  # type: ignore[override]
    try:
        return ast.literal_eval(node)
    except Exception:
        return None


def parse_spec_file(path: Path) -> SpecOptions:
    content = path.read_text(encoding="utf-8")
    tree = ast.parse(content, filename=str(path))
    options = SpecOptions()
    for node in ast.walk(tree):
        if not isinstance(node, ast.Assign) or not isinstance(node.value, ast.Call):
            continue
        call_name = _get_call_name(node.value)
        if call_name == "Analysis":
            if node.value.args:
                scripts = _safe_literal_eval(node.value.args[0])
                if isinstance(scripts, (list, tuple)):
                    options.scripts = [str(item) for item in scripts]
            for kw in node.value.keywords:
                value = _safe_literal_eval(kw.value)
                if value is None:
                    continue
                if kw.arg == "pathex" and isinstance(value, (list, tuple)):
                    options.pathex = [str(item) for item in value]
                elif kw.arg == "datas" and isinstance(value, (list, tuple)):
                    parsed_datas: List[Tuple[str, str]] = []
                    for entry in value:
                        if isinstance(entry, (list, tuple)) and len(entry) == 2:
                            parsed_datas.append((str(entry[0]), str(entry[1])))
                    options.datas = parsed_datas
                elif kw.arg == "binaries" and isinstance(value, (list, tuple)):
                    parsed_bins: List[Tuple[str, str]] = []
                    for entry in value:
                        if isinstance(entry, (list, tuple)) and len(entry) == 2:
                            parsed_bins.append((str(entry[0]), str(entry[1])))
                    options.binaries = parsed_bins
                elif kw.arg == "hiddenimports" and isinstance(value, (list, tuple)):
                    options.hiddenimports = [str(item) for item in value]
                elif kw.arg == "hookspath" and isinstance(value, (list, tuple)):
                    options.hookspath = [str(item) for item in value]
                elif kw.arg == "runtime_hooks" and isinstance(value, (list, tuple)):
                    options.runtime_hooks = [str(item) for item in value]
                elif kw.arg == "excludes" and isinstance(value, (list, tuple)):
                    options.excludes = [str(item) for item in value]
                elif kw.arg == "noarchive" and isinstance(value, bool):
                    options.noarchive = value
        elif call_name == "EXE":
            for kw in node.value.keywords:
                value = _safe_literal_eval(kw.value)
                if value is None:
                    continue
                if kw.arg == "name" and isinstance(value, str):
                    options.name = value
                elif kw.arg == "console" and isinstance(value, bool):
                    options.console = value
                elif kw.arg == "upx" and isinstance(value, bool):
                    options.upx = value
                elif kw.arg == "icon" and (isinstance(value, str) or value is None):
                    options.icon = value if isinstance(value, str) else None
    return options


def _repr_sequence(values: Sequence[str]) -> str:
    return "[" + ", ".join(repr(str(v)) for v in values) + "]"


def _repr_tuple_pairs(values: Sequence[Tuple[str, str]]) -> str:
    return "[" + ", ".join(repr((str(src), str(dest))) for src, dest in values) + "]"


def _quote_cmd_arg(arg: str) -> str:
    return shlex.quote(str(arg))


def _format_list_text(values: Iterable[str]) -> str:
    return "\n".join(str(value) for value in values)


def _parse_list_text(text: str) -> List[str]:
    cleaned: List[str] = []
    separators = [",", "；", "，"]
    for raw_line in text.splitlines():
        line = raw_line.strip().strip(";")
        if not line:
            continue
        normalized = line
        for sep in separators[1:]:
            normalized = normalized.replace(sep, separators[0])
        parts = [segment.strip() for segment in normalized.split(separators[0])]
        for part in parts:
            if part:
                cleaned.append(part)
    return cleaned


def _format_pairs_text(values: Iterable[Tuple[str, str]]) -> str:
    lines = []
    for src, dest in values:
        src_text = str(src)
        dest_text = str(dest or ".")
        if dest_text in {".", "./", ".\\"}:
            lines.append(src_text)
        else:
            lines.append(f"{src_text} -> {dest_text}")
    return "\n".join(lines)


def _parse_pairs_text(text: str) -> List[Tuple[str, str]]:
    pairs: List[Tuple[str, str]] = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        separator = None
        for candidate in ["->", "=>", "→", "|", ":"]:
            if candidate in line:
                separator = candidate
                break
        if separator:
            src, dest = line.split(separator, 1)
        else:
            src, dest = line, "."
        src = src.strip()
        dest = dest.strip() or "."
        if src:
            pairs.append((src, dest))
    return pairs


def sign_executable(exe_path: Path, config: SignatureConfig, logger):
    if not config.enabled:
        return
    signtool = config.signtool_path or "signtool"
    if not exe_path.exists():
        raise FileNotFoundError(f"找不到要簽章的檔案：{exe_path}")
    if not config.cert_path:
        raise ValueError("請選擇簽章用的憑證 (.pfx)")
    cmd = [signtool, "sign"]
    cmd.extend(["/f", config.cert_path])
    if config.password:
        cmd.extend(["/p", config.password])
    digest = config.digest_alg or "sha256"
    cmd.extend(["/fd", digest])
    if config.description:
        cmd.extend(["/d", config.description])
    if config.timestamp_url:
        cmd.extend(["/tr", config.timestamp_url, "/td", digest])
    cmd.append(str(exe_path))
    redact_indices = []
    if config.password:
        # 密碼在指令列中的索引為 '/p' 之後
        password_index = cmd.index(config.password)
        redact_indices.append(password_index)
    logger(f"開始簽章：{exe_path.name}")
    result = log_subprocess(cmd, logger, redact_indices=redact_indices)
    if result != 0:
        raise RuntimeError("signtool 執行失敗，請確認憑證與參數設定")
    logger("簽章完成")


class ScrollableFrame(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind(
            "<Configure>",
            lambda event: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self._window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.inner.bind("<Enter>", self._bind_mousewheel)
        self.inner.bind("<Leave>", self._unbind_mousewheel)

    def _on_mousewheel(self, event):
        delta = -1 * (event.delta // 120) if event.delta else 0
        if event.num == 4:
            delta = -1
        elif event.num == 5:
            delta = 1
        self.canvas.yview_scroll(delta, "units")

    def _bind_mousewheel(self, _event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

    def _unbind_mousewheel(self, _event):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")


def _is_sys_attribute(node: ast.AST, attr: str) -> bool:
    return (
        isinstance(node, ast.Attribute)
        and isinstance(node.value, ast.Name)
        and node.value.id == "sys"
        and node.attr == attr
    )

class _RuntimeGuardAnalyzer(ast.NodeVisitor):
    def __init__(self):
        self.uses_local_data = False
        self.guard_markers: set[str] = set()
        self.anchor_examples: list[str] = []

    def visit_FunctionDef(self, node: ast.FunctionDef):
        if node.name.lower().startswith("resolve_base_dir"):
            self.guard_markers.add("resolve_base_dir")
        self.generic_visit(node)

    def visit_Name(self, node: ast.Name):
        if node.id == "__file__":
            self.uses_local_data = True
            self._add_anchor("使用 __file__", node)
        self.generic_visit(node)

    def visit_Attribute(self, node: ast.Attribute):
        if isinstance(node.value, ast.Name) and node.value.id == "sys":
            if node.attr in {"executable", "frozen", "_MEIPASS"}:
                self.guard_markers.add(f"sys.{node.attr}")
            if node.attr == "argv":
                self.uses_local_data = True
                self._add_anchor("使用 sys.argv", node)
        self.generic_visit(node)

    def visit_Call(self, node: ast.Call):
        func = node.func
        if isinstance(func, ast.Name):
            if func.id == "resolve_base_dir":
                self.guard_markers.add("resolve_base_dir")
            elif func.id == "getattr" and len(node.args) >= 2:
                target, key = node.args[0], node.args[1]
                if isinstance(target, ast.Name) and target.id == "sys" and isinstance(key, ast.Constant):
                    if str(key.value) == "frozen":
                        self.guard_markers.add("getattr(sys,'frozen')")
            elif func.id in {"Path", "PurePath"} and node.args:
                if any(self._argument_has_anchor(arg) for arg in node.args):
                    self.uses_local_data = True
        elif isinstance(func, ast.Attribute):
            # 例如 os.path.join / os.path.dirname / Path(__file__).resolve
            if (
                isinstance(func.value, ast.Attribute)
                and isinstance(func.value.value, ast.Name)
                and func.value.value.id == "os"
                and func.value.attr == "path"
            ):
                if func.attr in {"join", "dirname", "abspath", "realpath"}:
                    if any(self._argument_has_anchor(arg) for arg in node.args):
                        self.uses_local_data = True
            elif self._argument_has_anchor(func.value):
                self.uses_local_data = True
        self.generic_visit(node)

    def visit_Subscript(self, node: ast.Subscript):
        if _is_sys_attribute(node.value, "argv"):
            self.uses_local_data = True
            self._add_anchor("索引 sys.argv", node)
        self.generic_visit(node)

    def _argument_has_anchor(self, node: ast.AST) -> bool:
        found = False
        for sub in ast.walk(node):
            if isinstance(sub, ast.Name) and sub.id == "__file__":
                self._add_anchor("參數使用 __file__", sub)
                found = True
            elif isinstance(sub, ast.Subscript) and _is_sys_attribute(sub.value, "argv"):
                self._add_anchor("參數使用 sys.argv", sub)
                found = True
            elif _is_sys_attribute(sub, "argv"):
                self._add_anchor("參數使用 sys.argv", sub)
                found = True
        return found

    def _add_anchor(self, message: str, node: ast.AST):
        lineno = getattr(node, "lineno", "?")
        self.anchor_examples.append(f"行 {lineno}: {message}")

def inspect_runtime_guard(script_path: Path):
    try:
        source = script_path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        source = script_path.read_text(encoding="utf-8", errors="ignore")
    tree = ast.parse(source, filename=str(script_path))
    analyzer = _RuntimeGuardAnalyzer()
    analyzer.file_path = script_path
    analyzer.visit(tree)
    return (
        analyzer.uses_local_data,
        bool(analyzer.guard_markers),
        sorted(analyzer.guard_markers),
        analyzer.anchor_examples,
    )

TEMP_REGISTRY = []
TEMP_LOCK = threading.Lock()


def register_temp_path(path: Path):
    path_obj = Path(path)
    with TEMP_LOCK:
        TEMP_REGISTRY.append(path_obj)


def cleanup_temp_resources():
    with TEMP_LOCK:
        while TEMP_REGISTRY:
            candidate = TEMP_REGISTRY.pop()
            if candidate.is_dir():
                _robust_rmtree(candidate)
            else:
                try:
                    candidate.unlink()
                except FileNotFoundError:
                    pass


atexit.register(cleanup_temp_resources)


def log_subprocess(cmd, logger, cwd=None, redact_indices: Optional[Iterable[int]] = None):
    """啟動子行程並將 stdout 即時寫回 GUI 日誌，便於追蹤外部工具執行情況"""
    actual_cmd = [str(part) for part in cmd]
    redact_set = set(redact_indices or [])
    display_parts = [
        ("<REDACTED>" if idx in redact_set else part)
        for idx, part in enumerate(actual_cmd)
    ]
    logger(f"執行命令：{' '.join(display_parts)}")
    process = subprocess.Popen(
        actual_cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        cwd=cwd,
        creationflags=CREATE_NO_WINDOW,
        startupinfo=STARTUPINFO_HIDE,
    )
    for line in process.stdout:
        logger(line.rstrip())
    return_code = process.wait()
    if return_code != 0:
        logger(f"命令結束，代碼 {return_code}")
    return return_code

def ensure_pyinstaller(logger):
    """確認 PyInstaller 可用，若缺少則以 pip 安裝"""
    logger("檢查 PyInstaller 是否可用...")
    cmd = [sys.executable, "-m", "PyInstaller", "--version"]
    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        creationflags=CREATE_NO_WINDOW,
        startupinfo=STARTUPINFO_HIDE,
    )
    if result.returncode == 0:
        logger(f"PyInstaller 版本：{result.stdout.strip()}")
        return
    logger("未偵測到 PyInstaller，開始安裝...")
    install_cmd = [sys.executable, "-m", "pip", "install", "pyinstaller"]
    install_result = log_subprocess(install_cmd, logger)
    if install_result != 0:
        raise RuntimeError("PyInstaller 安裝失敗")

def ensure_upx(logger):
    """檢查 UPX 是否已下載，否則自動抓取並解壓"""
    global UPX_TEMP_DIR_CREATED
    if UPX_EXE.exists():
        logger(f"已找到 UPX：{UPX_EXE}")
        return
    logger("開始下載 UPX...")
    TOOLS_DIR.mkdir(parents=True, exist_ok=True)
    tmp_dir = Path(tempfile.mkdtemp(prefix="upx_dl_", dir=BASE_DIR))
    register_temp_path(tmp_dir)
    archive_path = tmp_dir / "upx.zip"
    version = "4.2.4"
    package = f"upx-{version}-win64"
    url = f"https://github.com/upx/upx/releases/download/v{version}/{package}.zip"
    try:
        logger(f"下載 {url}")
        urllib.request.urlretrieve(url, archive_path)
        with zipfile.ZipFile(archive_path, "r") as zip_ref:
            zip_ref.extractall(tmp_dir)
        extracted = tmp_dir / package / "upx.exe"
        UPX_DIR.mkdir(parents=True, exist_ok=True)
        shutil.copy2(extracted, UPX_EXE)
        UPX_TEMP_DIR_CREATED = True
        logger(f"UPX 已安裝於 {UPX_EXE}")
    finally:
        _robust_rmtree(tmp_dir)

def discover_sources(root: Path):
    """遍歷專案路徑並回傳所有 Python 檔案路徑"""
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in EXCLUDED_DIRS]
        for filename in filenames:
            lower = filename.lower()
            if lower.endswith(".py") or lower.endswith(".pyw"):
                yield Path(dirpath) / filename

def collect_imports(root: Path):
    """分析所有來源檔案的 import 語句並彙整模組清單"""
    modules = set()
    for source in discover_sources(root):
        try:
            with source.open("r", encoding="utf-8", errors="ignore") as handler:
                tree = ast.parse(handler.read(), filename=str(source))
        except Exception:
            continue
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    modules.add(alias.name.split(".")[0])
            elif isinstance(node, ast.ImportFrom):
                if node.level == 0 and node.module:
                    modules.add(node.module.split(".")[0])
    return modules

def classify_module(name: str):
    """判斷模組來源（內建/標準庫/第三方/缺失），供依賴掃描使用"""
    if name in {"__future__", "__main__"}:
        return "ignored"
    if name in sys.builtin_module_names:
        return "builtin"
    spec = importlib.util.find_spec(name)
    if spec is None or spec.origin is None:
        return "missing"
    origin = Path(spec.origin).resolve()
    if str(origin).startswith(str(STDLIB_ROOT)):
        return "stdlib"
    return "third_party"

def scan_dependencies(root: Path):
    """取得專案的模組分類，幫助後續自動排除或提示缺失依賴"""
    modules = collect_imports(root)
    buckets = {"missing": [], "third_party": [], "stdlib": [], "builtin": []}
    for mod in sorted(modules):
        kind = classify_module(mod)
        if kind == "ignored":
            continue
        if kind not in buckets:
            buckets[kind] = []
        buckets[kind].append(mod)
    return buckets

def install_missing(modules, logger):
    """依序安裝缺少的第三方套件，並在失敗時立即拋出錯誤"""
    installed: Set[str] = set()
    for mod in modules:
        package = resolve_dependency_package(mod)
        if package in installed:
            continue
        installed.add(package)
        if package != mod:
            logger(f"正在安裝 {mod} (pip 套件 {package}) ...")
        else:
            logger(f"正在安裝 {mod} ...")
        result = log_subprocess([sys.executable, "-m", "pip", "install", package], logger)
        if result != 0:
            raise RuntimeError(f"模組 {package} 安裝失敗")

def create_runtime_hook():
    """動態建立 runtime hook，確保凍結後程式能切換到正確工作目錄"""
    hook_dir = Path(tempfile.mkdtemp(prefix="hook_", dir=BASE_DIR))
    register_temp_path(hook_dir)
    hook_path = hook_dir / "force_cwd.py"
    hook_path.write_text(
        "import os\n"
        "import sys\n"
        "def _ensure_cwd():\n"
        "    if getattr(sys, 'frozen', False):\n"
        "        base = os.path.dirname(sys.executable)\n"
        "    else:\n"
        "        base = os.path.dirname(os.path.abspath(sys.argv[0]))\n"
        "    try:\n"
        "        os.chdir(base)\n"
        "    except OSError:\n"
        "        pass\n"
        "_ensure_cwd()\n",
        encoding="utf-8",
    )
    return hook_dir, hook_path

def cleanup_artifacts(root: Path, preserve: Iterable[Path] | None = None):
    """清除 PyInstaller 中間產物與暫存檔，避免髒資料干擾下次打包"""
    preserve_set = {Path(item).resolve() for item in preserve or []}
    targets = [root / "build", root / "dist"]
    for target in targets:
        if target.exists():
            _robust_rmtree(target)
    for cache_dir in root.rglob("__pycache__"):
        _robust_rmtree(cache_dir)
    for spec_file in root.glob("*.spec"):
        if spec_file.resolve() in preserve_set:
            continue
        try:
            spec_file.unlink()
        except FileNotFoundError:
            pass
    if TOOLS_DIR.exists():
        _robust_rmtree(TOOLS_DIR)

def compress_executable(source_path: Path, dest_path: Path, logger) -> bool:
    """使用 UPX 壓縮輸出 exe 並回報壓縮成果，失敗時可回退"""
    if not source_path.exists():
        logger(f"找不到輸出檔：{source_path}")
        return False
    before_size = source_path.stat().st_size
    logger(
        f"使用 UPX 壓縮 {source_path.name}，將輸出為 {dest_path.name}，原始大小 {_format_size(before_size)}"
    )
    try:
        if dest_path.exists():
            dest_path.unlink()
        shutil.copy2(source_path, dest_path)
    except OSError as exc:
        logger(f"無法建立壓縮檔案：{exc}")
        return False
    cmd = [str(UPX_EXE), "--best", "--lzma", str(dest_path)]
    result = log_subprocess(cmd, logger)
    if result != 0:
        logger("UPX 第一次壓縮失敗，嘗試加入 --force 重新執行")
        cmd_force = [str(UPX_EXE), "--best", "--lzma", "--force", str(dest_path)]
        result = log_subprocess(cmd_force, logger)
        if result != 0:
            logger("UPX 無法在啟用 Guard CF 的 EXE 上壓縮，將刪除壓縮檔")
            dest_path.unlink(missing_ok=True)
            return False
    after_size = dest_path.stat().st_size
    logger(f"壓縮完成：{dest_path.name} {_format_size(before_size)} -> {_format_size(after_size)}")
    return True

def _repr_path(path: Path) -> str:
    return repr(str(path))

@contextlib.contextmanager
def suppress_new_console():
    """覆寫 Popen，確保子行程不會彈出新主控台視窗"""
    original_popen = subprocess.Popen
    def silent_popen(*args, **kwargs):
        creation = kwargs.get("creationflags", 0) | CREATE_NO_WINDOW
        kwargs["creationflags"] = creation
        if STARTUPINFO_HIDE and not kwargs.get("startupinfo"):
            kwargs["startupinfo"] = STARTUPINFO_HIDE
        return original_popen(*args, **kwargs)
    subprocess.Popen = silent_popen
    mp_module = None
    original_mp_popen = None
    if sys.platform.startswith("win"):
        try:
            import multiprocessing.popen_spawn_win32 as mp_module  # type: ignore
            original_mp_popen = mp_module.Popen
            def silent_mp_popen(*args, **kwargs):
                creation = kwargs.get("creationflags", 0) | CREATE_NO_WINDOW
                kwargs["creationflags"] = creation
                if STARTUPINFO_HIDE and not kwargs.get("startupinfo"):
                    kwargs["startupinfo"] = STARTUPINFO_HIDE
                return original_mp_popen(*args, **kwargs)
            mp_module.Popen = silent_mp_popen
        except ImportError:
            mp_module = None
    try:
        yield
    finally:
        subprocess.Popen = original_popen
        if mp_module and original_mp_popen:
            mp_module.Popen = original_mp_popen

class _LoggerWriter(io.TextIOBase):
    """將 stdout/stderr 轉寫入 GUI logger，模擬 file-like 介面"""

    def __init__(self, logger):
        self.logger = logger
        self._buffer = ""

    def write(self, s: str) -> int:
        if not s:
            return 0
        self._buffer += s
        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            if line.strip():
                self.logger(line.rstrip())
        return len(s)

    def flush(self):
        if self._buffer.strip():
            self.logger(self._buffer.rstrip())
        self._buffer = ""

class _QueueLogHandler(logging.Handler):
    """將 logging.Handler 轉為呼叫 GUI 的 logger"""

    def __init__(self, logger):
        super().__init__()
        self._logger = logger

    def emit(self, record):
        try:
            message = self.format(record)
        except Exception:
            self.handleError(record)
            return
        if message:
            self._logger(message)

@contextlib.contextmanager
def _capture_logging(logger):
    # logger: Callable[[str], None]，用於將訊息推入 GUI 佇列
    handler = _QueueLogHandler(logger)
    handler.setFormatter(logging.Formatter("%(message)s"))
    root_logger = logging.getLogger()
    previous_handlers = root_logger.handlers[:]
    previous_level = root_logger.level
    root_logger.handlers = [handler]
    root_logger.setLevel(logging.INFO)
    try:
        yield
    finally:
        handler.flush()
        root_logger.handlers = previous_handlers
        root_logger.setLevel(previous_level)

def run_pyinstaller_internal(args, logger):
    from PyInstaller.__main__ import run as pyinstaller_run  # type: ignore
    writer = _LoggerWriter(logger)
    exit_code = 0
    with suppress_new_console(), _capture_logging(logger), contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
        try:
            pyinstaller_run(args)
        except SystemExit as exc:
            exit_code = exc.code if isinstance(exc.code, int) else 1
        finally:
            writer.flush()
    return exit_code

def create_spec_file(
    root: Path,
    target: Path,
    hook_path: Path,
    spec_options: SpecOptions,
    icon_path: Path | None,
):
    spec_dir = Path(tempfile.mkdtemp(prefix="spec_", dir=BASE_DIR))
    register_temp_path(spec_dir)
    spec_path = spec_dir / f"{target.stem}.spec"
    scripts = spec_options.scripts or [str(target)]
    scripts = _normalize_path_list(scripts, root)
    pathex = spec_options.pathex or [str(root)]
    pathex = _normalize_path_list(pathex, root)
    datas = _normalize_tuple_pairs(spec_options.datas, root)
    binaries = _normalize_tuple_pairs(spec_options.binaries, root)
    hiddenimports = _unique_preserve(spec_options.hiddenimports)
    hookspath = _normalize_path_list([*spec_options.hookspath, str(hook_path.parent)], root)
    runtime_hooks = _normalize_path_list([*spec_options.runtime_hooks, str(hook_path)], root)
    excludes = sorted(set(spec_options.excludes))
    runtime_tmpdir = "None"
    exe_name = spec_options.name or target.stem
    use_console = spec_options.console
    use_upx = spec_options.upx
    use_noarchive = spec_options.noarchive
    icon_override = _normalize_optional_path(spec_options.icon, root)
    final_icon = icon_override or (str(icon_path) if icon_path else None)
    icon_repr = _repr_path(Path(final_icon)) if final_icon else "None"
    use_strip = False  # Windows 無 strip，可依需要調整
    spec_content = f"""# -*- mode: python ; coding: utf-8 -*-
block_cipher = None

a = Analysis(
    {_repr_sequence(scripts)},
    pathex={_repr_sequence(pathex)},
    binaries={_repr_tuple_pairs(binaries)},
    datas={_repr_tuple_pairs(datas)},
    hiddenimports={repr(hiddenimports)},
    hookspath={_repr_sequence(hookspath)},
    runtime_hooks={_repr_sequence(runtime_hooks)},
    excludes={repr(excludes)},
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive={use_noarchive},
)
pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher,
)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='{exe_name}',
    debug=False,
    bootloader_ignore_signals=False,
    strip={use_strip},
    upx={use_upx},
    upx_exclude=[],
    runtime_tmpdir={runtime_tmpdir},
    console={use_console},
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon={icon_repr},
)
"""
    spec_path.write_text(spec_content, encoding="utf-8")
    return spec_dir, spec_path

def _format_size(num_bytes: int) -> str:
    """將位元組大小轉成人類易讀格式"""
    units = ["B", "KB", "MB", "GB"]
    size = float(num_bytes)
    for unit in units:
        if size < 1024.0 or unit == units[-1]:
            return f"{size:.2f}{unit}"
        size /= 1024.0


class BuilderGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("KeySounds 打包工具")
        self.geometry("940x780")
        self.resizable(True, True)
        self.log_queue = queue.Queue()
        self.status_queue = queue.Queue()
        self._build_vars()
        self._build_widgets()
        self._refresh_spec_widgets()
        self.after(100, self._drain_log)
        self.after(300, self._auto_scan_on_start)

    def _build_vars(self):
        self.project_var = tk.StringVar(value=str(BASE_DIR))
        self.target_var = tk.StringVar(value=str(DEFAULT_TARGET))
        self.icon_var = tk.StringVar(value=str(MAIN_ICON) if MAIN_ICON.exists() else "")
        self.compress_var = tk.BooleanVar(value=True)
        self.auto_install_var = tk.BooleanVar(value=True)
        self.clean_build_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="就緒")
        self.script_source_var = tk.StringVar(value="single")
        self.script_summary_var = tk.StringVar(value="每次以主程式欄為準")
        self._spec_scripts_available = False
        self.spec_path_var = tk.StringVar(value="")
        self.spec_options = SpecOptions()
        self.spec_text_widgets: Dict[str, tk.Text] = {}
        self.spec_field_configs = [
            ("scripts", "Scripts (每行一個路徑)", 3),
            ("pathex", "Pathex (每行一個資料夾)", 3),
            ("datas", "資料檔 (來源 -> 目的)", 4),
            ("binaries", "二進位 (來源 -> 目的)", 3),
            ("hiddenimports", "Hidden Imports", 4),
            ("excludes", "排除模組", 3),
            ("runtime_hooks", "Runtime Hooks", 3),
            ("hookspath", "Hook 路徑", 3),
        ]
        self.spec_pairs_fields = {"datas", "binaries"}
        self.spec_field_modes = {
            key: ("pairs" if key in self.spec_pairs_fields else "list")
            for key, *_ in self.spec_field_configs
        }
        self.spec_name_var = tk.StringVar(value="")
        self.spec_icon_override_var = tk.StringVar(value="")
        self.spec_console_var = tk.BooleanVar(value=False)
        self.spec_upx_var = tk.BooleanVar(value=False)
        self.spec_noarchive_var = tk.BooleanVar(value=False)
        self.signature_enabled_var = tk.BooleanVar(value=False)
        self.signature_signtool_var = tk.StringVar(value="signtool")
        self.signature_cert_var = tk.StringVar(value="")
        self.signature_password_var = tk.StringVar(value="")
        self.signature_timestamp_var = tk.StringVar(value="https://timestamp.digicert.com")
        self.signature_description_var = tk.StringVar(value="")
        self.signature_digest_var = tk.StringVar(value="sha256")
        self.signature_generate_tool_var = tk.BooleanVar(value=False)
        self.signature_tool_path_var = tk.StringVar(value=str(BASE_DIR / "sign_tool.cmd"))
        self._script_radio_single: ttk.Radiobutton | None = None
        self._script_radio_spec: ttk.Radiobutton | None = None
        self._signature_tool_entry: ttk.Entry | None = None
        self.target_var.trace_add("write", lambda *_: self._refresh_script_summary())
        self.script_source_var.trace_add("write", lambda *_: self._on_script_source_change())

    def _build_widgets(self):
        padding = {"padx": 10, "pady": 5}
        frame = tk.Frame(self)
        frame.pack(fill=tk.X, **padding)
        frame.grid_columnconfigure(0, weight=0)
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_columnconfigure(2, weight=0)
        button_opts = {"width": 12}

        tk.Label(frame, text="專案根目錄").grid(row=0, column=0, sticky="w", padx=(0, 8))
        tk.Entry(frame, textvariable=self.project_var).grid(row=0, column=1, sticky="we")
        tk.Button(frame, text="瀏覽", command=self._browse_project, **button_opts).grid(row=0, column=2, padx=(8, 0))

        tk.Label(frame, text="主程式 (.py/.pyw)").grid(row=1, column=0, sticky="w", padx=(0, 8))
        tk.Entry(frame, textvariable=self.target_var).grid(row=1, column=1, sticky="we")
        tk.Button(frame, text="選擇檔案", command=self._browse_target, **button_opts).grid(row=1, column=2, padx=(8, 0))

        tk.Label(frame, text="Icon (選填)").grid(row=2, column=0, sticky="w", padx=(0, 8))
        tk.Entry(frame, textvariable=self.icon_var).grid(row=2, column=1, sticky="we")
        tk.Button(frame, text="選擇圖示", command=self._browse_icon, **button_opts).grid(row=2, column=2, padx=(8, 0))
        tk.Label(frame, text="PyInstaller spec 檔").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=(5, 0))
        tk.Entry(frame, textvariable=self.spec_path_var).grid(row=3, column=1, sticky="we", pady=(5, 0))
        spec_btns = tk.Frame(frame)
        spec_btns.grid(row=3, column=2, sticky="nsew", padx=(8, 0), pady=(5, 0))
        tk.Button(spec_btns, text="載入", command=self._browse_spec, width=12).pack(fill=tk.X, pady=1)
        tk.Button(spec_btns, text="重新讀取", command=self._reload_spec, width=12).pack(fill=tk.X, pady=1)
        tk.Label(frame, text="若載入既有 spec，表單即會帶入其內容，可在下方調整", fg="#495057").grid(
            row=4,
            column=0,
            columnspan=3,
            sticky="w",
            pady=(0, 0),
        )

        options = tk.Frame(self)
        options.pack(fill=tk.X, padx=10, pady=0)
        checkbox_items = [
            ("壓縮打包", self.compress_var),
            ("自動安裝缺少的第三方模組", self.auto_install_var),
            ("PyInstaller 清理暫存", self.clean_build_var),
        ]
        for idx, (text, var) in enumerate(checkbox_items):
            cb = tk.Checkbutton(options, text=text, variable=var, anchor="w")
            cb.grid(row=0, column=idx, sticky="w", padx=(0, 8), pady=0)
            options.grid_columnconfigure(idx, weight=1)

        advanced = tk.Frame(self)
        advanced.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 10))
        advanced.grid_columnconfigure(0, weight=3)
        advanced.grid_columnconfigure(1, weight=1)

        spec_frame = tk.LabelFrame(advanced, text="PyInstaller Spec 設定")
        spec_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        spec_frame.grid_rowconfigure(3, weight=1)

        meta = tk.Frame(spec_frame)
        meta.pack(fill=tk.X, pady=(8, 4))
        ttk.Label(meta, text="輸出名稱").grid(row=0, column=0, sticky="w")
        ttk.Entry(meta, textvariable=self.spec_name_var).grid(row=0, column=1, sticky="we", padx=(5, 15))
        ttk.Label(meta, text="Icon 覆寫").grid(row=0, column=2, sticky="w")
        ttk.Entry(meta, textvariable=self.spec_icon_override_var).grid(row=0, column=3, sticky="we", padx=(5, 0))
        meta.grid_columnconfigure(1, weight=1)
        meta.grid_columnconfigure(3, weight=1)

        checks = tk.Frame(spec_frame)
        checks.pack(fill=tk.X, pady=(0, 4))
        for idx, (label, var) in enumerate(
            [
                ("啟用 console", self.spec_console_var),
                ("允許 PyInstaller 使用 UPX", self.spec_upx_var),
                ("停用 noarchive", self.spec_noarchive_var),
            ]
        ):
            tk.Checkbutton(checks, text=label, variable=var, anchor="w").grid(
                row=0,
                column=idx,
                sticky="w",
                padx=(0, 10),
            )
        self._spec_fields_frame = ScrollableFrame(spec_frame)
        self._spec_fields_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 4))
        self._spec_fields_inner = self._spec_fields_frame.inner

        signature = tk.LabelFrame(advanced, text="簽章設定")
        signature.grid(row=0, column=1, sticky="nsew")
        ttk.Checkbutton(signature, text="啟用簽章", variable=self.signature_enabled_var).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(8, 4)
        )
        ttk.Label(signature, text="signtool 指令").grid(row=1, column=0, sticky="w")
        ttk.Entry(signature, textvariable=self.signature_signtool_var).grid(row=1, column=1, sticky="we", pady=2)
        ttk.Label(signature, text="憑證 (.pfx)").grid(row=2, column=0, sticky="w")
        cert_row = tk.Frame(signature)
        cert_row.grid(row=2, column=1, sticky="we", pady=2)
        ttk.Entry(cert_row, textvariable=self.signature_cert_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(cert_row, text="瀏覽", command=self._browse_cert).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(signature, text="憑證密碼").grid(row=3, column=0, sticky="w")
        ttk.Entry(signature, textvariable=self.signature_password_var, show="*").grid(row=3, column=1, sticky="we", pady=2)
        ttk.Label(signature, text="Timestamp URL").grid(row=4, column=0, sticky="w")
        ttk.Entry(signature, textvariable=self.signature_timestamp_var).grid(row=4, column=1, sticky="we", pady=2)
        ttk.Label(signature, text="描述").grid(row=5, column=0, sticky="w")
        ttk.Entry(signature, textvariable=self.signature_description_var).grid(row=5, column=1, sticky="we", pady=2)
        ttk.Label(signature, text="雜湊演算法").grid(row=6, column=0, sticky="w")
        digest_box = ttk.Combobox(signature, textvariable=self.signature_digest_var, values=["sha256", "sha1"], state="readonly")
        digest_box.grid(row=6, column=1, sticky="we", pady=2)
        signature.grid_columnconfigure(1, weight=1)

        buttons = tk.Frame(self)
        buttons.pack(fill=tk.X, **padding)
        tk.Button(buttons, text="開始打包", command=self._handle_build, width=30, height=2).pack(side=tk.LEFT, padx=0, pady=(0, 5))
        tk.Label(buttons, textvariable=self.status_var, fg="#2d6a4f").pack(side=tk.RIGHT)

        self.log_text = ScrolledText(self, state=tk.DISABLED, width=110, height=18)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

    def _browse_project(self):
        path = filedialog.askdirectory(initialdir=self.project_var.get() or str(BASE_DIR))
        if path:
            self.project_var.set(path)

    def _browse_target(self):
        path = filedialog.askopenfilename(
            initialdir=self.project_var.get() or str(BASE_DIR),
            filetypes=[("Python 檔案", "*.py *.pyw")],
        )
        if path:
            self.target_var.set(path)

    def _browse_icon(self):
        path = filedialog.askopenfilename(
            initialdir=self.project_var.get() or str(BASE_DIR),
            filetypes=[("Icon", "*.ico"), ("所有檔案", "*.*")],
        )
        if path:
            self.icon_var.set(path)

    def _browse_spec(self):
        path = filedialog.askopenfilename(
            initialdir=self.project_var.get() or str(BASE_DIR),
            filetypes=[("PyInstaller spec", "*.spec"), ("所有檔案", "*.*")],
        )
        if path:
            self._load_spec_from_path(Path(path))

    def _reload_spec(self):
        path_text = self.spec_path_var.get().strip()
        if not path_text:
            messagebox.showwarning("spec 檔", "請先選擇要載入的 spec 檔")
            return
        path = Path(path_text)
        if not path.exists():
            messagebox.showwarning("spec 檔", "指定的 spec 檔不存在")
            return
        self._load_spec_from_path(path)

    def _load_spec_from_path(self, path: Path):
        try:
            options = parse_spec_file(path)
        except Exception as exc:
            messagebox.showerror("載入失敗", f"無法載入 spec：{exc}")
            return
        self.spec_options = options
        self.spec_path_var.set(str(path))
        self.spec_name_var.set(options.name or path.stem)
        self.spec_icon_override_var.set(options.icon or "")
        self.spec_console_var.set(bool(options.console))
        self.spec_upx_var.set(bool(options.upx))
        self.spec_noarchive_var.set(bool(options.noarchive))
        self._populate_spec_fields(reset_meta=False)
        self._log(f"已載入 spec：{path}")

    def _browse_cert(self):
        path = filedialog.askopenfilename(
            initialdir=self.project_var.get() or str(BASE_DIR),
            filetypes=[("PFX 憑證", "*.pfx"), ("所有檔案", "*.*")],
        )
        if path:
            self.signature_cert_var.set(path)

    def _refresh_spec_widgets(self):
        if not hasattr(self, "_spec_fields_inner"):
            return
        for child in self._spec_fields_inner.winfo_children():
            child.destroy()
        self.spec_text_widgets.clear()
        current_row = 0
        for key, label, height in self.spec_field_configs:
            ttk.Label(self._spec_fields_inner, text=label).grid(
                row=current_row,
                column=0,
                sticky="w",
                pady=(4 if current_row else 0, 0),
            )
            current_row += 1
            text_widget = tk.Text(self._spec_fields_inner, height=height, wrap="word")
            text_widget.grid(row=current_row, column=0, sticky="nsew", pady=(0, 6))
            self._spec_fields_inner.grid_rowconfigure(current_row, weight=1)
            self.spec_text_widgets[key] = text_widget
            current_row += 1
        self._spec_fields_inner.grid_columnconfigure(0, weight=1)
        self._populate_spec_fields(reset_meta=True)

    def _populate_spec_fields(self, reset_meta: bool = False):
        if not self.spec_text_widgets:
            return
        for key, widget in self.spec_text_widgets.items():
            widget.configure(state=tk.NORMAL)
            widget.delete("1.0", tk.END)
            values = getattr(self.spec_options, key, [])
            if self.spec_field_modes.get(key) == "pairs":
                text = _format_pairs_text(values)
            else:
                text = _format_list_text(values)
            if text:
                widget.insert("1.0", text)
        if reset_meta:
            self.spec_name_var.set(self.spec_options.name or self.spec_name_var.get())
            self.spec_icon_override_var.set(self.spec_options.icon or self.spec_icon_override_var.get())
            self.spec_console_var.set(bool(self.spec_options.console))
            self.spec_upx_var.set(bool(self.spec_options.upx))
            self.spec_noarchive_var.set(bool(self.spec_options.noarchive))

    def _gather_spec_options(self) -> SpecOptions:
        options = replace(self.spec_options)
        for key, widget in self.spec_text_widgets.items():
            raw = widget.get("1.0", tk.END).strip()
            if not raw:
                data: Any = []
            elif self.spec_field_modes.get(key) == "pairs":
                data = _parse_pairs_text(raw)
            else:
                data = _parse_list_text(raw)
            setattr(options, key, data)
        options.name = self.spec_name_var.get().strip() or None
        icon_text = self.spec_icon_override_var.get().strip()
        options.icon = icon_text or None
        options.console = bool(self.spec_console_var.get())
        options.upx = bool(self.spec_upx_var.get())
        options.noarchive = bool(self.spec_noarchive_var.get())
        self.spec_options = replace(options)
        return options

    def _gather_signature_config(self) -> SignatureConfig:
        return SignatureConfig(
            enabled=bool(self.signature_enabled_var.get()),
            signtool_path=self.signature_signtool_var.get().strip() or "signtool",
            cert_path=self.signature_cert_var.get().strip(),
            password=self.signature_password_var.get(),
            timestamp_url=self.signature_timestamp_var.get().strip(),
            description=self.signature_description_var.get().strip(),
            digest_alg=self.signature_digest_var.get().strip() or "sha256",
        )

    def _get_preserve_spec_paths(self) -> List[Path]:
        path_text = self.spec_path_var.get().strip()
        if not path_text:
            return []
        try:
            return [Path(path_text).resolve()]
        except Exception:
            return []

    def _handle_build(self):
        if not self._preflight_guard_check():
            return
        self._run_task(self._build_task)

    def _preflight_guard_check(self) -> bool:
        target_text = self.target_var.get().strip()
        if not target_text:
            return True
        try:
            target_path = Path(target_text).expanduser().resolve()
        except Exception:
            return True
        if not target_path.exists() or not target_path.is_file():
            return True
        uses_data, has_guard, guard_markers, anchor_examples = inspect_runtime_guard(target_path)
        self._guard_check_result = (uses_data, has_guard, guard_markers, anchor_examples)
        self._guard_warning_ack = False
        if uses_data and not has_guard:
            self._log(PATH_GUARD_HINT)
            proceed = messagebox.askyesno(
                "路徑保護警告",
                f"{PATH_GUARD_HINT}\n\n缺少保護時輸出的 exe 將無法使用。\n仍要繼續打包嗎？",
            )
            if not proceed:
                self._set_status("已取消")
                self._log("已取消打包：使用者拒絕在未設置路徑保護下進行。")
                if anchor_examples:
                    for entry in anchor_examples:
                        self._log(f"關鍵片段：{entry}")
                else:
                    self._log("關鍵片段：偵測到 __file__/sys.argv，但無法取得詳細行號")
                self._log(
                    "範例：請在主程式中加入類似下列函式，並用它生成 Input/Output 路徑：\n"
                    "def resolve_base_dir():\n"
                    "    import sys, os\n"
                    "    if getattr(sys, 'frozen', False):\n"
                    "        return Path(sys.executable).resolve().parent\n"
                    "    return Path(__file__).resolve().parent\n"
                    "base_dir = resolve_base_dir()\n"
                    "input_dir = base_dir / 'Input'\n"
                    "output_dir = base_dir / 'Output'"
                )
                self._guard_check_result = None
                return False
            self._log("使用者確認在無保護情況下繼續打包。")
            self._guard_warning_ack = True
        return True

    def _run_task(self, target):
        thread = threading.Thread(target=target, daemon=True)
        thread.start()

    def _log(self, message: str):
        self.log_queue.put(message)

    def _set_status(self, text: str):
        self.status_queue.put(text)

    def _async_warning(self, title: str, message: str):
        self.after(0, lambda: messagebox.showwarning(title, message))

    def _drain_log(self):
        processed = 0
        max_per_cycle = 200
        while processed < max_per_cycle and not self.log_queue.empty():
            line = self.log_queue.get()
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, line + "\n")
            self.log_text.configure(state=tk.DISABLED)
            self.log_text.see(tk.END)
            processed += 1
        while not self.status_queue.empty():
            status = self.status_queue.get()
            self.status_var.set(status)
        delay = 50 if not self.log_queue.empty() else 100
        self.after(delay, self._drain_log)

    def _auto_scan_on_start(self):
        self._run_task(self._scan_task)

    def _scan_task(self):
        root = Path(self.project_var.get()).expanduser().resolve()
        if not root.exists():
            self._log("錯誤：專案根目錄不存在")
            self._set_status("偵測失敗")
            return
        self._set_status("偵測依賴中...")
        self._log(f"開始掃描 {root}")
        try:
            buckets = scan_dependencies(root)
        except Exception as exc:
            self._log(f"偵測失敗：{exc}")
            self._set_status("偵測失敗")
            return
        summary = (
            f"第三方：{', '.join(buckets['third_party']) or '無'}\n"
            f"缺少：{', '.join(buckets['missing']) or '無'}\n"
            f"標準庫：{len(buckets['stdlib'])} 項"
        )
        self._log(summary)
        self._set_status("偵測完成")

    def _build_task(self):
        root = Path(self.project_var.get()).expanduser().resolve()
        target = Path(self.target_var.get()).expanduser().resolve()
        if not target.exists():
            self._log("錯誤：主程式檔案不存在")
            self._set_status("失敗")
            return
        if not target.is_file():
            self._log("錯誤：請選擇有效的 .py 或 .pyw 檔案")
            self._set_status("失敗")
            return
        signature_config = self._gather_signature_config()
        preserve_specs = set(self._get_preserve_spec_paths())
        spec_config = self._gather_spec_options()
        guard_info = getattr(self, "_guard_check_result", None)
        if guard_info is None:
            guard_info = inspect_runtime_guard(target)
        uses_data, has_guard, guard_markers, anchor_examples = guard_info
        if uses_data and not has_guard and not getattr(self, "_guard_warning_ack", False):
            self._log(PATH_GUARD_HINT)
            self._set_status("已取消")
            self._log("已取消打包：缺少路徑保護且未取得使用者授權。")
            if anchor_examples:
                for entry in anchor_examples:
                    self._log(f"關鍵片段：{entry}")
            return
        if uses_data and has_guard:
            markers = ", ".join(guard_markers) or "未知"
            self._log(f"偵測到路徑保護機制：{markers}")
        if uses_data and anchor_examples:
            for entry in anchor_examples:
                self._log(f"關鍵片段：{entry}")
        self._guard_check_result = None
        self._guard_warning_ack = False
        self._set_status("打包中...")
        self._log("===== 打包流程開始 =====")
        hook_dir = None
        spec_dir = None
        try:
            buckets = scan_dependencies(root)
            missing = buckets.get("missing", [])
            third_party = buckets.get("third_party", [])
            stdlib_mods = set(buckets.get("stdlib", []))
            used_modules = set(third_party) | stdlib_mods
            self._log(f"偵測到第三方模組：{', '.join(third_party) or '無'}")
            if missing:
                self._log(f"缺少模組：{', '.join(missing)}")
                if self.auto_install_var.get():
                    install_missing(missing, self._log)
                else:
                    raise RuntimeError("存在缺少的模組，且未啟用自動安裝")
            validated_excludes = []
            for mod in spec_config.excludes:
                mod = mod.strip()
                if not mod:
                    continue
                if mod in used_modules:
                    self._log(f"排除請求 {mod} 已被專案使用，已略過")
                else:
                    validated_excludes.append(mod)
            auto_excludes = [mod for mod in DEFAULT_EXCLUDE_MODULES if mod not in used_modules]
            spec_config.excludes = sorted(set(auto_excludes + validated_excludes))
            self._log(f"將排除模組：{', '.join(spec_config.excludes) or '無'}")
            if spec_config.hiddenimports:
                spec_config.hiddenimports = _unique_preserve([*spec_config.hiddenimports, *third_party])
            else:
                spec_config.hiddenimports = sorted(set(third_party))
            if not spec_config.scripts:
                spec_config.scripts = [str(target)]
            if not spec_config.pathex:
                spec_config.pathex = [str(root)]
            spec_config.hiddenimports = _unique_preserve(spec_config.hiddenimports)
            spec_config.runtime_hooks = _unique_preserve(spec_config.runtime_hooks)
            spec_config.hookspath = _unique_preserve(spec_config.hookspath)
            self._log(f"Spec hiddenimports：{', '.join(spec_config.hiddenimports) or '無'}")
            ensure_pyinstaller(self._log)
            hook_dir, hook_path = create_runtime_hook()
            icon_path = Path(self.icon_var.get()) if self.icon_var.get() else None
            if icon_path and not icon_path.exists():
                icon_path = None
            spec_dir, spec_path = create_spec_file(root, target, hook_path, spec_config, icon_path)
            pi_args = [
                str(spec_path),
                "--noconfirm",
                "--distpath",
                str(BASE_DIR),
                "--workpath",
                str(spec_dir / "build"),
            ]
            if self.clean_build_var.get():
                pi_args.append("--clean")
            result = run_pyinstaller_internal(pi_args, self._log)
            if result != 0:
                raise RuntimeError("PyInstaller 執行失敗")
            output_exe = BASE_DIR / f"{target.stem}.exe"
            output_org = BASE_DIR / f"{target.stem}_org.exe"
            if output_org.exists():
                output_org.unlink()
            if not output_exe.exists():
                raise RuntimeError("找不到 PyInstaller 產生的 exe")
            shutil.move(output_exe, output_org)
            self._log(f"已將未壓縮版本命名為 {output_org.name}")
            final_name = f"{target.stem}.exe"
            final_dest = BASE_DIR / final_name
            if self.compress_var.get():
                ensure_upx(self._log)
                success = compress_executable(output_org, final_dest, self._log)
                if not success:
                    shutil.copy2(output_org, final_dest)
                    self._log(f"UPX 失敗，已還原未壓縮版本為 {final_name}")
            else:
                if final_dest.exists():
                    final_dest.unlink()
                shutil.move(output_org, final_dest)
                self._log(f"未勾選壓縮打包，已輸出未壓縮版本 {final_name}")
                output_org = final_dest
            if output_org.exists() and output_org.name.endswith("_org.exe"):
                output_org.unlink()
            sign_executable(final_dest, signature_config, self._log)
            self._log("===== 打包完成 =====")
            self._set_status("完成")
        except Exception as exc:
            self._log(f"發生錯誤：{exc}")
            self._set_status("失敗")
        finally:
            if hook_dir:
                _robust_rmtree(hook_dir)
            if spec_dir:
                _robust_rmtree(spec_dir)
            cleanup_artifacts(root, preserve=preserve_specs)


def _robust_rmtree(path: Path, retries: int = 3, delay: float = 0.2):
    """刪除資料夾時加入重試與微延遲，避免刪除尚未完成的情況"""
    path = Path(path)
    for attempt in range(retries):
        shutil.rmtree(path, ignore_errors=True)
        if not path.exists():
            return
        time.sleep(delay)
    if path.exists():
        shutil.rmtree(path, ignore_errors=True)

if __name__ == "__main__":
    app = BuilderGUI()
    app.mainloop()
