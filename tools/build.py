"""可重现的 PyInstaller 便携目录、许可证和 ZIP 打包流程。"""
import importlib.metadata as metadata
import json
import shutil
import subprocess
import sys
import urllib.request
import zipfile
from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
OUTPUTS = ROOT / "outputs"
WORK = ROOT / "work"


def make_icon(path: Path):
    image = Image.new("RGBA", (256, 256), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((8, 8, 248, 248), 52, fill="#2563eb")
    draw.rounded_rectangle((48, 54, 200, 166), 10, fill="white")
    draw.line((77, 87, 166, 87), fill="#93c5fd", width=12)
    draw.line((77, 115, 129, 115), fill="#93c5fd", width=12)
    draw.ellipse((135, 131, 195, 191), fill="white", outline="#1e40af", width=10)
    draw.line((184, 181, 212, 213), fill="#1e40af", width=15)
    image.save(path, sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])


def qt_license_texts():
    directory = ROOT / "licenses" / "Qt"
    directory.mkdir(parents=True, exist_ok=True)
    for name in ("LGPL-3.0-only.txt", "GPL-3.0-only.txt"):
        target = directory / name
        if target.exists():
            continue
        url = f"https://raw.githubusercontent.com/qt/qtbase/v6.8.3/LICENSES/{name}"
        with urllib.request.urlopen(url, timeout=30) as response:
            data = response.read()
        if len(data) < 1000 or b"GNU" not in data:
            raise ValueError(f"Invalid Qt license download: {name}")
        target.write_bytes(data)


def copy_licenses(bundle: Path):
    target_root = bundle / "THIRD_PARTY_LICENSES"
    target_root.mkdir(exist_ok=True)
    for distribution in metadata.distributions():
        name = distribution.metadata["Name"]
        if name.lower() in {"pip", "setuptools"}:
            continue
        for relative in distribution.files or []:
            if not any(part in str(relative).lower() for part in ("license", "copying", "copyright")):
                continue
            source = Path(distribution.locate_file(relative))
            if source.is_file() and source.suffix.lower() not in {".py", ".pyc", ".pyd"}:
                # 保留路径以防同一依赖中不同组件的 LICENSE 同名。
                target = target_root / name / str(relative).replace("../", "_parent_/")
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source, target)
    shutil.copytree(ROOT / "licenses" / "Qt", target_root / "Qt", dirs_exist_ok=True)
    python_license = Path(sys.base_prefix) / "LICENSE.txt"
    if python_license.exists():
        shutil.copy2(python_license, target_root / "Python-LICENSE.txt")


def source_archive(path: Path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name in ("main.py", "pyproject.toml", "requirements-lock.txt", "README.md", "LICENSE",
                     "THIRD_PARTY_NOTICES.md", "build.ps1", ".gitignore", ".gitattributes"):
            archive.write(ROOT / name, "PPTSearch-source/" + name)
        for directory in ("ppt_search", "tests", "tools", "licenses", "docs"):
            for file in (ROOT / directory).rglob("*"):
                if file.is_file() and "__pycache__" not in file.parts:
                    archive.write(file, "PPTSearch-source/" + file.relative_to(ROOT).as_posix())


def main():
    OUTPUTS.mkdir(exist_ok=True)
    WORK.mkdir(exist_ok=True)
    icon = WORK / "ppt-search.ico"
    make_icon(icon)
    qt_license_texts()
    subprocess.run([sys.executable, "-m", "PyInstaller", "--name", "PPTSearch", "--noconsole",
                    "--onedir", "--noconfirm", "--clean", "--noupx", "--icon", str(icon),
                    "--distpath", str(WORK / "portable"), "--workpath", str(WORK / "build"),
                    "--specpath", str(WORK), "--paths", str(ROOT), "--collect-data", "pptx",
                    "--hidden-import", "win32timezone", "--exclude-module", "pytest",
                    "--exclude-module", "tkinter", str(ROOT / "main.py")], cwd=ROOT, check=True)
    # 构建目录与用户可能正在运行的便携版分离，不覆盖正在使用的 EXE 或 data。
    bundle = WORK / "portable" / "PPTSearch"
    copy_licenses(bundle)
    for name in ("README.md", "LICENSE", "THIRD_PARTY_NOTICES.md", "requirements-lock.txt"):
        shutil.copy2(ROOT / name, bundle / name)
    shutil.copytree(ROOT / "docs", bundle / "docs", dirs_exist_ok=True)
    for name in ("acceptance.json", "gui-preview.png"):
        source = WORK / name
        if source.exists():
            shutil.copy2(source, OUTPUTS / name)
    shutil.copy2(ROOT / "README.md", OUTPUTS / "README.md")
    release = OUTPUTS / "PPTSearch-1.0.0"
    shutil.copytree(bundle, release, dirs_exist_ok=True)
    portable = shutil.make_archive(str(OUTPUTS / "PPTSearch-1.0.0-win64"), "zip", bundle.parent, bundle.name)
    source_archive(OUTPUTS / "PPTSearch-1.0.0-source.zip")
    print(json.dumps({"executable": str(release / "PPTSearch.exe"), "portable_zip": portable}, ensure_ascii=False))


if __name__ == "__main__":
    main()
