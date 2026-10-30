# Third-party notices

PPT Search's own Python source is MIT licensed. Third-party software keeps its
respective license; this document does not relicense bundled dependencies.
The portable distribution includes original license files in THIRD_PARTY_LICENSES.
Exact versions of Python packages are recorded in requirements-lock.txt.

## Qt / PySide6 / Shiboken 6.8.3

This distribution uses the open-source LGPLv3 option of Qt for Python and the
applicable open-source license options of dynamically linked Qt modules.
Qt and Shiboken are unmodified. Compatible replacement shared libraries can be
placed in the portable distribution's _internal directory without rebuilding
PPT Search. No restriction on reverse engineering for debugging modifications
to LGPL-covered libraries is imposed by this application.

License texts: THIRD_PARTY_LICENSES/Qt/LGPL-3.0-only.txt and GPL-3.0-only.txt.

Upstream sources, build instructions, notices and module-specific licenses:

- Qt base 6.8.3: https://code.qt.io/cgit/qt/qtbase.git/tag/?h=v6.8.3
- Qt base source archive: https://download.qt.io/archive/qt/6.8/6.8.3/submodules/qtbase-everywhere-src-6.8.3.tar.xz
- Qt SVG 6.8.3: https://code.qt.io/cgit/qt/qtsvg.git/tag/?h=v6.8.3
- Qt image formats 6.8.3: https://code.qt.io/cgit/qt/qtimageformats.git/tag/?h=v6.8.3
- PySide / Shiboken: https://code.qt.io/cgit/pyside/pyside-setup.git/tag/?h=v6.8.3
- Qt for Python license documentation: https://doc.qt.io/qtforpython-6/licenses.html

## Other dependencies

- python-pptx: MIT — https://github.com/scanny/python-pptx
- pywin32: Python Software Foundation and component-specific licenses — https://github.com/mhammond/pywin32
- lxml: BSD and component-specific libxml2 / libxslt licenses — https://lxml.de/
- Pillow: HPND and bundled codec-specific licenses — https://python-pillow.github.io/
- XlsxWriter: BSD-2-Clause — https://github.com/jmcnamara/XlsxWriter
- typing_extensions: PSF — https://github.com/python/typing_extensions
- Python: PSF and component-specific notices — https://www.python.org/
- PyInstaller build tool: GPL with bootloader exception permitting distribution of the bundled application — https://pyinstaller.org/

Build-time tests and tooling are not imported by the desktop application.
Their original license files are nevertheless included for provenance.
