from distutils.core import setup
import py2exe

setup(windows=["Text2PDF"])

setup(
    windows=[{"script": "prcheck.py", "icon_resources": [(1, "prcheck.ico")]}],
    options={"py2exe": {"packages": ["xml", "gzip"]}},
    data_files=[("", ["PR Structure.ods", "prchecker_splash.gif", "uudeview.exe"])],
)


# Text2PDF must come first or prcheck.exe will fail due to module import errors
