import pathlib
from setuptools import Extension, setup
from Cython.Build import cythonize


MODULE_NAME = "clipd_core"
SOURCE_FILE = "clipboard_guardian.py"

extensions = [
    Extension(
        MODULE_NAME,
        [SOURCE_FILE],
    )
]

setup(
    name=MODULE_NAME,
    ext_modules=cythonize(
        extensions,
        compiler_directives={"language_level": "3"},
    ),
)
