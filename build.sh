#!/bin/bash

rm -rf build/ dist/
docker run --rm -v "$(pwd):/src/" -e PYPI_INDEX_URL=https://pypi.tuna.tsinghua.edu.cn/simple cdrx/pyinstaller-windows:python3-32bit "pyinstaller -i ok.ico -F main.py"