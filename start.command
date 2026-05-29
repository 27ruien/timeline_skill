#!/bin/zsh
cd "$(dirname "$0")"

echo "Starting Timeline Maker..."

if ! command -v python3 >/dev/null 2>&1; then
  echo "未找到 python3。请先安装 Python 3。"
  echo "可以从 https://www.python.org/downloads/ 下载。"
  read "?按回车退出"
  exit 1
fi

if [ ! -x ".venv/bin/python" ]; then
  echo "首次启动：正在创建本地运行环境..."
  python3 -m venv .venv
  if [ $? -ne 0 ]; then
    echo "创建本地运行环境失败。"
    read "?按回车退出"
    exit 1
  fi
fi

echo "正在检查依赖..."
.venv/bin/python - <<'PY'
import importlib.util
missing = [
    package
    for package in ("openpyxl", "PIL")
    if not importlib.util.find_spec(package)
]
raise SystemExit(0 if not missing else 1)
PY

if [ $? -ne 0 ]; then
  echo "首次启动：正在安装 Excel 生成依赖..."
  .venv/bin/python -m pip install -r requirements.txt
  if [ $? -ne 0 ]; then
    echo "依赖安装失败。请检查网络后重新双击 start.command。"
    read "?按回车退出"
    exit 1
  fi
fi

echo "启动本地网页..."
.venv/bin/python local_app.py
