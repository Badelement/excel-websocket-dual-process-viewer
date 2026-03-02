"""
read_data.py（读取数据进程）
===========================

职责非常单一：
1. 读取 data.xlsx
2. 转成 Python 二维数组
3. 通过 WebSocket 发给连接过来的客户端（show_data 进程）
4. 限制自己只能启动一个实例
"""

import asyncio
import json
import socket
import sys
from pathlib import Path

import pandas as pd
import websockets

# 仅本机通信
HOST = "127.0.0.1"
PORT = 8765
# read_data 的单实例锁端口
READ_LOCK_PORT = 54320


def ensure_single_instance(lock_port: int, process_name: str):
    """
    单实例控制：
    通过绑定本地端口来判断是否已有同类进程在运行。
    """
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.bind(("127.0.0.1", lock_port))
    except OSError:
        print(f"{process_name}程序已启动！")
        sys.exit(1)
    return sock


def _resource_base_dir() -> Path:
    """
    资源目录定位，兼容 PyInstaller onefile。
    """
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def load_excel_2d_array() -> list:
    """
    读取 Excel 并返回二维数组：
    - 第 1 行（表头）=> list 第 0 个元素
    - 后续数据行 => list 后续元素
    """
    excel_path = _resource_base_dir() / "data.xlsx"
    if not excel_path.exists():
        raise FileNotFoundError(f"未找到数据文件: {excel_path}")

    df = pd.read_excel(excel_path)
    return [df.columns.tolist()] + df.fillna("").values.tolist()


async def read_handler(websocket):
    """
    WebSocket 请求处理：
    每个客户端连接进来后，发送一次完整二维数组。
    """
    data = load_excel_2d_array()
    message = json.dumps({"type": "table", "data": data}, ensure_ascii=False)
    await websocket.send(message)
    await websocket.wait_closed()


async def run_server():
    """
    启动 WebSocket 服务端并常驻运行。
    """
    async with websockets.serve(read_handler, HOST, PORT):
        print(f"read_data WebSocket已启动: ws://{HOST}:{PORT}")
        await asyncio.Future()


def main():
    """
    进程入口：
    先抢占单实例锁，再运行服务。
    """
    lock = ensure_single_instance(READ_LOCK_PORT, "read_data")
    try:
        asyncio.run(run_server())
    finally:
        lock.close()


if __name__ == "__main__":
    main()

