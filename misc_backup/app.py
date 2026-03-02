"""
项目主入口（总控程序）
====================

这个文件是整个程序的“调度中心”，核心职责是：
1. 以 --mode 参数切换三种运行角色：
   - launch：主启动模式（默认），负责拉起 read + show 两个子进程
   - read  ：读数据进程，读取 Excel 并通过 WebSocket 对外提供数据
   - show  ：展示进程，从 WebSocket 拿数据并用 PyQt 表格展示
2. 保证单实例运行：
   - 主程序只能开一个
   - read 子进程只能开一个
   - show 子进程只能开一个
3. 处理生命周期：
   - 主程序先启动 read，再启动 show
   - 当 show 退出时，主程序会主动结束 read，避免后台残留进程
"""

import argparse
import asyncio
import ctypes
import json
import socket
import subprocess
import sys
import time
from pathlib import Path

import pandas as pd
import websockets
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QTableWidget, QTableWidgetItem

# 本机回环地址，只允许本机进程通信（不对外网开放）
HOST = "127.0.0.1"
# WebSocket 服务端端口（read 监听、show 连接）
PORT = 8765
# 三种角色对应的“进程锁端口”。
# 原理：绑定本地端口成功 = 当前进程是首个实例；失败 = 已有实例在运行。
READ_LOCK_PORT = 54320
SHOW_LOCK_PORT = 54321
LAUNCH_LOCK_PORT = 54322


def ensure_single_instance(lock_port: int, process_name: str):
    """
    单实例控制（通用函数）

    参数:
    - lock_port: 用于“抢占锁”的本地端口
    - process_name: 进程名称（用于提示文案）

    返回:
    - sock: 成功抢占后返回 socket 对象。调用方必须在退出时 close() 释放锁。
    """
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        # 绑定成功 => 没有其他同类进程在运行
        sock.bind(("127.0.0.1", lock_port))
    except OSError:
        # 绑定失败 => 端口已被占用，说明已有实例
        message = "主程序已启动！" if process_name == "main_app" else f"{process_name}程序已启动！"

        # 打包后是无控制台窗口（--noconsole），因此这里用系统弹窗提示用户
        if getattr(sys, "frozen", False):
            try:
                ctypes.windll.user32.MessageBoxW(0, message, "提示", 0x40)
            except Exception:
                pass

        # 开发模式下仍打印到控制台
        print(message)
        sys.exit(1)
    return sock


def _resource_base_dir() -> Path:
    """
    资源目录定位（关键：兼容 PyInstaller onefile）

    - 源码运行时：data.xlsx 在脚本同目录
    - 打包 onefile 运行时：资源会先解压到 sys._MEIPASS 临时目录
    """
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def load_excel_2d_array() -> list:
    """
    读取 Excel -> 转成二维数组（Python list[list]）

    输出格式示例:
    [
      ["列1", "列2", "列3"],   # 表头行
      [值11, 值12, 值13],      # 数据行1
      [值21, 值22, 值23],      # 数据行2
      ...
    ]
    """
    excel_path = _resource_base_dir() / "data.xlsx"
    if not excel_path.exists():
        raise FileNotFoundError(f"未找到数据文件: {excel_path}")

    df = pd.read_excel(excel_path)
    # fillna("") 把空值转为空字符串，避免 UI 里显示 NaN
    table = [df.columns.tolist()] + df.fillna("").values.tolist()
    return table


async def read_handler(websocket):
    """
    WebSocket 请求处理函数（read 进程使用）
    每有一个客户端连进来，就把当前表格快照发送一次。
    """
    data = load_excel_2d_array()
    message = json.dumps({"type": "table", "data": data}, ensure_ascii=False)
    await websocket.send(message)
    await websocket.wait_closed()


async def run_read_server():
    """
    启动 WebSocket 服务端并常驻。
    """
    async with websockets.serve(read_handler, HOST, PORT):
        print(f"read_data WebSocket已启动: ws://{HOST}:{PORT}")
        # 用一个永不完成的 Future 让协程常驻
        await asyncio.Future()


def run_read_process():
    """
    read 角色入口：
    1) 抢占 read 进程锁
    2) 启动 WebSocket 服务
    3) 退出时释放锁
    """
    lock = ensure_single_instance(READ_LOCK_PORT, "read_data")
    try:
        asyncio.run(run_read_server())
    finally:
        lock.close()


class WSClientThread(QThread):
    """
    WebSocket 客户端线程（show 角色使用）

    为什么要放在线程里？
    - 网络操作可能阻塞
    - 若直接在 UI 主线程做网络 I/O，会导致窗口卡死、无响应
    """

    received = pyqtSignal(list)  # 成功拿到二维数组后发出
    failed = pyqtSignal(str)     # 失败后发出错误文案

    def run(self):
        asyncio.run(self._consume())

    async def _consume(self):
        ws_url = f"ws://{HOST}:{PORT}"
        # 给 read 进程预留启动时间：最多重试约 6 秒（20 * 0.3）
        for _ in range(20):
            try:
                async with websockets.connect(ws_url) as ws:
                    message = await ws.recv()
                    payload = json.loads(message)
                    self.received.emit(payload["data"])
                    return
            except Exception:
                await asyncio.sleep(0.3)

        self.failed.emit("无法连接 read_data 进程，请先启动 read_data.py 或总程序。")


class MainWindow(QMainWindow):
    """
    PyQt 主窗口：把二维数组渲染到 QTableWidget。
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel数据展示")
        self.resize(900, 600)

        self.table = QTableWidget()
        self.setCentralWidget(self.table)

        self.worker = WSClientThread()
        self.worker.received.connect(self.show_table)
        self.worker.failed.connect(self.show_error)
        self.worker.start()

    def show_table(self, table_data: list):
        """
        把二维数组映射到表格控件：
        - 第 0 行作为表头
        - 其余行作为数据内容
        """
        if not table_data:
            self.show_error("收到空数据。")
            return

        headers = [str(x) for x in table_data[0]]
        rows = table_data[1:]

        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(len(rows))

        for i, row in enumerate(rows):
            for j, value in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(str(value)))

    def show_error(self, text: str):
        QMessageBox.critical(self, "错误", text)


def run_show_process():
    """
    show 角色入口：
    1) 抢占 show 进程锁
    2) 启动 PyQt 窗口
    3) 窗口退出时释放锁
    """
    lock = ensure_single_instance(SHOW_LOCK_PORT, "show_data")
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    code = app.exec_()
    lock.close()
    sys.exit(code)


def spawn_launcher_pair():
    """
    launch 模式核心逻辑（主程序入口）：
    - read 先启动（先把数据服务端准备好）
    - show 后启动（再去连接 read）
    - show 关闭后，主程序负责结束 read，防止 read 成为孤儿进程
    """
    script_path = Path(__file__).resolve()

    # 打包后（frozen）启动自身时不用再带脚本路径
    if getattr(sys, "frozen", False):
        cmd = [sys.executable]
    else:
        cmd = [sys.executable, str(script_path)]

    read_proc = subprocess.Popen([*cmd, "--mode", "read"])
    time.sleep(0.8)
    show_proc = subprocess.Popen([*cmd, "--mode", "show"])

    show_proc.wait()

    read_proc.terminate()
    try:
        read_proc.wait(timeout=3)
    except subprocess.TimeoutExpired:
        read_proc.kill()


def main():
    """
    程序总入口：
    - 解析 --mode
    - 分发到 read/show/launch 对应逻辑
    """
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=["read", "show", "launch"], default="launch")
    args = parser.parse_args()

    if args.mode == "read":
        run_read_process()
    elif args.mode == "show":
        run_show_process()
    else:
        # 主启动器也限制单实例，防止重复点击后拉起多套子进程
        lock = ensure_single_instance(LAUNCH_LOCK_PORT, "main_app")
        try:
            spawn_launcher_pair()
        finally:
            lock.close()


if __name__ == "__main__":
    main()

