"""
show_data.py（展示进程）
=======================

职责：
1. 作为 WebSocket 客户端连接 read_data
2. 接收二维数组
3. 使用 PyQt 表格控件展示数据
4. 限制自身单实例
"""

import asyncio
import json
import socket
import sys

import websockets
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QTableWidget, QTableWidgetItem

HOST = "127.0.0.1"
PORT = 8765
SHOW_LOCK_PORT = 54321


def ensure_single_instance(lock_port: int, process_name: str):
    """
    单实例控制：
    绑定固定端口失败则说明已有 show_data 在运行。
    """
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.bind(("127.0.0.1", lock_port))
    except OSError:
        print(f"{process_name}程序已启动！")
        sys.exit(1)
    return sock


class WSClientThread(QThread):
    """
    WebSocket 客户端工作线程

    说明：
    - UI 主线程只负责界面绘制
    - 网络连接放到子线程，避免窗口卡顿
    """

    # 收到二维数组后发给主线程
    received = pyqtSignal(list)
    # 失败时发错误信息给主线程
    failed = pyqtSignal(str)

    def run(self):
        asyncio.run(self._consume())

    async def _consume(self):
        """
        连接 read_data 并接收数据。
        若 read_data 还没启动，短暂重试后再报错。
        """
        ws_url = f"ws://{HOST}:{PORT}"
        for _ in range(20):
            try:
                async with websockets.connect(ws_url) as ws:
                    message = await ws.recv()
                    payload = json.loads(message)
                    self.received.emit(payload["data"])
                    return
            except Exception:
                await asyncio.sleep(0.3)
        self.failed.emit("无法连接 read_data 进程，请先启动 read_data.py。")


class MainWindow(QMainWindow):
    """
    主窗口：把二维数组显示到 QTableWidget。
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
        渲染规则：
        - table_data[0] 是表头
        - table_data[1:] 是数据行
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


def main():
    """
    进程入口：
    抢占单实例锁 -> 启动 UI -> 退出时释放锁。
    """
    lock = ensure_single_instance(SHOW_LOCK_PORT, "show_data")
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    code = app.exec_()
    lock.close()
    sys.exit(code)


if __name__ == "__main__":
    main()

