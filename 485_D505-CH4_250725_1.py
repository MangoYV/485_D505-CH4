import sys
import struct
import serial
import serial.tools.list_ports
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGroupBox,
                             QLabel, QComboBox, QLineEdit, QPushButton, QTextEdit, QTabWidget,
                             QGridLayout, QMessageBox, QCheckBox, QScrollBar)
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QFont


class ModbusRTUTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("D505-CH4四通道测力显示仪通信工具 - Modbus-RTU")
        self.setGeometry(750, 280, 850, 695)
        self.serial_port = None

        font = QFont("微软雅黑", 12)
        self.setFont(font)

        self.init_ui()
        self.serial_connected = False
        self.read_count = 0
        self.write_count = 0

        self.auto_connect_timer = QTimer()
        self.auto_connect_timer.timeout.connect(self.auto_open_serial)
        self.auto_connect_timer.setSingleShot(True)
        self.auto_connect_timer.start(1000)

    def auto_open_serial(self):
        self.open_serial()

    def init_ui(self):
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

        tabs = QTabWidget()
        main_layout.addWidget(tabs)

        settings_tab = QWidget()
        settings_layout = QVBoxLayout(settings_tab)

        port_group = QGroupBox("串口设置")
        port_layout = QGridLayout()

        port_layout.addWidget(QLabel("串口号:"), 0, 0)
        self.port_combo = QComboBox()
        self.port_combo.setMinimumWidth(100)
        self.refresh_ports()
        com3_index = self.port_combo.findText("COM3")
        if com3_index >= 0:
            self.port_combo.setCurrentIndex(com3_index)
        port_layout.addWidget(self.port_combo, 0, 1)

        self.refresh_button = QPushButton("刷新端口")
        self.refresh_button.clicked.connect(self.refresh_ports)
        port_layout.addWidget(self.refresh_button, 0, 2)

        port_layout.setColumnStretch(1, 1)

        port_layout.addWidget(QLabel("波特率:"), 1, 0)
        self.baud_combo = QComboBox()
        self.baud_combo.setMinimumWidth(100)
        self.baud_combo.addItems(["9600", "19200", "38400", "57600", "115200"])
        self.baud_combo.setCurrentText("115200")
        port_layout.addWidget(self.baud_combo, 1, 1)

        port_layout.addWidget(QLabel("数据位:"), 2, 0)
        self.data_bits_combo = QComboBox()
        self.data_bits_combo.setMinimumWidth(100)
        self.data_bits_combo.addItems(["5", "6", "7", "8"])
        self.data_bits_combo.setCurrentText("8")
        port_layout.addWidget(self.data_bits_combo, 2, 1)

        port_layout.addWidget(QLabel("停止位:"), 3, 0)
        self.stop_bits_combo = QComboBox()
        self.stop_bits_combo.setMinimumWidth(100)
        self.stop_bits_combo.addItems(["1", "1.5", "2"])
        self.stop_bits_combo.setCurrentText("1")
        port_layout.addWidget(self.stop_bits_combo, 3, 1)

        port_layout.addWidget(QLabel("校验位:"), 4, 0)
        self.parity_combo = QComboBox()
        self.parity_combo.setMinimumWidth(100)
        self.parity_combo.addItems(["无", "奇校验", "偶校验"])
        self.parity_combo.setCurrentText("无")
        port_layout.addWidget(self.parity_combo, 4, 1)

        port_layout.addWidget(QLabel("从站地址:"), 5, 0)
        self.slave_address_edit = QLineEdit("1")
        self.slave_address_edit.setValidator(self.create_int_validator(1, 247))
        port_layout.addWidget(self.slave_address_edit, 5, 1)

        self.connect_button = QPushButton("打开串口")
        self.connect_button.clicked.connect(self.toggle_connection)
        port_layout.addWidget(self.connect_button, 6, 0, 1, 3)

        port_group.setLayout(port_layout)
        settings_layout.addWidget(port_group)

        info_group = QGroupBox("操作说明")
        info_layout = QVBoxLayout()
        info_text = QTextEdit()
        info_text.setReadOnly(True)
        info_text.setPlainText(
            "1. 默认串口设置：115200波特率, 8数据位, 1停止位, 无校验\n"
            "2. 从站地址范围：1-247\n"
            "3. 寄存器地址：0-65535 （32位数据占用2个连续寄存器）\n"
            "4. 读取数据：选择起始地址和读取寄存器数量（必须为2的倍数）\n"
            "5. 写入数据：输入32位数据（长整型或浮点型）\n"
            "6. 修改通讯参数后需重新上电生效"
        )
        info_layout.addWidget(info_text)
        info_group.setLayout(info_layout)
        settings_layout.addWidget(info_group)

        data_tab = QWidget()
        data_layout = QVBoxLayout(data_tab)

        read_write_layout = QHBoxLayout()

        read_group = QGroupBox("读取操作")
        read_layout = QGridLayout()

        self.read_address_edits = []
        for i in range(8):
            row = i // 2
            col = (i % 2) * 2

            label = QLabel(f"地址{i + 1}:")
            read_layout.addWidget(label, row, col)

            if i < 4:
                default_address = str(2000 + i * 2)
            else:
                default_address = "0"

            address_edit = QLineEdit(default_address)
            address_edit.setValidator(self.create_int_validator(0, 65535))
            read_layout.addWidget(address_edit, row, col + 1)
            self.read_address_edits.append(address_edit)

        read_layout.addWidget(QLabel("数据缩放:"), 4, 0)
        self.scale_factor_edit = QLineEdit("0.1")
        self.scale_factor_edit.setValidator(self.create_float_validator())
        read_layout.addWidget(self.scale_factor_edit, 4, 1)

        read_layout.addWidget(QLabel("数据类型:"), 4, 2)
        self.data_type_combo = QComboBox()
        self.data_type_combo.addItems(["浮点型 （Float）", "长整型 （Long）"])
        self.data_type_combo.setCurrentIndex(1)
        read_layout.addWidget(self.data_type_combo, 4, 3)

        self.read_button = QPushButton("读取数据")
        self.read_button.clicked.connect(self.read_data)
        read_layout.addWidget(self.read_button, 5, 0, 1, 4)

        read_group.setLayout(read_layout)
        read_write_layout.addWidget(read_group, 3)

        write_group = QGroupBox("写入操作")
        write_layout = QGridLayout()

        write_layout.addWidget(QLabel("起始地址:"), 0, 0)
        self.write_address_edit = QLineEdit("2000")
        self.write_address_edit.setValidator(self.create_int_validator(0, 65535))
        write_layout.addWidget(self.write_address_edit, 0, 1)

        write_layout.addWidget(QLabel("写入值:"), 1, 0)
        self.write_value_edit = QLineEdit("0.0")
        write_layout.addWidget(self.write_value_edit, 1, 1)

        write_layout.addWidget(QLabel("数据类型:"), 2, 0)
        self.write_type_combo = QComboBox()
        self.write_type_combo.addItems(["浮点型 （Float）", "长整型 （Long）"])
        self.write_type_combo.setCurrentIndex(0)
        write_layout.addWidget(self.write_type_combo, 2, 1)

        write_layout.addWidget(QLabel(""), 3, 0)
        self.write_button = QPushButton("写入数据")
        self.write_button.clicked.connect(self.write_data)
        write_layout.addWidget(self.write_button, 4, 0, 1, 2)

        write_group.setLayout(write_layout)
        read_write_layout.addWidget(write_group, 2)

        data_layout.addLayout(read_write_layout)

        result_group = QGroupBox("通信结果")
        result_layout = QHBoxLayout()
        result_layout.setContentsMargins(5, 15, 5, 5)

        comm_group = QGroupBox("通信命令")
        comm_layout = QVBoxLayout()
        self.comm_text = QTextEdit()
        self.comm_text.setReadOnly(True)
        comm_layout.addWidget(self.comm_text)
        comm_group.setLayout(comm_layout)
        result_layout.addWidget(comm_group, 3)

        result_display_group = QGroupBox("读取结果")
        result_display_layout = QVBoxLayout()
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        result_display_layout.addWidget(self.result_text)
        result_display_group.setLayout(result_display_layout)
        result_layout.addWidget(result_display_group, 2)

        result_group.setLayout(result_layout)
        data_layout.addWidget(result_group, 1)

        clear_layout = QHBoxLayout()
        self.clear_button = QPushButton("清空结果")
        self.clear_button.clicked.connect(self.clear_results)
        clear_layout.addStretch(1)
        clear_layout.addWidget(self.clear_button)
        clear_layout.addStretch(1)
        data_layout.addLayout(clear_layout)

        tabs.addTab(data_tab, "数据操作")
        tabs.addTab(settings_tab, "串口设置")
        
        tabs.setCurrentIndex(0)

        self.status_bar = self.statusBar()
        self.status_bar.showMessage("准备就绪")

        self.serial_connected = False

        self.comm_text.verticalScrollBar().valueChanged.connect(self.on_comm_scroll)
        self.result_text.verticalScrollBar().valueChanged.connect(self.on_result_scroll)

        self.scrolling = False

    def on_comm_scroll(self, value):
        if not self.scrolling:
            self.scrolling = True
            self.result_text.verticalScrollBar().setValue(value)
            self.scrolling = False

    def on_result_scroll(self, value):
        if not self.scrolling:
            self.scrolling = True
            self.comm_text.verticalScrollBar().setValue(value)
            self.scrolling = False

    def create_int_validator(self, min_val, max_val):
        from PyQt5.QtGui import QIntValidator
        validator = QIntValidator()
        validator.setRange(min_val, max_val)
        return validator

    def create_float_validator(self):
        from PyQt5.QtGui import QDoubleValidator
        validator = QDoubleValidator()
        validator.setBottom(0.0001)
        return validator

    def refresh_ports(self):
        self.port_combo.clear()
        ports = serial.tools.list_ports.comports()
        for port in ports:
            self.port_combo.addItem(port.device)
        if not ports:
            self.port_combo.addItem("无可用串口")

    def toggle_connection(self):
        if self.serial_connected:
            self.close_serial()
        else:
            self.open_serial()

    def open_serial(self):
        port = self.port_combo.currentText()
        if port == "无可用串口" or not port:
            QMessageBox.warning(self, "错误", "没有可用的串口")
            return

        try:
            baudrate = int(self.baud_combo.currentText())
            bytesize = int(self.data_bits_combo.currentText())
            stopbits = float(self.stop_bits_combo.currentText())
            if stopbits == 1.0:
                stopbits = serial.STOPBITS_ONE
            elif stopbits == 1.5:
                stopbits = serial.STOPBITS_ONE_POINT_FIVE
            elif stopbits == 2.0:
                stopbits = serial.STOPBITS_TWO

            parity_text = self.parity_combo.currentText()
            if parity_text == "无":
                parity = serial.PARITY_NONE
            elif parity_text == "奇校验":
                parity = serial.PARITY_ODD
            elif parity_text == "偶校验":
                parity = serial.PARITY_EVEN

            self.serial_port = serial.Serial(
                port=port,
                baudrate=baudrate,
                bytesize=bytesize,
                parity=parity,
                stopbits=stopbits,
                timeout=1.0
            )

            self.serial_connected = True
            self.connect_button.setText("关闭串口")
            self.status_bar.showMessage(f"已连接到 {port}, {baudrate}波特率")
            connection_info = f"串口已打开: {port}, {baudrate}波特率"
            self.comm_text.append(f'<span style="color:black">{connection_info}</span>')
            self.result_text.append(connection_info)
        except Exception as e:
            QMessageBox.critical(self, "连接错误", f"无法打开串口: {str(e)}")
            self.status_bar.showMessage(f"连接失败: {str(e)}")

    def close_serial(self):
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
        self.serial_connected = False
        self.connect_button.setText("打开串口")
        self.status_bar.showMessage("串口已关闭")
        self.comm_text.append("--------------------------------")
        self.comm_text.append('<span style="color:blue">串口已关闭</span>')
        self.result_text.append("--------------------------------")
        self.result_text.append("串口已关闭")

    def calculate_crc(self, data):
        crc = 0xFFFF
        for byte in data:
            crc ^= byte
            for _ in range(8):
                if crc & 0x0001:
                    crc >>= 1
                    crc ^= 0xA001
                else:
                    crc >>= 1
        return crc.to_bytes(2, 'little')

    def read_data(self):
        if not self.serial_connected:
            QMessageBox.warning(self, "错误", "请先打开串口")
            return

        try:
            slave_address = int(self.slave_address_edit.text())
            scale_factor = float(self.scale_factor_edit.text())
            data_type = self.data_type_combo.currentText()

            self.read_count += 1

            self.comm_text.append("--------------------------------")
            self.result_text.append("--------------------------------")

            valid_addresses = [int(addr.text()) for addr in self.read_address_edits if int(addr.text()) != 0]

            if not valid_addresses:
                self.comm_text.append(f"读取结果：第{self.read_count}次")
                self.comm_text.append("")
                self.result_text.append(f"读取结果：第{self.read_count}次")
                self.result_text.append("")
                self.scroll_to_bottom()
                return

            self.comm_text.append(f"读取结果：第{self.read_count}次")
            self.comm_text.append("")
            self.result_text.append(f"读取结果：第{self.read_count}次")
            self.result_text.append("")

            for i, address_edit in enumerate(self.read_address_edits):
                start_address = int(address_edit.text())

                if start_address == 0:
                    continue

                command = bytearray()
                command.append(slave_address)
                command.append(0x03)
                command.append((start_address >> 8) & 0xFF)
                command.append(start_address & 0xFF)
                command.append(0x00)
                command.append(0x02)

                crc = self.calculate_crc(command)
                command.extend(crc)

                self.serial_port.write(command)
                self.comm_text.append(f'<span style="color:red">发送读取命令（地址{start_address}）：{command.hex(" ").upper()}</span>')

                expected_length = 9
                response = self.serial_port.read(expected_length)

                if not response:
                    self.comm_text.append(f'<span style="color:blue">收到响应数据（地址{start_address}）：读取超时，未收到响应</span>')
                    self.result_text.append(f'<span style="color:red">地址{start_address}：</span>')
                    self.result_text.append(f'<span style="color:blue">\t读取超时，未收到响应</span>')
                    continue

                self.comm_text.append(f'<span style="color:blue">收到响应数据（地址{start_address}）：{response.hex(" ").upper()}</span>')

                if len(response) < 5:
                    self.result_text.append(f'<span style="color:red">地址{start_address}：</span>')
                    self.result_text.append(f'<span style="color:blue">\t响应长度不足</span>')
                    continue

                received_crc = response[-2:]
                calculated_crc = self.calculate_crc(response[:-2])
                if received_crc != calculated_crc:
                    self.result_text.append(f'<span style="color:red">地址{start_address}：</span>')
                    self.result_text.append(f'<span style="color:blue">\tCRC校验失败</span>')
                    continue

                if response[0] != slave_address:
                    self.result_text.append(f'<span style="color:red">地址{start_address}：</span>')
                    self.result_text.append(f'<span style="color:blue">\t从站地址不匹配</span>')
                    continue

                if response[1] != 0x03:
                    self.result_text.append(f'<span style="color:red">地址{start_address}：</span>')
                    self.result_text.append(f'<span style="color:blue">\t功能码错误</span>')
                    continue

                byte_count = response[2]
                data_bytes = response[3:3 + byte_count]

                if "浮点型" in data_type:
                    try:
                        value = struct.unpack('>f', data_bytes)[0]
                        scaled_value = value * scale_factor
                        formatted_value = f"{scaled_value:.5f}".rstrip('0').rstrip('.')
                        self.result_text.append(f'<span style="color:red">地址{start_address}：</span>')
                        self.result_text.append(f'<span style="color:blue">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;数值：{formatted_value}</span>')
                    except:
                        self.result_text.append(f'<span style="color:red">地址{start_address}：</span>')
                        self.result_text.append(f'<span style="color:blue">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;数值：解析浮点数错误</span>')
                else:
                    try:
                        value = (data_bytes[0] << 24) | (data_bytes[1] << 16) | (data_bytes[2] << 8) | data_bytes[3]
                        scaled_value = value * scale_factor
                        if scale_factor == 1:
                            self.result_text.append(f'<span style="color:red">地址{start_address}：</span>')
                            self.result_text.append(f'<span style="color:blue">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;数值：{int(scaled_value)}</span>')
                        else:
                            formatted_value = f"{scaled_value:.5f}".rstrip('0').rstrip('.')
                            self.result_text.append(f'<span style="color:red">地址{start_address}：</span>')
                            self.result_text.append(f'<span style="color:blue">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;数值：{formatted_value}</span>')
                    except:
                        self.result_text.append(f'<span style="color:red">地址{start_address}：</span>')
                        self.result_text.append(f'<span style="color:blue">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;数值：解析长整型错误</span>')

            self.scroll_to_bottom()

        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取数据时发生错误: {str(e)}")
            self.result_text.append(f"错误: {str(e)}")
            self.scroll_to_bottom()

    def write_data(self):
        if not self.serial_connected:
            QMessageBox.warning(self, "错误", "请先打开串口")
            return

        try:
            slave_address = int(self.slave_address_edit.text())
            start_address = int(self.write_address_edit.text())
            data_type = self.write_type_combo.currentText()
            value_str = self.write_value_edit.text()

            self.write_count += 1

            self.comm_text.append("--------------------------------")
            self.result_text.append("--------------------------------")

            if "浮点型" in data_type:
                try:
                    value = float(value_str)
                    value_bytes = struct.pack('>f', value)
                except ValueError:
                    QMessageBox.warning(self, "错误", "无效的浮点数值")
                    self.scroll_to_bottom()
                    return
            else:
                try:
                    value = int(value_str)
                    value_bytes = bytes([
                        (value >> 24) & 0xFF,
                        (value >> 16) & 0xFF,
                        (value >> 8) & 0xFF,
                        value & 0xFF
                    ])
                except ValueError:
                    QMessageBox.warning(self, "错误", "无效的长整型值")
                    self.scroll_to_bottom()
                    return

            register_count = 2
            byte_count = 4

            command = bytearray()
            command.append(slave_address)
            command.append(0x10)
            command.append((start_address >> 8) & 0xFF)
            command.append(start_address & 0xFF)
            command.append((register_count >> 8) & 0xFF)
            command.append(register_count & 0xFF)
            command.append(byte_count)
            command.extend(value_bytes)

            crc = self.calculate_crc(command)
            command.extend(crc)

            self.serial_port.write(command)
            self.comm_text.append(f'发送结果：第{self.write_count}次')
            self.comm_text.append(f'<span style="color:red">发送写入命令：{command.hex(" ").upper()}</span>')

            expected_length = 8
            response = self.serial_port.read(expected_length)

            if not response:
                self.comm_text.append('<span style="color:blue">收到响应数据：读取超时，未收到响应</span>')
                self.result_text.append("写入超时，未收到响应")
                self.scroll_to_bottom()
                return

            self.comm_text.append(f'<span style="color:blue">收到响应数据：{response.hex(" ").upper()}</span>')

            if len(response) < 6:
                self.comm_text.append('<span style="color:red">响应长度不足</span>')
                self.result_text.append("响应长度不足")
                self.scroll_to_bottom()
                return

            received_crc = response[-2:]
            calculated_crc = self.calculate_crc(response[:-2])
            if received_crc != calculated_crc:
                self.comm_text.append(f'<span style="color:red">CRC校验失败： 收到 {received_crc.hex()} 计算 {calculated_crc.hex()}</span>')
                self.result_text.append(
                    f"CRC校验失败： 收到 {received_crc.hex()} 计算 {calculated_crc.hex()}")
                self.scroll_to_bottom()
                return

            if response[0] != slave_address:
                self.comm_text.append(f'<span style="color:red">从站地址不匹配： 收到 {response[0]}, 期望 {slave_address}</span>')
                self.result_text.append(f"从站地址不匹配： 收到 {response[0]}, 期望 {slave_address}")
                self.scroll_to_bottom()
                return

            if response[1] != 0x10:
                self.comm_text.append(f'<span style="color:red">功能码错误： 收到 {hex(response[1])}, 期望 0x10</span>')
                self.result_text.append(f"功能码错误： 收到 {hex(response[1])}, 期望 0x10")
                self.scroll_to_bottom()
                return

            resp_start_address = (response[2] << 8) | response[3]
            resp_register_count = (response[4] << 8) | response[5]

            if resp_start_address != start_address:
                self.comm_text.append(f'<span style="color:red">写入地址不匹配： 收到 {resp_start_address}, 期望 {start_address}</span>')
                self.result_text.append(f"写入地址不匹配： 收到 {resp_start_address}, 期望 {start_address}")
                self.scroll_to_bottom()
                return

            if resp_register_count != register_count:
                self.comm_text.append(f'<span style="color:red">寄存器数量不匹配： 收到 {resp_register_count}, 期望 {register_count}</span>')
                self.result_text.append(
                    f"寄存器数量不匹配： 收到 {resp_register_count}, 期望 {register_count}")
                self.scroll_to_bottom()
                return

            self.result_text.append(f"写入成功：第{self.write_count}次")
            self.result_text.append(f'<span style="color:blue">写入地址：{start_address}</span>')
            self.result_text.append(f'<span style="color:blue">返回读值：{value_str}</span>')
            self.scroll_to_bottom()

        except Exception as e:
            QMessageBox.critical(self, "错误", f"写入数据时发生错误: {str(e)}")
            error_msg = f"错误: {str(e)}"
            self.comm_text.append(f'<span style="color:red">{error_msg}</span>')
            self.result_text.append(error_msg)
            self.scroll_to_bottom()

    def scroll_to_bottom(self):
        self.comm_text.verticalScrollBar().setValue(self.comm_text.verticalScrollBar().maximum())
        self.result_text.verticalScrollBar().setValue(self.result_text.verticalScrollBar().maximum())

    def clear_results(self):
        self.comm_text.clear()
        self.result_text.clear()

    def closeEvent(self, event):
        self.close_serial()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ModbusRTUTool()
    window.show()
    sys.exit(app.exec_())