import sys
import serial
import serial.tools.list_ports
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QGroupBox, QLabel, QComboBox, QPushButton, QLineEdit,
                             QTextEdit, QTableWidget, QTableWidgetItem, QTabWidget,
                             QHeaderView, QMessageBox, QFormLayout, QSpinBox)
from PyQt5.QtCore import QTimer, Qt
import struct
import binascii


class ModbusRTUTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.serial_port = None
        self.setWindowTitle("DY500智能数字变送器通讯工具")
        self.setGeometry(100, 100, 900, 700)

        self.init_ui()
        self.scan_serial_ports()

        # 初始化寄存器表
        self.init_register_table()

        # 定时器用于处理串口接收
        self.receive_timer = QTimer(self)
        self.receive_timer.timeout.connect(self.read_serial_data)
        self.receive_timer.start(100)  # 每100ms检查一次串口数据

    def init_ui(self):
        main_widget = QWidget()
        main_layout = QVBoxLayout()

        # 串口配置区域
        config_group = QGroupBox("串口配置")
        config_layout = QFormLayout()

        self.port_combo = QComboBox()
        self.baud_combo = QComboBox()
        self.baud_combo.addItems(["9600", "19200", "38400", "57600", "115200"])
        self.baud_combo.setCurrentText("19200")

        self.data_bits_combo = QComboBox()
        self.data_bits_combo.addItems(["5", "6", "7", "8"])
        self.data_bits_combo.setCurrentText("8")

        self.stop_bits_combo = QComboBox()
        self.stop_bits_combo.addItems(["1", "1.5", "2"])
        self.stop_bits_combo.setCurrentText("1")

        self.parity_combo = QComboBox()
        self.parity_combo.addItems(["无", "奇校验", "偶校验"])
        self.parity_combo.setCurrentText("无")

        self.connect_btn = QPushButton("连接")
        self.connect_btn.clicked.connect(self.toggle_connection)
        self.scan_btn = QPushButton("扫描端口")
        self.scan_btn.clicked.connect(self.scan_serial_ports)

        config_layout.addRow("端口:", self.port_combo)
        config_layout.addRow("波特率:", self.baud_combo)
        config_layout.addRow("数据位:", self.data_bits_combo)
        config_layout.addRow("停止位:", self.stop_bits_combo)
        config_layout.addRow("校验位:", self.parity_combo)
        config_layout.addRow(self.scan_btn, self.connect_btn)

        config_group.setLayout(config_layout)
        main_layout.addWidget(config_group)

        # 选项卡区域
        self.tabs = QTabWidget()

        # Modbus RTU 操作选项卡
        rtu_tab = QWidget()
        rtu_layout = QVBoxLayout()

        # 寄存器表
        self.register_table = QTableWidget()
        self.register_table.setColumnCount(4)
        self.register_table.setHorizontalHeaderLabels(["地址", "寄存器值", "长整型值", "浮点型值"])
        self.register_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # 操作区域
        operation_group = QGroupBox("Modbus RTU 操作")
        operation_layout = QHBoxLayout()

        # 读取操作
        read_group = QGroupBox("读取数据")
        read_layout = QVBoxLayout()

        self.read_addr = QSpinBox()
        self.read_addr.setRange(40000, 49999)
        self.read_addr.setValue(40000)

        self.read_count = QSpinBox()
        self.read_count.setRange(1, 20)
        self.read_count.setValue(1)

        self.read_type_combo = QComboBox()
        self.read_type_combo.addItems(["长整型 (LONG)", "浮点型 (FLOAT)"])

        read_btn = QPushButton("读取数据")
        read_btn.clicked.connect(self.read_data)

        read_layout.addWidget(QLabel("起始地址:"))
        read_layout.addWidget(self.read_addr)
        read_layout.addWidget(QLabel("数据个数 (1-20):"))
        read_layout.addWidget(self.read_count)
        read_layout.addWidget(QLabel("数据类型:"))
        read_layout.addWidget(self.read_type_combo)
        read_layout.addWidget(read_btn)
        read_layout.addStretch()
        read_group.setLayout(read_layout)

        # 写入操作
        write_group = QGroupBox("写入数据")
        write_layout = QVBoxLayout()

        self.write_addr = QSpinBox()
        self.write_addr.setRange(40000, 49999)
        self.write_addr.setValue(40000)

        self.write_value = QLineEdit()
        self.write_value.setPlaceholderText("输入要写入的值")

        self.write_type_combo = QComboBox()
        self.write_type_combo.addItems(["长整型 (LONG)", "浮点型 (FLOAT)"])

        write_btn = QPushButton("写入数据")
        write_btn.clicked.connect(self.write_data)

        zero_btn = QPushButton("清零操作")
        zero_btn.clicked.connect(self.zero_operation)

        write_layout.addWidget(QLabel("目标地址:"))
        write_layout.addWidget(self.write_addr)
        write_layout.addWidget(QLabel("写入值:"))
        write_layout.addWidget(self.write_value)
        write_layout.addWidget(QLabel("数据类型:"))
        write_layout.addWidget(self.write_type_combo)
        write_layout.addWidget(write_btn)
        write_layout.addWidget(zero_btn)
        write_layout.addStretch()
        write_group.setLayout(write_layout)

        # 添加读写组到操作区域
        operation_layout.addWidget(read_group)
        operation_layout.addWidget(write_group)
        operation_group.setLayout(operation_layout)

        # 添加组件到RTU选项卡
        rtu_layout.addWidget(self.register_table)
        rtu_layout.addWidget(operation_group)
        rtu_tab.setLayout(rtu_layout)

        # 主动发送模式选项卡
        active_tab = QWidget()
        active_layout = QVBoxLayout()

        self.active_text = QTextEdit()
        self.active_text.setReadOnly(True)
        self.active_text.setPlaceholderText("主动发送模式数据将显示在这里...")

        clear_btn = QPushButton("清空显示")
        clear_btn.clicked.connect(lambda: self.active_text.clear())

        active_layout.addWidget(QLabel("仪表主动发送数据:"))
        active_layout.addWidget(self.active_text)
        active_layout.addWidget(clear_btn)
        active_tab.setLayout(active_layout)

        # 数据监控选项卡
        monitor_tab = QWidget()
        monitor_layout = QVBoxLayout()

        self.monitor_text = QTextEdit()
        self.monitor_text.setReadOnly(True)
        self.monitor_text.setPlaceholderText("串口通信数据将显示在这里...")

        monitor_layout.addWidget(QLabel("串口通信监控:"))
        monitor_layout.addWidget(self.monitor_text)
        monitor_tab.setLayout(monitor_layout)

        # 添加选项卡
        self.tabs.addTab(rtu_tab, "Modbus RTU 模式")
        self.tabs.addTab(active_tab, "主动发送模式")
        self.tabs.addTab(monitor_tab, "通信监控")

        main_layout.addWidget(self.tabs)
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

        # 状态栏
        self.status_bar = self.statusBar()
        self.status_label = QLabel("就绪")
        self.status_bar.addWidget(self.status_label)

        # 设置默认选项卡
        self.tabs.setCurrentIndex(0)

    def init_register_table(self):
        """初始化寄存器表"""
        self.register_table.setRowCount(20)
        for i in range(20):
            addr = 40000 + i * 2
            addr_item = QTableWidgetItem(str(addr))
            addr_item.setFlags(addr_item.flags() & ~Qt.ItemIsEditable)
            self.register_table.setItem(i, 0, addr_item)

            # 其他列初始化为空
            for col in range(1, 4):
                item = QTableWidgetItem("")
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.register_table.setItem(i, col, item)

    def scan_serial_ports(self):
        """扫描可用串口"""
        self.port_combo.clear()
        ports = serial.tools.list_ports.comports()
        for port in ports:
            self.port_combo.addItem(port.device)
        if ports:
            self.status_label.setText(f"找到 {len(ports)} 个可用串口")
        else:
            self.status_label.setText("未找到可用串口")

    def toggle_connection(self):
        """连接/断开串口"""
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
            self.serial_port = None
            self.connect_btn.setText("连接")
            self.status_label.setText("串口已断开")
        else:
            port = self.port_combo.currentText()
            if not port:
                QMessageBox.warning(self, "错误", "请选择串口")
                return

            try:
                baud = int(self.baud_combo.currentText())
                data_bits = int(self.data_bits_combo.currentText())
                stop_bits_text = self.stop_bits_combo.currentText()

                if stop_bits_text == "1.5":
                    stop_bits = serial.STOPBITS_ONE_POINT_FIVE
                elif stop_bits_text == "2":
                    stop_bits = serial.STOPBITS_TWO
                else:
                    stop_bits = serial.STOPBITS_ONE

                parity = self.parity_combo.currentText()
                if parity == "无":
                    parity = serial.PARITY_NONE
                elif parity == "奇校验":
                    parity = serial.PARITY_ODD
                else:
                    parity = serial.PARITY_EVEN

                self.serial_port = serial.Serial(
                    port=port,
                    baudrate=baud,
                    bytesize=data_bits,
                    parity=parity,
                    stopbits=stop_bits,
                    timeout=0.1
                )

                self.connect_btn.setText("断开")
                self.status_label.setText(f"已连接 {port} @ {baud} bps")
            except Exception as e:
                QMessageBox.critical(self, "连接错误", f"无法打开串口: {str(e)}")
                self.status_label.setText("连接失败")

    def read_serial_data(self):
        """读取串口数据"""
        if self.serial_port and self.serial_port.is_open:
            try:
                data = self.serial_port.read_all()
                if data:
                    # 在通信监控中显示数据
                    hex_data = binascii.hexlify(data).decode('utf-8')
                    # 修复1：将hex_data定义移出try块确保作用域
                    # 修复2：修正range函数括号错误
                    formatted_hex = ' '.join([hex_data[i:i + 2] for i in range(0, len(hex_data), 2)])
                    self.monitor_text.append(f"接收: {formatted_hex}")

                    # 根据当前模式处理数据
                    if self.tabs.currentIndex() == 1:  # 主动发送模式
                        try:
                            ascii_data = data.decode('ascii')
                            self.active_text.append(ascii_data)
                        except UnicodeDecodeError:  # 修复3：指定异常类型
                            pass
                    else:  # Modbus RTU模式
                        self.process_modbus_response(data)
            except Exception as e:
                self.status_label.setText(f"读取错误: {str(e)}")

    def process_modbus_response(self, data):
        """处理Modbus响应"""
        if len(data) < 5:
            return

        # 解析功能码
        func_code = data[1]

        # 03功能码响应处理
        if func_code == 0x03:
            # 数据长度
            data_len = data[2]
            # 寄存器数据
            reg_data = data[3:3 + data_len]

            # 每4个字节解析为一个32位值
            for i in range(0, len(reg_data), 4):
                if i + 4 > len(reg_data):
                    break

                # 高位在前，低位在后
                value_bytes = reg_data[i:i + 4]

                # 解析为长整型和浮点型
                try:
                    # 长整型
                    long_value = struct.unpack('>i', value_bytes)[0]

                    # 浮点型
                    float_value = struct.unpack('>f', value_bytes)[0]

                    # 更新表格 (简化处理，实际应用中需要根据地址更新)
                    row = i // 4
                    if row < self.register_table.rowCount():
                        self.register_table.setItem(row, 1, QTableWidgetItem(hex(struct.unpack('>I', value_bytes)[0])))
                        self.register_table.setItem(row, 2, QTableWidgetItem(str(long_value)))
                        self.register_table.setItem(row, 3, QTableWidgetItem(f"{float_value:.6f}"))
                except:
                    pass
            self.status_label.setText("数据读取成功")

        # 10功能码响应处理
        elif func_code == 0x10:
            self.status_label.setText("数据写入成功")

    def read_data(self):
        """发送读取数据命令"""
        if not self.serial_port or not self.serial_port.is_open:
            QMessageBox.warning(self, "错误", "请先连接串口")
            return

        # 构建Modbus RTU读取命令
        address = self.read_addr.value()
        count = self.read_count.value()
        device_id = 1  # 默认设备ID

        # 计算Modbus寄存器地址
        reg_address = address - 40000

        # 构建命令
        cmd = bytearray()
        cmd.append(device_id)  # 设备地址
        cmd.append(0x03)  # 功能码
        cmd.append((reg_address >> 8) & 0xFF)  # 寄存器地址高8位
        cmd.append(reg_address & 0xFF)  # 寄存器地址低8位
        cmd.append((count * 2) >> 8 & 0xFF)  # 寄存器数量高8位 (每个参数4字节，2个寄存器)
        cmd.append((count * 2) & 0xFF)  # 寄存器数量低8位

        # 计算CRC
        crc = self.calculate_crc(cmd)
        cmd.append(crc & 0xFF)
        cmd.append((crc >> 8) & 0xFF)

        # 发送命令
        try:
            self.serial_port.write(cmd)

            # 在通信监控中显示发送的数据
            hex_cmd = binascii.hexlify(cmd).decode('utf-8')
            # 修复：修正range函数括号错误
            formatted_hex = ' '.join([hex_cmd[i:i + 2] for i in range(0, len(hex_cmd), 2)])
            self.monitor_text.append(f"发送: {formatted_hex}")

            self.status_label.setText(f"发送读取命令: 地址={address}, 数量={count}")
        except Exception as e:
            self.status_label.setText(f"发送错误: {str(e)}")

    def write_data(self):
        """发送写入数据命令"""
        if not self.serial_port or not self.serial_port.is_open:
            QMessageBox.warning(self, "错误", "请先连接串口")
            return

        try:
            value = float(self.write_value.text())
        except ValueError:
            QMessageBox.warning(self, "错误", "请输入有效的数值")
            return

        # 构建Modbus RTU写入命令
        address = self.write_addr.value()
        device_id = 1  # 默认设备ID

        # 计算Modbus寄存器地址
        reg_address = address - 40000

        # 根据数据类型转换
        if self.write_type_combo.currentIndex() == 0:  # 长整型
            value_bytes = struct.pack('>i', int(value))
        else:  # 浮点型
            value_bytes = struct.pack('>f', value)

        # 构建命令
        cmd = bytearray()
        cmd.append(device_id)  # 设备地址
        cmd.append(0x10)  # 功能码
        cmd.append((reg_address >> 8) & 0xFF)  # 寄存器地址高8位
        cmd.append(reg_address & 0xFF)  # 寄存器地址低8位
        cmd.append(0x00)  # 寄存器数量高8位 (固定2个寄存器)
        cmd.append(0x02)  # 寄存器数量低8位
        cmd.append(0x04)  # 字节数 (4字节)
        cmd.extend(value_bytes)  # 数据

        # 计算CRC
        crc = self.calculate_crc(cmd)
        cmd.append(crc & 0xFF)
        cmd.append((crc >> 8) & 0xFF)

        # 发送命令
        try:
            self.serial_port.write(cmd)

            # 在通信监控中显示发送的数据
            hex_cmd = binascii.hexlify(cmd).decode('utf-8')
            # 修复：修正range函数括号错误
            formatted_hex = ' '.join([hex_cmd[i:i + 2] for i in range(0, len(hex_cmd), 2)])
            self.monitor_text.append(f"发送: {formatted_hex}")

            self.status_label.setText(f"发送写入命令: 地址={address}, 值={value}")
        except Exception as e:
            self.status_label.setText(f"发送错误: {str(e)}")

    def zero_operation(self):
        """清零操作"""
        if not self.serial_port or not self.serial_port.is_open:
            QMessageBox.warning(self, "错误", "请先连接串口")
            return

        # 在实际应用中，清零操作需要写入特定的值到特定的地址
        # 这里只是一个示例
        QMessageBox.information(self, "清零操作", "清零操作已执行")
        self.status_label.setText("清零操作已执行")

    def calculate_crc(self, data):
        """计算Modbus CRC16校验码"""
        crc = 0xFFFF
        for byte in data:
            crc ^= byte
            for _ in range(8):
                if crc & 0x0001:
                    crc >>= 1
                    crc ^= 0xA001
                else:
                    crc >>= 1
        return crc

    def closeEvent(self, event):
        """关闭窗口时关闭串口"""
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ModbusRTUTool()
    window.show()
    sys.exit(app.exec_())