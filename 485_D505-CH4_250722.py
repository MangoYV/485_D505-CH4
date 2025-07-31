import sys
import struct
import serial
import serial.tools.list_ports
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGroupBox,
                             QLabel, QComboBox, QLineEdit, QPushButton, QTextEdit, QTabWidget,
                             QGridLayout, QMessageBox, QCheckBox)
from PyQt5.QtCore import QTimer


class ModbusRTUTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("D505-CH4四通道测力显示仪通信工具 - Modbus-RTU")
        self.setGeometry(900, 240, 770, 610)
        self.serial_port = None
        self.init_ui()

    def init_ui(self):
        # 创建主布局
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

        # 创建选项卡
        tabs = QTabWidget()
        main_layout.addWidget(tabs)

        # 串口设置选项卡
        settings_tab = QWidget()
        settings_layout = QVBoxLayout(settings_tab)

        # 串口设置组
        port_group = QGroupBox("串口设置")
        port_layout = QGridLayout()

        port_layout.addWidget(QLabel("串口号:"), 0, 0)
        self.port_combo = QComboBox()
        self.port_combo.setMinimumWidth(100)  # 增加最小宽度
        self.refresh_ports()  # 刷新端口列表
        # 设置默认选择COM3
        com3_index = self.port_combo.findText("COM3")
        if com3_index >= 0:
            self.port_combo.setCurrentIndex(com3_index)
        port_layout.addWidget(self.port_combo, 0, 1)  # 确保添加到布局中

        self.refresh_button = QPushButton("刷新端口")
        self.refresh_button.clicked.connect(self.refresh_ports)  # 绑定刷新事件
        port_layout.addWidget(self.refresh_button, 0, 2)

        # 设置列拉伸因子确保空间分配合理
        port_layout.setColumnStretch(1, 1)  # 给下拉框列分配更多空间

        port_layout.addWidget(QLabel("波特率:"), 1, 0)
        self.baud_combo = QComboBox()
        self.baud_combo.setMinimumWidth(100)  # 增加最小宽度
        self.baud_combo.addItems(["9600", "19200", "38400", "57600", "115200"])
        self.baud_combo.setCurrentText("115200")
        port_layout.addWidget(self.baud_combo, 1, 1)

        port_layout.addWidget(QLabel("数据位:"), 2, 0)
        self.data_bits_combo = QComboBox()
        self.data_bits_combo.setMinimumWidth(100)  # 增加最小宽度
        self.data_bits_combo.addItems(["5", "6", "7", "8"])
        self.data_bits_combo.setCurrentText("8")
        port_layout.addWidget(self.data_bits_combo, 2, 1)

        port_layout.addWidget(QLabel("停止位:"), 3, 0)
        self.stop_bits_combo = QComboBox()
        self.stop_bits_combo.setMinimumWidth(100)  # 增加最小宽度
        self.stop_bits_combo.addItems(["1", "1.5", "2"])
        self.stop_bits_combo.setCurrentText("1")
        port_layout.addWidget(self.stop_bits_combo, 3, 1)

        port_layout.addWidget(QLabel("校验位:"), 4, 0)
        self.parity_combo = QComboBox()
        self.parity_combo.setMinimumWidth(100)  # 增加最小宽度
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

        # 操作说明组
        info_group = QGroupBox("操作说明")
        info_layout = QVBoxLayout()
        info_text = QTextEdit()
        info_text.setReadOnly(True)
        info_text.setPlainText(
            "1. 默认串口设置: 115200波特率, 8数据位, 1停止位, 无校验\n"
            "2. 从站地址范围: 1-247\n"
            "3. 寄存器地址: 0-65535 (32位数据占用2个连续寄存器)\n"
            "4. 读取数据: 选择起始地址和读取寄存器数量(必须为2的倍数)\n"
            "5. 写入数据: 输入32位数据(长整型或浮点型)\n"
            "6. 修改通讯参数后需重新上电生效"
        )
        info_layout.addWidget(info_text)
        info_group.setLayout(info_layout)
        settings_layout.addWidget(info_group)

        # 数据操作选项卡
        data_tab = QWidget()
        data_layout = QVBoxLayout(data_tab)

        # 创建水平布局容器用于放置读取和写入操作组
        read_write_layout = QHBoxLayout()

        # 读取操作组
        read_group = QGroupBox("读取操作")
        read_layout = QGridLayout()

        # 创建8个独立的地址输入框
        self.read_address_edits = []  # 存储8个地址输入框
        for i in range(8):
            row = i // 2  # 每行放置两个输入框
            col = (i % 2) * 2  # 列位置

            # 添加标签
            label = QLabel(f"地址{i + 1}:")
            read_layout.addWidget(label, row, col)

            # 添加输入框
            # 修改默认值设置：前4个地址保持递增，后4个默认设为0
            if i < 4:
                default_address = str(2000 + i * 2)  # 前4个：2000,2002,2004,2006
            else:
                default_address = "0"  # 后4个默认0

            address_edit = QLineEdit(default_address)
            address_edit.setValidator(self.create_int_validator(0, 65535))
            read_layout.addWidget(address_edit, row, col + 1)
            self.read_address_edits.append(address_edit)

        # 缩放因子设置和数据类型选择放在同一行
        read_layout.addWidget(QLabel("数据缩放:"), 4, 0)
        self.scale_factor_edit = QLineEdit("0.1")  # 默认缩放因子0.1（10倍）
        self.scale_factor_edit.setValidator(self.create_float_validator())
        read_layout.addWidget(self.scale_factor_edit, 4, 1)

        # 将数据类型选择框移动到缩放因子同一行
        read_layout.addWidget(QLabel("数据类型:"), 4, 2)
        self.data_type_combo = QComboBox()
        self.data_type_combo.addItems(["浮点型 (Float)", "长整型 (Long)"])
        self.data_type_combo.setCurrentIndex(1)  # 默认选择长整型
        read_layout.addWidget(self.data_type_combo, 4, 3)

        self.read_button = QPushButton("读取数据")
        self.read_button.clicked.connect(self.read_data)  # 添加事件绑定
        read_layout.addWidget(self.read_button, 5, 0, 1, 4)  # 按钮跨4列

        read_group.setLayout(read_layout)
        read_write_layout.addWidget(read_group, 1)  # 添加到水平布局

        # 写入操作组
        write_group = QGroupBox("写入操作")
        write_layout = QGridLayout()

        write_layout.addWidget(QLabel("起始地址:"), 0, 0)
        self.write_address_edit = QLineEdit("2000")  # 默认起始地址改为2000
        self.write_address_edit.setValidator(self.create_int_validator(0, 65535))
        write_layout.addWidget(self.write_address_edit, 0, 1)

        write_layout.addWidget(QLabel("写入值:"), 1, 0)
        self.write_value_edit = QLineEdit("0.0")
        write_layout.addWidget(self.write_value_edit, 1, 1)

        write_layout.addWidget(QLabel("数据类型:"), 2, 0)
        self.write_type_combo = QComboBox()
        self.write_type_combo.addItems(["浮点型 (Float)", "长整型 (Long)"])
        self.write_type_combo.setCurrentIndex(1)  # 默认选择长整型
        write_layout.addWidget(self.write_type_combo, 2, 1)

        self.write_button = QPushButton("写入数据")
        self.write_button.clicked.connect(self.write_data)
        write_layout.addWidget(self.write_button, 3, 0, 1, 2)

        write_group.setLayout(write_layout)
        read_write_layout.addWidget(write_group, 1)  # 添加到水平布局

        # 将水平布局添加到主垂直布局
        data_layout.addLayout(read_write_layout)

        # 结果输出组 - 拆分为左右两部分
        result_group = QGroupBox("通信结果")
        result_layout = QHBoxLayout()  # 改为水平布局
        result_layout.setContentsMargins(5, 15, 5, 5)

        # 左侧：通信命令区域
        comm_group = QGroupBox("通信命令")
        comm_layout = QVBoxLayout()
        self.comm_text = QTextEdit()
        self.comm_text.setReadOnly(True)
        comm_layout.addWidget(self.comm_text)
        comm_group.setLayout(comm_layout)
        result_layout.addWidget(comm_group, 1)  # 拉伸因子设为1

        # 右侧：读取结果区域
        result_display_group = QGroupBox("读取结果")
        result_display_layout = QVBoxLayout()
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        result_display_layout.addWidget(self.result_text)
        result_display_group.setLayout(result_display_layout)
        result_layout.addWidget(result_display_group, 1)  # 拉伸因子设为1

        result_group.setLayout(result_layout)
        data_layout.addWidget(result_group, 1)

        # 清空按钮
        clear_layout = QHBoxLayout()
        self.clear_button = QPushButton("清空结果")
        self.clear_button.clicked.connect(self.clear_results)
        clear_layout.addStretch(1)
        clear_layout.addWidget(self.clear_button)
        clear_layout.addStretch(1)
        data_layout.addLayout(clear_layout)

        # 添加选项卡
        tabs.addTab(settings_tab, "串口设置")
        tabs.addTab(data_tab, "数据操作")

        # 状态栏
        self.status_bar = self.statusBar()
        self.status_bar.showMessage("准备就绪")

        # 初始化串口状态
        self.serial_connected = False

    def create_int_validator(self, min_val, max_val):
        from PyQt5.QtGui import QIntValidator
        validator = QIntValidator()
        validator.setRange(min_val, max_val)
        return validator

    # 新增：创建浮点数验证器
    def create_float_validator(self):
        from PyQt5.QtGui import QDoubleValidator
        validator = QDoubleValidator()
        validator.setBottom(0.0001)  # 最小缩放因子0.0001
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
            self.comm_text.append(f"串口已打开: {port}, {baudrate}波特率")  # 显示在通信区域
        except Exception as e:
            QMessageBox.critical(self, "连接错误", f"无法打开串口: {str(e)}")
            self.status_bar.showMessage(f"连接失败: {str(e)}")

    def close_serial(self):
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
        self.serial_connected = False
        self.connect_button.setText("打开串口")
        self.status_bar.showMessage("串口已关闭")
        self.comm_text.append("串口已关闭")  # 显示在通信区域

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
            results = []

            # 添加分隔线区分不同次的读取命令
            self.comm_text.append("----------------------")

            # 分别读取8个地址
            for i, address_edit in enumerate(self.read_address_edits):
                start_address = int(address_edit.text())

                # 新增：如果地址为0则跳过读取
                if start_address == 0:
                    continue

                # 构建读取指令 (Modbus功能码03)
                command = bytearray()
                command.append(slave_address)  # 从站地址
                command.append(0x03)  # 功能码: 读保持寄存器
                command.append((start_address >> 8) & 0xFF)  # 起始地址高字节
                command.append(start_address & 0xFF)  # 起始地址低字节
                command.append(0x00)  # 寄存器数量高字节 (固定读取2个寄存器)
                command.append(0x02)  # 寄存器数量低字节 (固定读取2个寄存器)

                # 计算CRC并添加到命令
                crc = self.calculate_crc(command)
                command.extend(crc)

                # 发送命令
                self.serial_port.write(command)
                self.comm_text.append(f"发送读取命令(地址{start_address}): {command.hex(' ').upper()}")  # 显示在通信区域

                # 接收响应
                expected_length = 9  # 响应长度固定为9字节
                response = self.serial_port.read(expected_length)

                if not response:
                    results.append(f"地址{start_address}: 读取超时，未收到响应")
                    continue

                self.comm_text.append(f"收到响应数据(地址{start_address}): {response.hex(' ').upper()}")  # 显示在通信区域

                # 验证响应长度
                if len(response) < 5:
                    results.append(f"地址{start_address}: 响应长度不足")
                    continue

                # 验证CRC
                received_crc = response[-2:]
                calculated_crc = self.calculate_crc(response[:-2])
                if received_crc != calculated_crc:
                    results.append(f"地址{start_address}: CRC校验失败")
                    continue

                # 验证从站地址和功能码
                if response[0] != slave_address:
                    results.append(f"地址{start_address}: 从站地址不匹配")
                    continue

                if response[1] != 0x03:
                    results.append(f"地址{start_address}: 功能码错误")
                    continue

                # 获取数据字节数
                byte_count = response[2]
                data_bytes = response[3:3 + byte_count]

                # 解析数据
                if "浮点型" in data_type:
                    try:
                        value = struct.unpack('>f', data_bytes)[0]
                        scaled_value = value * scale_factor
                        results.append(f"地址{start_address}: {scaled_value:.2f}")
                    except:
                        results.append(f"地址{start_address}: 解析浮点数错误")
                else:  # 长整型
                    try:
                        value = (data_bytes[0] << 24) | (data_bytes[1] << 16) | (data_bytes[2] << 8) | data_bytes[3]
                        scaled_value = value * scale_factor
                        if scale_factor == 1:
                            results.append(f"地址{start_address}: {int(scaled_value)}")
                        else:
                            results.append(f"地址{start_address}: {scaled_value:.2f}")
                    except:
                        results.append(f"地址{start_address}: 解析长整型错误")

            # 显示结果（在右侧的读取结果区域）
            # 修改：添加标题行"读取结果："
            if results:  # 确保有结果时才添加标题
                self.result_text.append("读取结果：")  # 添加标题行
                for result in results:
                    self.result_text.append(result)
                self.result_text.append("")  # 添加空行分隔不同次读取
            else:
                self.result_text.append("读取结果：无有效数据")
                self.result_text.append("")  # 添加空行分隔不同次读取

        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取数据时发生错误: {str(e)}")
            self.result_text.append(f"错误: {str(e)}")

    def write_data(self):
        if not self.serial_connected:
            QMessageBox.warning(self, "错误", "请先打开串口")
            return

        try:
            slave_address = int(self.slave_address_edit.text())
            start_address = int(self.write_address_edit.text())
            data_type = self.write_type_combo.currentText()
            value_str = self.write_value_edit.text()

            # 根据数据类型转换值
            if "浮点型" in data_type:
                try:
                    value = float(value_str)
                    value_bytes = struct.pack('>f', value)
                except ValueError:
                    QMessageBox.warning(self, "错误", "无效的浮点数值")
                    return
            else:  # 长整型
                try:
                    value = int(value_str)
                    # 将32位整数转换为4字节（大端）
                    value_bytes = bytes([
                        (value >> 24) & 0xFF,
                        (value >> 16) & 0xFF,
                        (value >> 8) & 0xFF,
                        value & 0xFF
                    ])
                except ValueError:
                    QMessageBox.warning(self, "错误", "无效的长整型值")
                    return

            # 构建写入指令 (Modbus功能码16)
            # 格式: [地址(1)][功能码(1)][起始地址(2)][寄存器数量(2)][字节数(1)][数据(4)][CRC(2)]
            register_count = 2  # 32位数据占用2个寄存器
            byte_count = 4  # 4字节数据

            command = bytearray()
            command.append(slave_address)  # 从站地址
            command.append(0x10)  # 功能码: 写多个寄存器
            command.append((start_address >> 8) & 0xFF)  # 起始地址高字节
            command.append(start_address & 0xFF)  # 起始地址低字节
            command.append((register_count >> 8) & 0xFF)  # 寄存器数量高字节
            command.append(register_count & 0xFF)  # 寄存器数量低字节
            command.append(byte_count)  # 字节数
            command.extend(value_bytes)  # 数据

            # 计算CRC并添加到命令
            crc = self.calculate_crc(command)
            command.extend(crc)

            # 发送命令
            self.serial_port.write(command)
            self.comm_text.append(f"发送写入命令: {command.hex(' ').upper()}")  # 显示在通信区域

            # 接收响应
            # 响应格式: [地址(1)][功能码(1)][起始地址(2)][寄存器数量(2)][CRC(2)]
            expected_length = 8
            response = self.serial_port.read(expected_length)

            if not response:
                self.comm_text.append("写入超时，未收到响应")  # 显示在通信区域
                return

            self.comm_text.append(f"收到响应: {response.hex(' ').upper()}")  # 显示在通信区域

            # 验证响应长度
            if len(response) < 6:
                self.comm_text.append("响应长度不足")  # 显示在通信区域
                return

            # 验证CRC
            received_crc = response[-2:]
            calculated_crc = self.calculate_crc(response[:-2])
            if received_crc != calculated_crc:
                self.comm_text.append(f"CRC校验失败: 收到 {received_crc.hex()} 计算 {calculated_crc.hex()}")  # 显示在通信区域
                return

            # 验证从站地址和功能码
            if response[0] != slave_address:
                self.comm_text.append(f"从站地址不匹配: 收到 {response[0]}, 期望 {slave_address}")  # 显示在通信区域
                return

            if response[1] != 0x10:
                self.comm_text.append(f"功能码错误: 收到 {hex(response[1])}, 期望 0x10")  # 显示在通信区域
                return

            # 验证写入地址和寄存器数量
            resp_start_address = (response[2] << 8) | response[3]
            resp_register_count = (response[4] << 8) | response[5]

            if resp_start_address != start_address:
                self.comm_text.append(f"写入地址不匹配: 收到 {resp_start_address}, 期望 {start_address}")  # 显示在通信区域
                return

            if resp_register_count != register_count:
                self.comm_text.append(f"寄存器数量不匹配: 收到 {resp_register_count}, 期望 {register_count}")  # 显示在通信区域
                return

            self.comm_text.append(f"写入成功: 地址 {start_address}, 值 {value_str}")  # 显示在通信区域

        except Exception as e:
            QMessageBox.critical(self, "错误", f"写入数据时发生错误: {str(e)}")
            self.result_text.append(f"错误: {str(e)}")

    def clear_results(self):
        self.comm_text.clear()  # 清空通信区域
        self.result_text.clear()  # 清空读取结果区域

    def closeEvent(self, event):
        self.close_serial()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ModbusRTUTool()
    window.show()
    sys.exit(app.exec_())