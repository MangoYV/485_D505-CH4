import time
import struct
import threading
from pymodbus.client import ModbusSerialClient as ModbusClient

# 检查是否安装了 pyserial
try:
    import serial
except ImportError:
    raise RuntimeError("请先安装 pyserial 库，运行命令：pip install pyserial")


class ForceMeterReader:
    def __init__(self, port, slave_address=0x01):
        """
        初始化称重仪表连接
        :param port: 串口号 (如 'COM3' 或 '/dev/ttyUSB0')
        :param slave_address: 仪表地址 (默认0x01)
        """
        self.slave_address = slave_address  # 保存从站地址
        self.client = ModbusClient(
            port=port,
            baudrate=115200,
            bytesize=8,
            parity='N',
            stopbits=1,
            timeout=1
        )
        if not self.client.connect():
            raise ConnectionError(f"无法连接到端口 {port}")

    def read_32bit_value(self, register_address):
        """
        读取32位寄存器值 (自动处理高低位转换)
        :param register_address: 寄存器起始地址 (如 0x0010)
        :return: 解析后的浮点数或长整型
        """
        response = self.client.read_holding_registers(
            address=register_address,
            count=2,
            slave=self.slave_address
        )

        if not response.isError():
            # 组合高低位寄存器值 (高位在前，低位在后)
            combined_value = (response.registers[0] << 16) | response.registers[1]

            # 尝试解析为浮点数 (IEEE 754格式)
            try:
                # 使用struct模块将32位整数转换为浮点数
                byte_data = combined_value.to_bytes(4, 'big')  # 高位在前模式
                float_value = struct.unpack('>f', byte_data)[0]
                return round(float_value, 4)  # 保留4位小数
            except Exception as e:
                print(f"解析为浮点数失败: {e}")
                # 浮点解析失败则返回长整型
                return combined_value
        else:
            raise Exception(f"读取错误: {response}")

    def write_32bit_value(self, register_address, value):
        """
        写入32位值到寄存器
        :param register_address: 寄存器起始地址
        :param value: 要写入的数值 (支持浮点/长整型)
        """
        # 转换浮点数到字节数据
        if isinstance(value, float):
            byte_data = struct.pack('>f', value)
            int_value = struct.unpack('>I', byte_data)[0]
        else:  # 长整型
            int_value = value

        # 拆分32位值为高低位寄存器值
        high_byte = (int_value >> 16) & 0xFFFF
        low_byte = int_value & 0xFFFF

        response = self.client.write_registers(
            address=register_address,
            values=[high_byte, low_byte],
            slave=self.slave_address
        )

        if not response.isError():
            return True
        else:
            raise Exception(f"写入失败: {response}")

    def close(self):
        """关闭连接"""
        self.client.close()


# 使用示例
if __name__ == "__main__":
    # 1. 初始化仪表连接 (修改为实际串口)
    meter = None
    try:
        print("正在连接到力传感器仪表...")
        meter = ForceMeterReader(port='COM3')  # 示例端口
        print("连接成功!")
        
        # 只在首次连接时设置报警值
        print("设置报警值...")
        meter.write_32bit_value(0x0014, 500.0)  # AL1第一报警值地址
        print("报警值设置成功")
        
        # 2. 循环读取重量值
        print("开始读取重量数据 (按 Ctrl+C 退出):")
        try:
            count = 0
            while True:
                weight = meter.read_32bit_value(0x0010)  # ALV给定值地址
                print(f"当前重量读数: {weight}")
                
                # 每5次读取显示一次报警值（可选）
                count += 1
                if count % 5 == 0:
                    alarm = meter.read_32bit_value(0x0014)
                    print(f"当前报警值: {alarm}")
                
                # 等待1秒
                time.sleep(1)
        except KeyboardInterrupt:
            print("\n检测到 Ctrl+C，正在退出...")

    except ConnectionError as ce:
        print(f"连接错误: {ce}")
        print("请检查以下几点:")
        print("1. 确保串口连接正确")
        print("2. 确认串口号是否正确 (当前使用的是 COM3)")
        print("3. 检查设备是否已连接并开启")
        print("4. 确保没有其他程序占用该串口")
    except KeyboardInterrupt:
        print("\n用户中断: 程序被 Ctrl+C 终止")
        raise  # 重新抛出异常，让finally块可以执行
    except Exception as e:
        print(f"操作错误: {e}")
    finally:
        if meter:
            meter.close()
            print("连接已关闭")