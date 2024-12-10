import serial
import serial.tools.list_ports
resolution = 4095
VRef = 4.0

ports = serial.tools.list_ports.comports()
for port, desc, hwid in sorted(ports):
    if "STM" in desc:
        stm32_port = port
        print(f"STM32 port is {stm32_port}.\nMicrocontroller Connected")
        break

while True:
    with serial.Serial(stm32_port, 115200) as s:
        response = s.readline().strip()
        try:
            adc_val = int(response.decode('utf-8', errors='ignore'))
            voltage = (VRef * adc_val / resolution) - 2
            print(f"{voltage:.3f}")
        except ValueError:
            None
