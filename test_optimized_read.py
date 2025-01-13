import serial
import serial.tools.list_ports
import threading
import pandas as pd
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import queue
import time

read_enable = False

VRef = 4.0
resolution = 2**12
stm32_port = None
baud_rate = 115200
conversion_factor = VRef / resolution
voltage_queue = queue.Queue()

def read_sensor_data():
    if stm32_port is None:
        display_message("No Device Found")
        return
    try:
        while True:
            if read_enable is True:
                with serial.Serial(stm32_port, baud_rate) as s:
                    while not exit_flag:
                        read_enable.wait()
                        if s.in_waiting:
                            response = s.readline().strip()
                            try:
                                adc_val = int(response.decode('utf-8', errors='ignore'))
                                voltage = (conversion_factor * adc_val) - 2
                                voltage_queue.put(voltage)

                                if(voltage_queue.qsize() > 99):
                                    display_message(f"{voltage_queue.get()} V")
                                    with open("voltage_data.txt", "a") as file:
                                        file.write("\n".join(map(str, list(voltage_queue.queue))) + "\n")
                                    while not voltage_queue.empty():
                                        voltage_queue.get_nowait()
                            except ValueError:
                                None

                        if not read_enable.is_set():
                            remaining_data = s.read_all().decode('utf-8', errors='ignore').splitlines()
                            for response in remaining_data:
                                try:
                                    adc_val = int(response)
                                    voltage = (VRef * adc_val / resolution) - 2
                                    voltage_queue.put(voltage)

                                    with open("voltage_data.txt", "a") as file:
                                        file.write("\n".join(map(str, list(voltage_queue.queue))) + "\n")
                                    while not voltage_queue.empty():
                                        voltage_queue.get_nowait()
                                except ValueError:
                                    display_message(f"Invalid data received: {response}")
            else:
                time.sleep(1)
    except serial.SerialException as e:
        display_message(f"Serial error: {e}")

def main(args=None):

    global read_enable, sensor_thread, text_widget, start_btn, exit_flag

    #Gui init
    root = Tk()
    root.title("Animal Bite Force Datastream")
    root.geometry("600x600")

    # Create the text display area
    frame = Frame(root)
    frame.pack(pady=20)

    text_widget = Text(frame, width=60, height=15)
    text_widget.pack(side=LEFT, fill=BOTH, expand=True)

    scrollbar = Scrollbar(frame, command=text_widget.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    text_widget.config(yscrollcommand=scrollbar.set)

    # Dropdown menu for COM Port Selection

    ports = serial.tools.list_ports.comports()
    port_list = [port.device for port in ports]

    if not port_list:
        display_message("No COM Ports Found")

    port_var = StringVar()
    port_var.set(port_list[0])
    dropdown = ttk.Combobox(root, textvariable=port_var, values=port_list, state="readonly")
    dropdown.pack(pady=10)


    # Buttons
    start_btn = Button(root, text="Click to Start Sensor Read", width=40, command=toggle_enable)
    start_btn.pack(pady=10)

    clear_btn = Button(root, text="Clear Data", width=40, command=clear_text_widget)
    clear_btn.pack(pady=10)

    export_saveas_btn = Button(root, text="Save Data As .xlsx file", width=40, command=export_saveas_excel)
    export_saveas_btn.pack(pady=10)

    connect_btn = Button(root, text="Connect Microcontroller", width=40, command=lambda: connect_mcu(port_var))
    connect_btn.pack(pady=10)

    exit_button = Button(root, text="Exit", width=40, command=root.destroy)
    exit_button.pack(pady=10)
    root.mainloop()

def toggle_enable():
    global read_enable, sensor_thread

    read_enable = not read_enable
    if read_enable:
        start_btn.config(text="Reading Sensor Data. Click to Stop Sensor Read")
        if not (sensor_thread and sensor_thread.is_alive()):
            sensor_thread = threading.Thread(target=read_sensor_data, daemon=True)
            sensor_thread.start()
    else:
        start_btn.config(text="Click to Start Sensor Read")

def connect_mcu(port_var):
    global stm32_port

    if sensor_thread.is_alive():  
        exit_flag = True
        sensor_thread.join()

    stm32_port = port_var.get()
    if stm32_port:
        display_message(f"Selected port: {stm32_port}\nMicrocontroller Connected")
    else:
        display_message("No port selected")

    sensor_thread.start()

def display_message(message):
    text_widget.insert(END, f"{message}\n")
    text_widget.yview(END)

def clear_text_widget():
    with open("voltage_data.txt", "w"):
        pass
    text_widget.delete(1.0, END)

def export_saveas_excel():
    try:
        file_path = filedialog.asksaveasfilename(
            title="Select or Create an Excel file",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not file_path:
            display_message("Export cancelled.")
            return

        with open("voltage_data.txt", "r") as file:
            data = [float(line.strip()) for line in file if line.strip()]

        df = pd.DataFrame(data, columns=["Voltage (V)"])
        df.to_excel(file_path, index=True)
        display_message(f"Data exported to {file_path}")
    except Exception as e:
        display_message(f"An error occurred: {e}")

if __name__ == '__main__':
    main()
