import serial
import serial.tools.list_ports
import threading
import pandas as pd
from tkinter import *
from tkinter import filedialog

# Global variables
read_enable = False
VRef = 4.0
resolution = 4095
stm32_port = None
baud_rate = 115200
sensor_thread = None

def main():
    global gui, text_widget, start_btn
    
    # Initialize the GUI
    gui = Tk()
    gui.title("Sensor Reader")
    gui.geometry("400x600")
    
    # Create the text display area
    frame = Frame(gui)
    frame.pack(pady=20)

    text_widget = Text(frame, width=40, height=15)
    text_widget.pack(side=LEFT, fill=BOTH, expand=True)

    scrollbar = Scrollbar(frame, command=text_widget.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    text_widget.config(yscrollcommand=scrollbar.set)

    # Create the buttons
    start_btn = Button(gui, text="Click to Start Sensor Read", width=40, command=toggle_enable)
    start_btn.pack(pady=10)

    clear_btn = Button(gui, text="Clear Data", width=40, command=clear_text_widget)
    clear_btn.pack(pady=10)

    export_saveas_btn = Button(gui, text="Save As .xlsx file", width=40, command=export_saveas_excel)
    export_saveas_btn.pack(pady=10)

    export_openas_btn = Button(gui, text="Open and Export to existing .xlsx file", width=40, command=export_openas_excel)
    export_openas_btn.pack(pady=10)

    connect_btn = Button(gui, text="Connect Microcontroller", width=40, command=connect_mcu)
    connect_btn.pack(pady=10)

    exit_button = Button(gui, text="Exit", width=40, command=gui.destroy)
    exit_button.pack(pady=10)

    # Check for initial STM32 connection
    initialize_stm32_connection()
    display_message("Voltages Displayed May be Delayed")
    # Run the GUI main loop
    gui.mainloop()

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

def read_sensor_data():
    global read_enable
    if stm32_port is None:
        display_message("No Device Found")
        return
    try:
        with serial.Serial(stm32_port, baud_rate) as s, open("voltage_data.txt", "a") as file:
            display_message(f"Serial port {stm32_port} is open.")
            while read_enable:
                if s.in_waiting:
                    response = s.readline().strip()
                    try:
                        adc_val = int(response.decode('utf-8', errors='ignore'))
                        voltage = (VRef * adc_val / resolution) - 2
                        file.write(f"{voltage:.3f}\n")
                        display_message(f"Voltage = {voltage:.3f} V")
                    except ValueError:
                        display_message(f"Invalid data received: {response}")
            
            if s.in_waiting:
                display_message("Stop Request Receieved. Data logging still in progress")
                display_message("Do not export or save until data logging is complete")
                while s.in_waiting:
                    try:
                        adc_val = int(response.decode('utf-8', errors='ignore'))
                        voltage = (VRef * adc_val / resolution) - 2
                        file.write(f"{voltage:.3f}\n")
                        #display_message(f"Voltage = {voltage:.3f} V")
                    except ValueError:
                        display_message(f"Invalid data received: {response}")
                display_message("Data Logging is Complete")   
    except serial.SerialException as e:
        display_message(f"Serial error: {e}")

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

        # Read data from text file
        with open("voltage_data.txt", "r") as file:
            data = [float(line.strip()) for line in file if line.strip()]

        # Create a DataFrame with proper index
        df = pd.DataFrame(data, columns=["Voltage (V)"])

        # Export to Excel
        df.to_excel(file_path, index=True)  # index=True keeps the natural indices
        display_message(f"Data exported to {file_path}")
    except Exception as e:
        display_message(f"An error occurred: {e}")

def export_openas_excel():
    try:
        file_path = filedialog.askopenfilename(
            title="Select an Existing Excel file",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not file_path:
            display_message("Export cancelled.")
            return

        # Read data from text file
        with open("voltage_data.txt", "r") as file:
            data = [float(line.strip()) for line in file if line.strip()]

        # Open the existing Excel file and append data
        existing_df = pd.read_excel(file_path, engine='openpyxl', index_col=0)
        new_df = pd.DataFrame(data, columns=["Voltage (V)"])
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)

        # Export back to Excel
        combined_df.to_excel(file_path, index=True)  # index=True keeps unique indices
        display_message(f"Data appended and exported to {file_path}")
    except Exception as e:
        display_message(f"An error occurred: {e}")

def connect_mcu():
    global stm32_port
    ports = serial.tools.list_ports.comports()
    if stm32_port is None:
        for port, desc, hwid in sorted(ports):
            if "STM" in desc:
                stm32_port = port
                display_message(f"STM32 port is {stm32_port}.\nMicrocontroller Connected")
                break
        if stm32_port is None:
            display_message("Microcontroller Not Found")
    else:
        display_message("Microcontroller Already Connected")

def initialize_stm32_connection():
    global stm32_port
    ports = serial.tools.list_ports.comports()
    for port, desc, hwid in sorted(ports):
        if "STM" in desc:
            stm32_port = port
            display_message(f"STM32 port is {stm32_port}.\nMicrocontroller Connected")
            break
    else:
        display_message("Microcontroller Not Found")
    with open("voltage_data.txt", "w"):
        pass

if __name__ == "__main__":
    main()
