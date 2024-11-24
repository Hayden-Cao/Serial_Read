import serial
import serial.tools.list_ports
import threading
import pandas as pd
from tkinter import *
from tkinter import filedialog

ports = serial.tools.list_ports.comports()

# Global variables
read_enable = False
VRef = 4.0
resolution = 4095
stm32_port = None
baud_rate = 115200

# Create the GUI window
gui = Tk()
gui.title("Sensor Reader")
gui.geometry("400x600")

# Create a frame to hold the Text widget and scrollbar
frame = Frame(gui)
frame.pack(pady=20)

# Create a Text widget to display the file contents
text_widget = Text(frame, width=40, height=15)
text_widget.pack(side=LEFT, fill=BOTH, expand=True)

# Create a vertical scrollbar for the Text widget
scrollbar = Scrollbar(frame, command=text_widget.yview)
scrollbar.pack(side=RIGHT, fill=Y)

# Configure the Text widget to use the scrollbar
text_widget.config(yscrollcommand=scrollbar.set)

# Initialize a list to store voltage readings
voltage_data = []


sensor_thread = None

def toggle_enable():
    global read_enable, sensor_thread

    read_enable = not read_enable
    if read_enable:
        start_btn.config(text="Reading Sensor Data. Click to Stop Sensor Read")

        # Start the background thread only if it's not already running
        if not (sensor_thread and sensor_thread.is_alive()):
            sensor_thread = threading.Thread(target=read_sensor_data, daemon=True)
            sensor_thread.start()
    else:
        start_btn.config(text="Click to Start Sensor Read")

# Read sensor data from serial port
def read_sensor_data():
    global read_enable
    if stm32_port is None:
        display_message("No Device Found")
        return
    else:
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
        except serial.SerialException as e:
            display_message(f"Serial error: {e}")

# Display message in the Text widget
def display_message(message):
    text_widget.insert(END, f"{message}\n")
    text_widget.yview(END)  # Auto-scroll to the end

# Clear the contents of the text widget
def clear_text_widget():
    with open("voltage_data.txt", "w") as file:
        pass
    text_widget.delete(1.0, END)

# Export voltage data to an Excel file
def export_saveas_excel():
    try:
        # Prompt user to choose a file to save
        file_path = filedialog.asksaveasfilename(
            title="Select or Create an Excel file",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not file_path:  # User canceled the dialog
            display_message("Export cancelled.")
            return

        # Create the Excel writer
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            # Write column header
            pd.DataFrame(columns=["Voltage (V)"]).to_excel(writer, index=False, header=True)

            # Open the text file and process each line
            with open("voltage_data.txt", "r") as file:

                for i, line in enumerate(file):
                    try:
                        voltage = float(line.strip())
                    except ValueError:
                        display_message(f"Skipping invalid line {i + 1}: {line.strip()}")

                    # Write the chunk to Excel when chunk_size is reached
                        df = pd.DataFrame(voltage, columns=["Voltage (V)"])
                        df.to_excel(writer, index=True, header=False)

        display_message(f"Data exported to {file_path}")
    except FileNotFoundError:
        display_message("Text file not found.")
    except ValueError:
        display_message("Error in data format.")
    except Exception as e:
        display_message(f"An error occurred: {e}")
        
def export_openas_excel():
    try:
        # Prompt user to choose a file to save
        file_path = filedialog.askopenfilename(
            title="Select or Create an Excel file",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not file_path:  # User canceled the dialog
            display_message("Export cancelled.")
            return

        # Create the Excel writer
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            # Write column header
            pd.DataFrame(columns=["Voltage (V)"]).to_excel(writer, index=True, header=True)

            # Open the text file and process each line
            with open("voltage_data.txt", "r") as file:

                for i, line in enumerate(file):
                    try:
                        voltage = float(line.strip())
                    except ValueError:
                        display_message(f"Skipping invalid line {i + 1}: {line.strip()}")

                    # Write the chunk to Excel when chunk_size is reached
                        df = pd.DataFrame(voltage, columns=["Voltage (V)"])
                        df.to_excel(writer, index=True, header=False)

        display_message(f"Data exported to {file_path}")

    except FileNotFoundError:
        display_message("Text file not found.")
    except ValueError:
        display_message("Error in data format.")
    except Exception as e:
        display_message(f"An error occurred: {e}")
        
def connect_mcu():
    global stm32_port
    if stm32_port is None:
        for port, desc, hwid in sorted(ports):
            if "STM" in desc:
                stm32_port = port
                display_message(f"STM32 port is {stm32_port}.\nMicrocontroller Connected")
                break
    else:
        display_message("Microcontroller Already Connected")

# Identify STM32 port
for port, desc, hwid in sorted(ports):
    if "STM" in desc:
        stm32_port = port
        display_message(f"STM32 port is {stm32_port}.\nMicrocontroller Connected")
        break

with open("voltage_data.txt", "w") as file:
        pass

# Create a button to start/stop reading sensor data
start_btn = Button(gui, text="Click to Start Sensor Read", width=40, command=toggle_enable)
start_btn.pack(pady=10)

# Create a button to clear the text widget
clear_btn = Button(gui, text="Clear Data", width=40, command=clear_text_widget)
clear_btn.pack(pady=10)

# Create a button to export data to Excel
export_saveas_btn = Button(gui, text="Save As .xlsx file", width=40, command=export_saveas_excel)
export_saveas_btn.pack(pady=10)

# Create a button to export data to Excel
export_openas_btn = Button(gui, text="Open and Export to existing .xlsx file", width=40, command=export_openas_excel)
export_openas_btn.pack(pady=10)

connect_btn = Button(gui, text="Connect Microcontroller", width=40, command=connect_mcu)
connect_btn.pack(pady=10)

# Exit button
exit_button = Button(gui, text="Exit", width=40, command=gui.destroy)
exit_button.pack(pady=10)


# Run the GUI main loop
gui.mainloop()
