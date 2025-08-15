import time
import os
import sys
from ctypes import *
import tkinter as tk
from tkinter import messagebox, filedialog
import pyvisa
import numpy as np

# Function to move the stage
def move_stage(serial_num, target_pos_real, n_steps, direction, save_path, wavelength):
    lib: CDLL = cdll.LoadLibrary("Thorlabs.MotionControl.TCube.DCServo.dll")
    
    STEPS_PER_REV = c_double(512)
    gbox_ratio = c_double(67.49)
    pitch = c_double(1.0)

    lib.CC_SetMotorParamsExt(serial_num, STEPS_PER_REV, gbox_ratio, pitch)
    lib.CC_RequestPosition(serial_num)
    time.sleep(0.2)
    current_dev_pos = c_int(lib.CC_GetPosition(serial_num))

    current_real_pos = c_double()
    lib.CC_GetRealValueFromDeviceUnit(serial_num, current_dev_pos, byref(current_real_pos), 0)
    print(f'Current position: {current_real_pos.value} mm')

    step_size_real = (target_pos_real.value - current_real_pos.value) / n_steps
    step_size_real = -abs(step_size_real) if direction == "backward" else abs(step_size_real)
    
    with open(save_path, "w") as power_file:
        power_file.write("Step\tPosition (mm)\tPower (W)\tError (W)")
        
        for step in range(1, n_steps + 1):
            next_target_real = c_double(current_real_pos.value - step * abs(step_size_real)) if direction == "backward" else c_double(current_real_pos.value + step * abs(step_size_real))
            next_target_dev = c_int()
            lib.CC_GetDeviceUnitFromRealValue(serial_num, next_target_real, byref(next_target_dev), 0)

            print(f'Step {step}: Moving to {next_target_real.value} mm')
            lib.CC_SetMoveAbsolutePosition(serial_num, next_target_dev)
            time.sleep(0.25)
            lib.CC_MoveAbsolute(serial_num)
            time.sleep(2)

            lib.CC_RequestPosition(serial_num)
            time.sleep(0.2)
            updated_dev_pos = c_int(lib.CC_GetPosition(serial_num))
            updated_real_pos = c_double()
            lib.CC_GetRealValueFromDeviceUnit(serial_num, updated_dev_pos, byref(updated_real_pos), 0)
            print(f'Position after step {step}: {updated_real_pos.value} mm')

            mean_power, std_power = measure_power(wavelength)
            power_file.write(f"\n{step}\t{updated_real_pos.value:.2f}\t{mean_power:.5f}\t{std_power:.5f}")
    
    lib.CC_Close(serial_num)

# Function to measure power with error calculation
def measure_power(wavelength):
    rm = pyvisa.ResourceManager()
    res_found = rm.list_resources('USB?*::0x1313::0x8078::?*::INSTR')
    if not res_found:
        raise Exception('Could not find the PM100D power meter.')
    
    meter = rm.open_resource(res_found[0])
    meter.read_termination = '\n'
    meter.write_termination = '\n'
    meter.timeout = 2000
    
    meter.write('sense:power:unit mW')  # Ensure the power meter is set to milliwatts
    meter.write('sense:power:range:auto 1')
    meter.write(f'sense:correction:wavelength {wavelength}')
    
    readings = [meter.query_ascii_values('read?')[0] for _ in range(5)]
    mean_power = np.mean(readings)
    std_power = np.std(readings)

    print(f'Measured Power: {mean_power:.5f} ± {std_power:.5f} mW')
    return mean_power, std_power

class StageControlApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Thorlabs Stage Control")

        tk.Label(master, text="Target Position (mm):").grid(row=0, column=0, padx=10, pady=10)
        self.target_pos_entry = tk.Entry(master)
        self.target_pos_entry.grid(row=0, column=1, padx=10, pady=10)

        tk.Label(master, text="Number of Steps:").grid(row=1, column=0, padx=10, pady=10)
        self.steps_entry = tk.Entry(master)
        self.steps_entry.grid(row=1, column=1, padx=10, pady=10)

        tk.Label(master, text="Direction:").grid(row=2, column=0, padx=10, pady=10)
        self.direction_var = tk.StringVar(value="forward")
        tk.Radiobutton(master, text="Forward", variable=self.direction_var, value="forward").grid(row=2, column=1, padx=10, pady=10, sticky="w")
        tk.Radiobutton(master, text="Backward", variable=self.direction_var, value="backward").grid(row=3, column=1, padx=10, pady=10, sticky="w")
        
        tk.Label(master, text="Wavelength (nm):").grid(row=4, column=0, padx=10, pady=10)
        self.wavelength_entry = tk.Entry(master)
        self.wavelength_entry.grid(row=4, column=1, padx=10, pady=10)
        
        tk.Button(master, text="Select Save Folder", command=self.select_save_path).grid(row=5, column=0, columnspan=2, pady=10)
        tk.Button(master, text="Move Stage", command=self.move_stage).grid(row=6, column=0, columnspan=2, pady=20)
        
        self.save_path = ""
    
    def select_save_path(self):
        folder = filedialog.askdirectory()
        if folder:
            file_name = filedialog.asksaveasfilename(initialdir=folder, defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
            if file_name:
                self.save_path = file_name
                print(f"Save path selected: {self.save_path}")
            else:
                messagebox.showerror("Error", "No file name specified.")
        else:
            messagebox.showerror("Error", "No folder selected.")
    
    def move_stage(self):
        try:
            target_pos = float(self.target_pos_entry.get())
            n_steps = int(self.steps_entry.get())
            direction = self.direction_var.get()
            wavelength = float(self.wavelength_entry.get())
            
            if n_steps <= 0:
                raise ValueError("Number of steps must be greater than zero.")
            if not self.save_path:
                raise ValueError("Please select a folder to save the power data.")
            
            os.add_dll_directory(r"C:\\Program Files\\Thorlabs\\Kinesis")
            lib: CDLL = cdll.LoadLibrary("Thorlabs.MotionControl.TCube.DCServo.dll")
            serial_num = c_char_p(b"83859973")
            
            if lib.TLI_BuildDeviceList() == 0:
                lib.CC_Open(serial_num)
                lib.CC_StartPolling(serial_num, c_int(200))
                move_stage(serial_num, c_double(target_pos), n_steps, direction, self.save_path, wavelength)
                lib.CC_Close(serial_num)
                messagebox.showinfo("Success", "Stage movement complete!")
            else:
                messagebox.showerror("Error", "Device initialization failed.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

def main():
    root = tk.Tk()
    app = StageControlApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()