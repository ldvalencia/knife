import time
import os
import sys
from ctypes import *
import tkinter as tk
from tkinter import messagebox

# Function to move the stage
def move_stage(serial_num, target_pos_real, n_steps, direction):
    """
    Moves the stage in `n` steps from the current position to a target position `x` in mm.
    :return: None
    """
    lib: CDLL = cdll.LoadLibrary("Thorlabs.MotionControl.TCube.DCServo.dll")

    # Set constants
    STEPS_PER_REV = c_double(512)  # for the MTS25-Z8
    gbox_ratio = c_double(67.49)  # gearbox ratio
    pitch = c_double(1.0)

    # Apply these values to the device
    lib.CC_SetMotorParamsExt(serial_num, STEPS_PER_REV, gbox_ratio, pitch)

    # Get the device's current position in dev units
    lib.CC_RequestPosition(serial_num)
    time.sleep(0.2)
    current_dev_pos = c_int(lib.CC_GetPosition(serial_num))

    # Convert current device position to real units (mm)
    current_real_pos = c_double()
    lib.CC_GetRealValueFromDeviceUnit(serial_num, current_dev_pos, byref(current_real_pos), 0)
    print(f'Current position: {current_real_pos.value} mm')

    # Calculate the step size in real units (mm)
    step_size_real = (target_pos_real.value - current_real_pos.value) / n_steps
    if direction == "backward":
        step_size_real = -abs(step_size_real)  # Ensure step size is negative for backward movement
    else:
        step_size_real = abs(step_size_real)   # Ensure step size is positive for forward movement

    # Move the stage in `n` steps
    for step in range(1, n_steps + 1):
        # Calculate the next target position in real units
        if direction == "backward":
            next_target_real = c_double(current_real_pos.value - step * abs(step_size_real))  # Moving backward
        else:
            next_target_real = c_double(current_real_pos.value + step * abs(step_size_real))  # Moving forward

        # Convert the next target position to device units
        next_target_dev = c_int()
        lib.CC_GetDeviceUnitFromRealValue(serial_num, next_target_real, byref(next_target_dev), 0)

        print(f'Step {step}: Moving to {next_target_real.value} mm in Device Units: {next_target_dev.value}')

        # Move to the next target position as an absolute move
        lib.CC_SetMoveAbsolutePosition(serial_num, next_target_dev)
        time.sleep(0.25)
        lib.CC_MoveAbsolute(serial_num)

        # Wait for the motor to move
        time.sleep(2)

        # Get the updated position to confirm the move
        lib.CC_RequestPosition(serial_num)
        time.sleep(0.2)
        updated_dev_pos = c_int(lib.CC_GetPosition(serial_num))

        # Convert updated device position to real units
        updated_real_pos = c_double()
        lib.CC_GetRealValueFromDeviceUnit(serial_num, updated_dev_pos, byref(updated_real_pos), 0)

        print(f'Position after step {step}: {updated_real_pos.value} mm')

    # Final position check
    if abs(updated_real_pos.value - target_pos_real.value) < 0.1:  # Adjust tolerance as needed
        print("The motor has reached the target position.")
    else:
        print("The motor has not reached the target position.")

    # Close the device
    lib.CC_Close(serial_num)


# GUI Setup
class StageControlApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Thorlabs Stage Control")

        # GUI Labels and Inputs
        self.target_pos_label = tk.Label(master, text="Target Position (mm):")
        self.target_pos_label.grid(row=0, column=0, padx=10, pady=10)

        self.target_pos_entry = tk.Entry(master)
        self.target_pos_entry.grid(row=0, column=1, padx=10, pady=10)

        self.steps_label = tk.Label(master, text="Number of Steps:")
        self.steps_label.grid(row=1, column=0, padx=10, pady=10)

        self.steps_entry = tk.Entry(master)
        self.steps_entry.grid(row=1, column=1, padx=10, pady=10)

        self.direction_label = tk.Label(master, text="Direction:")
        self.direction_label.grid(row=2, column=0, padx=10, pady=10)

        self.direction_var = tk.StringVar(value="forward")
        self.direction_forward = tk.Radiobutton(master, text="Forward", variable=self.direction_var, value="forward")
        self.direction_forward.grid(row=2, column=1, padx=10, pady=10, sticky="w")

        self.direction_backward = tk.Radiobutton(master, text="Backward", variable=self.direction_var, value="backward")
        self.direction_backward.grid(row=3, column=1, padx=10, pady=10, sticky="w")

        self.move_button = tk.Button(master, text="Move Stage", command=self.move_stage)
        self.move_button.grid(row=4, column=0, columnspan=2, pady=20)

    def move_stage(self):
        try:
            target_pos = float(self.target_pos_entry.get())
            n_steps = int(self.steps_entry.get())
            direction = self.direction_var.get()  # Get selected direction

            if n_steps <= 0:
                raise ValueError("Number of steps must be greater than zero.")

            if sys.version_info < (3, 8):
                os.chdir(r"C:\Program Files\Thorlabs\Kinesis")
            else:
                os.add_dll_directory(r"C:\Program Files\Thorlabs\Kinesis")

            # Load the Thorlabs library
            lib: CDLL = cdll.LoadLibrary("Thorlabs.MotionControl.TCube.DCServo.dll")

            serial_num = c_char_p(b"83859973")  # Update with your serial number

            # Initialize the device
            if lib.TLI_BuildDeviceList() == 0:
                lib.CC_Open(serial_num)
                lib.CC_StartPolling(serial_num, c_int(200))

                # Set up device and move
                target_pos_real = c_double(target_pos)
                move_stage(serial_num, target_pos_real, n_steps, direction)

                # Close the device
                lib.CC_Close(serial_num)

                messagebox.showinfo("Success", "Stage has successfully moved!")
            else:
                messagebox.showerror("Error", "Device initialization failed.")
        except ValueError as e:
            messagebox.showerror("Input Error", f"Invalid input: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")


# Main function to run the GUI
def main():
    root = tk.Tk()
    app = StageControlApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
