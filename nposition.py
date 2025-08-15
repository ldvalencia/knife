import time
import os
import sys
from ctypes import *


def main():
    """
    main():
    ------

    Moves the stage in `n` steps from the current position to a target position `x` in mm.
    :return: None
    """

    if sys.version_info < (3, 8):
        os.chdir(r"C:\Program Files\Thorlabs\Kinesis")
    else:
        os.add_dll_directory(r"C:\Program Files\Thorlabs\Kinesis")

    lib: CDLL = cdll.LoadLibrary("Thorlabs.MotionControl.TCube.DCServo.dll")

    # Uncomment this line if you are using simulations
    # lib.TLI_InitializeSimulations()

    # Set constants
    serial_num = c_char_p(b"83859973")  # Update to your device's serial number

    # Open the device
    if lib.TLI_BuildDeviceList() == 0:
        lib.CC_Open(serial_num)
        lib.CC_StartPolling(serial_num, c_int(200))

        # Set up the device to convert real units to device units
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

        # Define the target position in real units (mm)
        target_pos_real = c_double(2)  # Target position in mm (change this value as needed)
        n_steps = 2  # Number of steps (change this value as needed)

        # Calculate the step size in real units (mm)
        step_size_real = (target_pos_real.value - current_real_pos.value) / n_steps

        # Move the stage in `n` steps
        for step in range(1, n_steps + 1):
            # Calculate the next target position in real units
            next_target_real = c_double(current_real_pos.value + step * step_size_real)

            # Convert the next target position to device units
            next_target_dev = c_int()
            lib.CC_GetDeviceUnitFromRealValue(serial_num, next_target_real, byref(next_target_dev), 0)

            print(f'Step {step}: Moving to {next_target_real.value} mm in Device Units: {next_target_dev.value}')

            # Move to the next target position as an absolute move
            lib.CC_SetMoveAbsolutePosition(serial_num, next_target_dev)
            time.sleep(0.25)
            lib.CC_MoveAbsolute(serial_num)

            # Wait for the motor to move
            time.sleep(2)  # Adjust this delay based on the step size

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
        # lib.TLI_UninitializeSimulations()

    return


if __name__ == "__main__":
    main()