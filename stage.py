import time
import os
import sys
from ctypes import *


def main():
    """
    main():
    ------

    Moves the stage to a specified position in mm without homing.
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
        dev_pos = c_int(lib.CC_GetPosition(serial_num))

        # Convert device units to real units
        real_pos = c_double()
        lib.CC_GetRealValueFromDeviceUnit(serial_num, dev_pos, byref(real_pos), 0)

        print(f'Current position: {real_pos.value} mm')

        # Define the target position in real units (mm)
        target_pos_real = c_double(0)  # Target position in mm (change this value as needed)
        target_pos_dev = c_int()

        # Convert the target position from real units to device units
        lib.CC_GetDeviceUnitFromRealValue(serial_num, target_pos_real, byref(target_pos_dev), 0)

        print(f'Moving to {target_pos_real.value} mm in Device Units: {target_pos_dev.value}')

        # Move to the new position as an absolute move
        lib.CC_SetMoveAbsolutePosition(serial_num, target_pos_dev)
        time.sleep(0.25)
        lib.CC_MoveAbsolute(serial_num)

        # Wait for the motor to move
        print("Waiting for the motor to reach the target position...")
        time.sleep(2)  # Adjust this delay based on the distance to move

        # Get the updated position to confirm the move
        lib.CC_RequestPosition(serial_num)
        time.sleep(0.2)
        updated_dev_pos = c_int(lib.CC_GetPosition(serial_num))

        # Convert updated device position to real units
        updated_real_pos = c_double()
        lib.CC_GetRealValueFromDeviceUnit(serial_num, updated_dev_pos, byref(updated_real_pos), 0)

        print(f'Position after moving: {updated_real_pos.value} mm')

        # Check if the motor has moved to the target position
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