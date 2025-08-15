# Import & initialize the PyVISA library
import pyvisa
import numpy as np

rm = pyvisa.ResourceManager()

# Find the power meter: we know it's a USB device from vendor 0x1313 (Thorlabs),
# and with model 0x8078 (PM100D).
res_found = rm.list_resources('USB?*::0x1313::0x8078::?*::INSTR')
if not res_found:
    raise Exception('Could not find the PM100D power meter. Is it connected and powered on?')

# Connect to the power meter
print('Connecting to PM100D...')
meter = rm.open_resource(res_found[0])
meter.read_termination = '\n'
meter.write_termination = '\n'
meter.timeout = 2000  # ms

meter.write('system:beeper')

print('*idn?')
print('--> ' + meter.query('*idn?'))

# Configure the power meter for laser power measurements
wavelength = 1064  # nm

meter.write('sense:power:unit mW')  # Change unit to mW (milliwatts)
meter.write('sense:power:range:auto 1')  # Auto range
meter.write('sense:average:count 50')  # Number of averages for reading
meter.write('configure:power')  # Configure for power measurement

meter.write('sense:correction:wavelength %.1f' % wavelength)

# Take 4 measurements and store them
measurements = []
for i in range(4):
    cur_power = meter.query_ascii_values('read?')[0]
    measurements.append(cur_power)
    print(f'Measurement {i+1}: {cur_power:.2f} mW')

# Calculate the average power and the error (standard deviation)
average_power = np.mean(measurements)
error = np.std(measurements)

print(f'\nAverage power: {average_power:.5f} mW')
print(f'Error (standard deviation): {error:.5f} mW')
