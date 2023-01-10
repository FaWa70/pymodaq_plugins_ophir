import serial as ser
import time as ti

"""For the test switch on the Ophir Nova II console (with a pyro head attached)
Then choose a sensitive range and tap on the detector to make an energy appear on 
the display. Then running the program below should give this output:
opened? True
b'* 0.326E-6\r\n'
still open? False

I did not succeed to use baud rates higher than 9600 baud.
Ranges are 0 indexed
Wavelengths are 1 indexed
"""

pyro_com = ser.Serial('COM7', baudrate=9600, timeout=1)
# std parameters: port=None, baudrate=9600, bytesize=EIGHTBITS, parity=PARITY_NONE,
# stopbits=STOPBITS_ONE, timeout=None, xonxoff=False, rtscts=False,
# write_timeout=None, dsrdtr=False, inter_byte_timeout=None, exclusive=None

print("opened?", pyro_com.is_open)
ti0 = ti.time()
pyro_com.write(b'$aw\r\n')     # write a string: fe, ar, aw, rn, wi, ef, se (28ms), sf
ti1 = ti.time()
print(pyro_com.readline())
ti2 = ti.time()

pyro_com.close()

print("still open?", pyro_com.is_open)

print(f"full time: {ti2 -ti0}")
print(f"send time: {ti1 -ti0}")
