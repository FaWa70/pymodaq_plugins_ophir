# Use of Ophir COM object. 
# Works with python 3.5.1 & 2.7.11. I tested with python 3.8, and it works
# Uses pywin32  # Python extensions for MS Win. Provides access to much of the Win32 API,
#               # the ability to create and use COM objects ...
import win32com.client  # conda install -c anaconda pywin32
import win32gui   # seems to be contained in pywin32.
# pip install win32gui ends with error message (but may be useful anyway?)
import time
import traceback

try:
    OphirCOM = win32com.client.Dispatch("OphirLMMeasurement.CoLMMeasurement")
    # Stop & Close all devices
    OphirCOM.StopAllStreams()
    OphirCOM.CloseAll()
    # Scan for connected Devices
    DeviceList = OphirCOM.ScanUSB()
    print(f"{DeviceList=}")
    for Device in DeviceList:  # if any device is connected
        DeviceHandle = OphirCOM.OpenUSBDevice(Device)  # open first device
        exists = OphirCOM.IsSensorExists(DeviceHandle, 0)
        if exists:
            print('\n----------Data for S/N {0} ---------------'.format(Device))

            # An Example for Range control. first get the ranges
            ranges = OphirCOM.GetRanges(DeviceHandle, 0)
            print(f"{ranges=}")
            # For PE25-C: ranges = (0, ('10.0J', '2.00J', '200mJ', '20.0mJ', '2.00mJ', '200uJ'))
            # The first entry, ranges[0], is the index of the active range in ranges[1]
            # new_range replaces ranges[0]
            new_range = 5  # For this head this corresponds to 200uJ
            print(f"{ranges[1][new_range]=}")
            # set new range
            OphirCOM.SetRange(DeviceHandle, 0, new_range)

            # An Example for data retrieving
            OphirCOM.StartStream(DeviceHandle, 0)  # start measuring
            for i in range(10):
                time.sleep(1)  # wait a little for data
                data = OphirCOM.GetData(DeviceHandle, 0)
                if len(data[0]) > 0:  # if any data available, print the first one from the batch
                    print(
                        'Reading = {0}, TimeStamp = {1}, Status = {2} '.format(
                            data[0][0], data[1][0], data[2][0]))

        else:
            print('\nNo Sensor attached to {0} !!!'.format(Device))
except OSError as err:
    print("OS error: {0}".format(err))
except:
    traceback.print_exc()

win32gui.MessageBox(0, 'finished', '', 0)
# Stop & Close all devices
OphirCOM.StopAllStreams()
OphirCOM.CloseAll()
# Release the object
OphirCOM = None
