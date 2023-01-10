# Use of Ophir COM object. 
# Tested with python 3.8, and it works
# Uses pywin32  # Python extensions for MS Win. Provides access to much of the Win32 API,
#               # the ability to create and use COM objects ...
import win32com.client  # conda install -c anaconda pywin32
import pythoncom  # for having the COM object available in all threads
import traceback
import time

class Ophir1EnergyMeter:  # Works if there is only ONE powermeter connected
    info = 'Wrapper for the interface with Ophir power meter console Nova2 0.0.1'
    # Any class variables ?

    def __init__(self):
        super().__init__()
        self._ophir_com = None  # will contain the windows API COM object provided by Ophir
        self._oph_device_list = None
        self._oph_device_handle = None
        self._oph_device_name = None
        self._oph_exists = None
        self._oph_sensor_name = None
        # The _list parameters hold the possible values.
        # The parameters without _list hold the presently selected values.
        # The parameters with _idx hold the presently selected index in the list (or iterable).
        self._range_list = None
        self._range_idx = None
        self._wavelength_list = None
        self._wavelength_idx = None
        self._reply_old = ""
        self._ti_out = False

    def open_communication(self):
        """
        Searches for Ophir consoles on USB. Goes further if exactly one was found.
        Then checks if a sensor is attached, and, if there is one, sets energy measurement mode.
        """
        try:
            pythoncom.CoInitialize()  # This line is necessary if different threads use the COM object.
            self._ophir_com = win32com.client.Dispatch("OphirLMMeasurement.CoLMMeasurement")
            # Stop & Close all devices
            self._ophir_com.StopAllStreams()
            self._ophir_com.CloseAll()
            # Scan for connected devices
            self._oph_device_list = self._ophir_com.ScanUSB()  # contains the serial numbers
            if len(self._oph_device_list) == 1:
                self._oph_device_handle = self._ophir_com.OpenUSBDevice(self._oph_device_list[0])  # open first device
                # Get the name of the console device
                _dinfo = self._ophir_com.GetDeviceInfo(self._oph_device_handle)
                print(f"{_dinfo=}")
                self._oph_device_name = _dinfo[0]
                self._oph_exists = self._ophir_com.IsSensorExists(self._oph_device_handle, 0)
                # All below could be done with legacy mode (write / read  mode)
                if self._oph_exists:
                    # Get the name of the sensor (attached on channel 0)
                    _sinfo = self._ophir_com.GetSensorInfo(self._oph_device_handle, 0)
                    print(f"{_sinfo=}")
                    self._oph_sensor_name = _sinfo[-1]
                    # Make sure that energy measurement mode is used
                    _mes_mode = self._ophir_com.GetMeasurementMode(self._oph_device_handle, 0)
                    print(f"{_mes_mode=}")
                    if _mes_mode[0] != 1:  # Set to energy measurement (index 1 in _mes_mode[1])
                        self._ophir_com.SetMeasurementMode(self._oph_device_handle, 0, 1)
                        # SetMeasurementMode(handle, channel, wanted_index)
                        _mes_mode = self._ophir_com.GetMeasurementMode(self._oph_device_handle, 0)
                    print(f"Uses {_mes_mode[1][_mes_mode[0]]} measurement mode")
                    return True
                else:
                    print(f'\nNo Sensor attached to {self._oph_device_list[0]} !!!')
                    return False
            else:
                print(f'\nThere are {len(self._oph_device_list)} consoles instead of 1 !!!')
                return False
        except OSError as err:
            print("OS error: {0}".format(err))
            return False
        except:
            traceback.print_exc()
            return False

    def close_communication(self):
        # Stop & Close all devices
        self._ophir_com.StopAllStreams()
        self._ophir_com.CloseAll()
        # Release the object
        self._ophir_com = None
        return True

    def str_2_num(self, s):
        if s[0] == "*":
            status_ok = True
            val = float(s[1:-1])
        else:
            status_ok = False
            val = -10
        return status_ok, val

    def get_data_1meas(self):
        print('----\nget_data_1meas called')

        self._ophir_com.Write(self._oph_device_handle, "SE")  # send energy
        # This sends the last measured energy. Duplicate readings arrive if the
        # reading rate is faster than the laser pulse repetition rate.
        # Some intermediate pulses may be lost if the PRR is too fast (no idea about the limit)
        reply = self._ophir_com.Read(self._oph_device_handle)
        status_ok, data = self.str_2_num(reply)
        # Avoid duplicate readings
        ti_out = 10  # seconds
        self._ti_out = False
        start_time = time.time()
        while self._reply_old == reply and not self._ti_out:
            time.sleep(50e-3)  # wait a little for data
            self._ti_out = time.time() - start_time > ti_out
            self._ophir_com.Write(self._oph_device_handle, "SE")
            reply = self._ophir_com.Read(self._oph_device_handle)
            # One could also ask repeatedly for the energy-flag (EF)
            # and read the value only if the reply is 1. This would
            # not exclude two successive pulses of same energy

        if reply[0] != "*":  # avoids calling str_2_num in the loop
            self._ti_out = True
            # *OVER is the reply for overrange values

        if self._ti_out:
            print('time out or error occured in get_data_1meas')
            return None
        print(reply[:-1])
        self._reply_old = reply
        status_ok, data = self.str_2_num(reply)
        print(f"hw: {data=}")  # Is just the energy in Joules
        return data

    def stop_streams(self):
        self._ophir_com.StopAllStreams()

    @property
    def oph_device_name(self):
        return self._oph_device_name

    @property
    def oph_sensor_name(self):
        return self._oph_sensor_name

    @property
    def range_list(self):
        """
        Fetches the tuple of possible measurement ranges
        Returns
        -------
        tuple of strings: possible measurement ranges
        """
        _ranges = self._ophir_com.GetRanges(self._oph_device_handle, 0)
        self._range_list = _ranges[1]
        return self._range_list

    @property
    def range(self):
        """
        fetches the presently chosen measurement range
        Returns
        -------
        string: presently chosen measurement range, an element of range_list
        """
        _ranges = self._ophir_com.GetRanges(self._oph_device_handle, 0)
        self._range_idx = _ranges[0]
        self._range_list = _ranges[1]
        return self._range_list[self._range_idx]

    @range.setter
    def range(self, value):
        """
        Set the measurement range to one of the possible choices
        Parameters
        ----------
        value: (string) an element of range_list
        """
        if value not in self._range_list:
            raise ValueError(f'The range {value} cannot be used. (Wrong sensor or wrong mode.)')
        else:
            self._ophir_com.SetRange(self._oph_device_handle, 0, self._range_list.index(value))
            # SetRange(handle, channel, wanted_index)
            print(f"new range is {self.range}")  # this calls the property getter self.range

    @property
    def wavelength_list(self):
        """
        Fetches the tuple of possible wavelengths. If your wavelengths is not
        available: Stop the communication with the console to unlock the console keypad.
        Then modify/create and save the wanted wavelength using the console buttons.
        Once it's available in the console it will appear here.
        Returns
        -------
        tuple of strings: possible wavelengths
        """
        _wavelengths = self._ophir_com.GetWavelengths(self._oph_device_handle, 0)
        self._wavelength_list = _wavelengths[1]
        return self._wavelength_list

    @property
    def wavelength(self):
        """
        Fetches the presently chosen wavelength
        Returns
        -------
        string: presently chosen wavelength, an element of wavelength_list
        """
        _wavelengths = self._ophir_com.GetWavelengths(self._oph_device_handle, 0)
        self._wavelength_idx = _wavelengths[0]
        self._wavelength_list = _wavelengths[1]
        return self._wavelength_list[self._wavelength_idx]

    @wavelength.setter
    def wavelength(self, value):
        """
        Set the characteristic time to reach a particular wavelength
        Parameters
        ----------
        value: (string) an element of wavelength_list
        """
        if value not in self._wavelength_list:
            raise ValueError(f'The wavelength {value} cannot be used. (First add it using the console)')
        else:
            self._ophir_com.SetWavelength(self._oph_device_handle, 0, self._wavelength_list.index(value))
            # SetWavelength(handle, channel, wanted_index)
            print(f"new wavelength is {self.wavelength}")  # this calls the property getter self.wavelength

if __name__ == '__main__':
    ophir_1_powermeter = Ophir1EnergyMeter()
    ophir_1_powermeter.open_communication()
    for i in range(10):
        time.sleep(0.8)  # wait a little for data
        dattta = ophir_1_powermeter.get_data_1meas()
        print(dattta)
    ophir_1_powermeter.close_communication()

