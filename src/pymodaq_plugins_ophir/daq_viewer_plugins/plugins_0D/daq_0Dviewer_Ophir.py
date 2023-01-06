import numpy as np
from pymodaq.utils.daq_utils import ThreadCommand
from pymodaq.utils.data import DataFromPlugins
from pymodaq.control_modules.viewer_utility_classes import DAQ_Viewer_base, comon_parameters, main
from pymodaq.utils.parameter import Parameter
from pymodaq_plugins_ophir.hardware.OphirPowermeter import Ophir1EnergyMeter


class DAQ_0DViewer_Ophir(DAQ_Viewer_base):
    """
    """
    params = comon_parameters + [
        {'title': 'Ophir console info', 'name': 'c_info', 'type': 'str',
         'value': 'notyetknown'},
        {'title': 'Ophir sensor info', 'name': 's_info', 'type': 'str',
         'value': 'notyetknown'},
        {'title': 'Wavelength', 'name': 'w_length', 'type': 'list',
         'limits': ['pos1', 'pos2'], 'tip': 'Choose your laser WL. If not present use console.'},
        {'title': 'Range', 'name': 'm_range', 'type': 'list',
         'limits': ['pos1', 'pos2'], 'tip': 'Choose a range that triggers without saturating.'}
    ]

    def ini_attributes(self):
        self.controller: Ophir1EnergyMeter = None
        # Declare here attributes you want/need to init with a default value

    def commit_settings(self, param: Parameter):
        """Apply the consequences of a change of value in the detector settings

        Parameters
        ----------
        param: Parameter
            A given parameter (within detector_settings) whose value has been changed by the user
        """
        if param.name() == "w_length":
            self.controller.wavelength = param.value()
        elif param.name() == "m_range":
            self.controller.range = param.value()
        else:
            print("Parameter change not yet implemented")
        # Attention sometimes changing a parameter triggers a reading.
        # Maybe delete the generated reading if it was caused by this parameter change action.

    def ini_detector(self, controller=None):
        """Detector communication initialization

        Parameters
        ----------
        controller: (object)
            custom object of a PyMoDAQ plugin (Slave case).
            None if only one actuator/detector is used by this controller (Master case)

        Returns
        -------
        info: str
        initialized: bool
            False if initialization failed otherwise True
        """
        self.ini_detector_init(old_controller=controller,
                               new_controller=Ophir1EnergyMeter())

        initialized = self.controller.open_communication()  # tries to open comm ?? again ??

        if initialized:
            # Get the values to put in parameters (at startup)
            self.settings.child('c_info').setValue(self.controller.oph_device_name)  # string
            self.settings.child('s_info').setValue(self.controller.oph_sensor_name)  # string
            self.settings.child('w_length').setLimits(self.controller.wavelength_list)  # populate the list
            self.settings.child('w_length').setValue(self.controller.wavelength)  # display this item
            self.settings.child('m_range').setLimits(self.controller.range_list)  # populate the list
            self.settings.child('m_range').setValue(self.controller.range)  # display this item
            # self.settings.child('peak_amp').setValue(self.controller.amplitude)  # float

            self.data_grabed_signal_temp.emit(
                [DataFromPlugins(name='Pyro', data=[np.array([0])],
                                 dim='Data0D', labels=['pyro1'])]
            )

            info = "Ophir console and detector successfully initialized"
        else:
            info = "Ophir console could not be initialized"
        return info, initialized

    def close(self):
        """Terminate the communication protocol"""
        self.controller.close_communication()

    def grab_data(self, Naverage=1, **kwargs):
        """Start a grab from the detector

        Parameters
        ----------
        Naverage: int
            Number of hardware averaging (if hardware averaging is possible, self.hardware_averaging should be set to
            True in class preamble and you should code this implementation)
        kwargs: dict
            others optional arguments
        """
        # synchronous version (blocking function)
        data_tot = self.controller.get_data_1meas()
        print(f"ODv: {data_tot=}")
        self.data_grabed_signal.emit([DataFromPlugins(name='Pyro', data=[np.array([data_tot])],
                                                      dim='Data0D', labels=['pyro1'])])

        """# asynchrone version (non-blocking function with callback)
        raise NotImplemented  # when writing your own plugin remove this line
        self.controller.your_method_to_start_a_grab_snap(
            self.callback)  # when writing your own plugin replace this line
        #########################################################"""

    def callback(self):
        """optional asynchrone method called when the detector has finished its acquisition of data"""
        data_tot = self.controller.your_method_to_get_data_from_buffer()
        self.data_grabed_signal.emit([DataFromPlugins(name='Mock1', data=data_tot,
                                                      dim='Data0D', labels=['dat0', 'data1'])])

    def stop(self):
        """Stop the current grab hardware wise if necessary"""
        self.controller.stop_streams()  # when writing your own plugin replace this line
        self.emit_status(ThreadCommand('Update_Status', ['Some info you want to log']))
        ##############################
        return ''


if __name__ == '__main__':
    main(__file__)  # starts with automatic initialization
    # main(__file__, False)  # waits for manual initialization
