pymodaq_plugins_ophir (Nova II console)
#######################################

.. the following must be adapted to your developped package, links to pypi, github  description...

.. image:: https://img.shields.io/pypi/v/pymodaq_plugins_thorlabs.svg
   :target: https://pypi.org/project/pymodaq_plugins_thorlabs/
   :alt: Latest Version

.. image:: https://readthedocs.org/projects/pymodaq/badge/?version=latest
   :target: https://pymodaq.readthedocs.io/en/stable/?badge=latest
   :alt: Documentation Status

.. image:: https://github.com/PyMoDAQ/pymodaq_plugins_thorlabs/workflows/Upload%20Python%20Package/badge.svg
   :target: https://github.com/PyMoDAQ/pymodaq_plugins_thorlabs
   :alt: Publication Status

Set of PyMoDAQ plugins for instruments from Ophir (www.ophiropt.com).

Restrictions
============
I only tested this on Windows 10. Probabaly it works on other versions of Windows too.
But, as the communication is made using a COM object from teh Win32 API, I doubt that
this works on Linux or Mac.

Installing the Driver and necessary Python packages
===================================================
I installed The Starlab 3.80 software from Ophir. This installed and registered the
`OphirLMMeasurement.dll`. A more detailed description of the process and how it can
be transferred to another computer without installing Starlab is given in chapter 3
of the `OphirLMMeasurement COM Object.doc` file that is provided iin the `\hardware`
subdirectory.

Python can talk with the WIN32 API COM object only if the `pywin32` package is installed.
I installed it using conda (conda install -c anaconda pywin32)
but I think it will work similarly with pip.


Authors
=======

* Frank Wagner  (frank.wagner@fresnel.fr)
* Other author (myotheremail@xxx.org)

.. if needed use this field

    Contributors
    ============

    * First Contributor
    * Other Contributors

Instruments
===========

Below is the list of instruments included in this plugin

Viewer0D
++++++++

* **Ophir1EnergyMeter**: Control of Ophir Laser Energy Meter 0D detector
(Tested with a Nova II console and pyroelectric detectors)

There is one Git branch for each stream mode:

- Standard mode:
The Windows COM object saves some values
in a buffer and transfers them after firing the DataReady event.

*State:* I have no idea how to use this event. For the moment I just call the GetData method
and throw away all but the first data entry. Many calls of GetData result in None being returned
which seems to be ok for the display, but I don't know what the h5 file looks like.
If a wait time is used in the general parameters,
it works more or less. (No idea how many pulses are missing.)

- Immediate mode:
The Windows COM object tries to transfer the data after each trigger. According to doc
this should work at up to 100 Hz. Sometimes the data contains also pulse repetition rate information.
These are at the positions where the status tuple gives 0x050000 (or 327680 in decimal) and it's
never the first position. For the moment this information is discarded

*State:* This works now. Best first push the 'single grab' button anyway.

* **xxx**: control of xxx 0D detector

Viewer1D
++++++++

* **yyy**: control of yyy 1D detector
* **xxx**: control of xxx 1D detector


Viewer2D
++++++++

* **yyy**: control of yyy 2D detector
* **xxx**: control of xxx 2D detector


Infos
=====

if needed for installation or other infos
