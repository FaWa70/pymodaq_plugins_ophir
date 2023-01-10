rem  1) remove ophir devices drivers (Ophir specific INF files)
swapinf StarlabDriversRemove
rem  2) remove COM (this is Ophir specific)
regsvr32 /u OphirLMMeasurement.dll
