# Sh-StartUp.py
## Description
This Python program acts as an alternative (to autoexec) method to start number of programs in Windows.
It selects programs batch from JSON file automatically (based on WiFi network SSID available) or 
manually and starts apps one by one.

May be useful when you work in different locations (home/office) or different environments.
Although the main mode is autostart at PC login, it may also be useful to start the new environment.

JSON configuration file may contain unlimited number of configurations and applications.

## Usage
pythonw.exe Sh-Start-Up.py \[*configname*]

## Dependencies
This program uses ***sh_messagebox*** module.
