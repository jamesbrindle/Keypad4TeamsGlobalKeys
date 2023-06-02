# Keypad4Teams

[![N|Solid](https://portfolio.jb-net.co.uk/shared/keypad4teams.png)](https://github.com/jamesbrindle/Keypad4TeamsGlobalKeys/releases)

This is a tiny program that allows the 'Keypad4Teams' 3D printed custom macro button set's keys to be 'global'.

https://www.thingiverse.com/thing:4883188

Normally, the keypad only works when the call Teams Window has focus. This programs helps the keypad work 'globally' - So even when Teams isn't in focus.


This program sits in the system tray that looks out for a particular key combination (default: Alt + 0). The keypad is set to send this key combination first, then 
send the usual Teams shortcut key combinations.

When Alt + 0 is handled, it will look for the relevant Teams window and bring it to the front.


## Setup

[Download Installer](https://github.com/jamesbrindle/Keypad4TeamsGlobalKeys/releases)

The installer will install to the Program Files folder and set the program to start on Windows statup. It can be exited from the right-click context menu on the icon in the system tray.

There are 3 types of installer:

1. Msi - 64 bit
2. Msi - 32 bit
3. ClickOnce (you don't need elevated permissions when installing using this method)

## Raspberry Pi Pico

After wiring the Pico (Refer to wiring diagram here: https://www.thingiverse.com/thing:4883188). 

1. Hold the 'boot' button on the Pico while plugging in the USB connected to the computer to load the bootloader.
2. Copy the '\PicoFiles\CircuitPython\adafruit-circuitpython-raspberry_pi_pico-de_DE-6.2.0.uf2' to the USB storage
3. This will reboot the Pico in standard serial mode. Then copy all files from \PicoFiles\Code to the Pico - Overwriting any existing

The keypad will now be ready to use.
