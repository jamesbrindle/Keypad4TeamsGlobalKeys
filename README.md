# Keypad4Teams

This is a tiny program that allows the 'Keypad4Teams' 3D printed custom macro button set's keys to be 'global'.

https://www.thingiverse.com/thing:4883188

Normally, the keypad only works when the call Teams Window has focus. This programs helps the keypad work 'globally' - So even when Teams isn't in focus.


This program sits in the system tray that looks out for a particular key combination (default: Alt + 0). The keypad is set to send this key combination first, then 
send the typically Teams shortcut key combinations.

When Alt + 0 is handled, it will look for the relevant Teams window and bring it to the front.


## Setup

The installer will install to the Program Files folder and set the program to start on Windows statup. It can be exitted from the right-click context menu on the icon in the system tray.