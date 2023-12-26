# keyhelpers - Python app running as background utility
Keystroke commands for correcting text in any editable field with Python, in Windows only.  
Tested on Windows 8 and a few times in 10.

The idea is to work with text using keyboard simulation, something like like Punto Switcher app does.  
This approach does not require to know which textfield is current is and where is is located on the screen.

Features implemented so far:
- Turn the word under cursor to lowercase & toggle case of its first letter on subsequent hits.
- Turn all the selescted text to lowercase.
- Remove all extra spaces leaving single spaces between words only.

