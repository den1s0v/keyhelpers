# keyhelpers - Python app running as background utility
Keystroke commands for correcting text in any editable field with Python, in Windows only.  
Tested on Windows 8 and 10 (using Python 3.9 and 3.12).




The idea is to work with text using keyboard simulation and clipboard read/write operations, something like [Punto Switcher](yandex.ru/soft/punto/win) app does.  
With this approach, it is not necessary to know which text field has the input focus and where it is located on the screen (however, one may need to check whether any text field is focused before simulating a lot of keystrokes).




Features implemented so far:
- Turn the word under cursor to lowercase & toggle the case of its first letter on subsequent hits.
- Turn all the selected text to lowercase.
- Remove all extra spaces leaving single spaces between words.







