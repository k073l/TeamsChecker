# TeamsChecker

A script built to make sure you won't miss the moment when the meeting had started

 ## Usage

- Install requirements by `pip install -r requirements.txt`

- Run `main.py` 

- Select main Teams window (in case there are multiple) by typing the number next to window title
- Script will check if meeting has started every 200 ms (on average)
- When meeting starts, you will hear the sound alarm which will not stop until more than one window of Teams is present (meaning, until you join the meeting)

### Known issues

It runs on Windows only.

Window flickering is normal and expected, due to screenshotting window using pywin32 (You might observe appearing and disappearing window title bar and borders)

If you have left the meeting which hasn't ended yet, expect the alarm to go off. To silence it, simply navigate to different team tab.

### But it doesn't work!

It is possible the provided join button image is different than the one provided. Simply run the script `main.py -d`. It will launch the script in debug mode, where you got access to debug information. Additionally, screenshots made by script (ex. `test_00000.png` in assets directory) aren't deleted each time, so you can take one with join button present and crop it to contain join button only (or, as provided image, only the color of that button). Remember to save the file as `join_whatever.png` in `assets` directory. Also note, that script uses first image that it sees that starts with `join`, so in case you have multiple files for testing which one is suitable for you, change their filename to ex. `_join` to make sure the script uses the one you want it to.
