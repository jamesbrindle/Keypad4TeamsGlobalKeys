import board
import digitalio
import time
import usb_hid
from adafruit_hid import Keyboard
from adafruit_hid import Keycode

# Define inputs
tast01 = digitalio.DigitalInOut(board.GP10)
tast01.direction = digitalio.Direction.INPUT
tast01.pull = digitalio.Pull.DOWN
tast02 = digitalio.DigitalInOut(board.GP11)
tast02.direction = digitalio.Direction.INPUT
tast02.pull = digitalio.Pull.DOWN
tast04 = digitalio.DigitalInOut(board.GP16)
tast04.direction = digitalio.Direction.INPUT
tast04.pull = digitalio.Pull.DOWN
tast05 = digitalio.DigitalInOut(board.GP12)
tast05.direction = digitalio.Direction.INPUT
tast05.pull = digitalio.Pull.DOWN
tast06 = digitalio.DigitalInOut(board.GP13)
tast06.direction = digitalio.Direction.INPUT
tast06.pull = digitalio.Pull.DOWN
tast03 = digitalio.DigitalInOut(board.GP15)
tast03.direction = digitalio.Direction.INPUT
tast03.pull = digitalio.Pull.DOWN
tast08 = digitalio.DigitalInOut(board.GP14)
tast08.direction = digitalio.Direction.INPUT
tast08.pull = digitalio.Pull.DOWN
tast07 = digitalio.DigitalInOut(board.GP17)
tast07.direction = digitalio.Direction.INPUT
tast07.pull = digitalio.Pull.DOWN

# Set as USB Serial device
keyboard = Keyboard.Keyboard(usb_hid.devices)
key = Keycode.Keycode

# Reset buttons flag
restart = 1

# Whether or not to send pre-command keys (alt + 0)
sendPreCommand = 1

# Method to sent pre-command keys
def sendThePreCommand():
    if sendPreCommand == 1:
        keyboard.press(key.LEFT_ALT)
        time.sleep(0.02)
        keyboard.press(key.ZERO)
        time.sleep(0.02)
        keyboard.release_all()
        time.sleep(0.1)

# Read from text file
def inhaltLesen(variant, zahl):
    f = open("config/"+variant+"/taster0"+zahl+".txt", "r")
    wert1 = f.readline()
    wert01 = wert1.replace("\n", "")
    wert001 = inhaltWandeln(wert01)
    wert2 = f.readline()
    wert02 = wert2.replace("\n", "")
    wert002 = inhaltWandeln(wert02)
    wert3 = f.readline()
    wert03 = wert3.replace("\n", "")
    wert003 = inhaltWandeln(wert03)
    wert4 = f.readline()
    wert04 = wert4.replace("\n", "")
    wert004 = inhaltWandeln(wert04)
    f.close()
    print(wert1, wert2, wert3, wert4)
    print(wert01, wert02, wert03, wert04)
    print(wert001, wert002, wert003, wert004)
    return wert001, wert002, wert003, wert004

def inhaltWandeln(w):
    if w == "A":
        q = key.A
    elif w == "B":
        q = key.B
    elif w == "C":
        q = key.C
    elif w == "D":
        q = key.D
    elif w == "E":
        q = key.E
    elif w == "F":
        q = key.F
    elif w == "G":
        q = key.G
    elif w == "H":
        q = key.H
    elif w == "I":
        q = key.I
    elif w == "J":
        q = key.J
    elif w == "K":
        q = key.K
    elif w == "L":
        q = key.L
    elif w == "M":
        q = key.M
    elif w == "N":
        q = key.N
    elif w == "O":
        q = key.O
    elif w == "P":
        q = key.P
    elif w == "Q":
        q = key.Q
    elif w == "R":
        q = key.R
    elif w == "S":
        q = key.S
    elif w == "T":
        q = key.T
    elif w == "U":
        q = key.U
    elif w == "V":
        q = key.V
    elif w == "W":
        q = key.W
    elif w == "X":
        q = key.X
    elif w == "Y":
        q = key.Y
    elif w == "Z":
        q = key.Z

    elif w == "1":
        q = key.ONE
    elif w == "2":
        q = key.TWO
    elif w == "3":
        q = key.THREE
    elif w == "4":
        q = key.FOUR
    elif w == "5":
        q = key.FIVE
    elif w == "6":
        q = key.SIX
    elif w == "7":
        q = key.SEVEN
    elif w == "8":
        q = key.EIGHT
    elif w == "9":
        q = key.NINE
    elif w == "0":
        q = key.ZERO

    elif w == "ENTER":
        q = key.ENTER
    elif w == "RETURN":
        q = key.ENTER
    elif w == "ESCAPE":
        q = key.ESCAPE
    elif w == "BACKSPACE":
        q = key.BACKSPACE
    elif w == "TAB":
        q = key.TAB
    elif w == "SPACEBAR":
        q = key.SPACEBAR
    elif w == "SPACE":
        q = key.SPACEBAR
    elif w == "SHIFT":
        q = key.SHIFT
    elif w == "CONTROL":
        q = key.CONTROL
    elif w == "ALT":
        q = key.LEFT_ALT
    elif w == "OPTION":
        q = key.LEFT_ALT
    elif w == "CMD":
        q = key.LEFT_GUI
    elif w == "WINDOWS":
        q = key.LEFT_GUI

    else:
        q = 0x00
    return int(q)

# Get program type (Teams, Zoom etc) - In order to get button mapping
g_var = open("get_var.txt", "r")
set_var = g_var.readline()
g_var.close()

# Read and set buttons 1 to 8 from text files
t1w1, t1w2, t1w3, t1w4 = inhaltLesen(set_var, "1")
t2w1, t2w2, t2w3, t2w4 = inhaltLesen(set_var, "2")
t3w1, t3w2, t3w3, t3w4 = inhaltLesen(set_var, "3")
t4w1, t4w2, t4w3, t4w4 = inhaltLesen(set_var, "4")
t5w1, t5w2, t5w3, t5w4 = inhaltLesen(set_var, "5")
t6w1, t6w2, t6w3, t6w4 = inhaltLesen(set_var, "6")
t7w1, t7w2, t7w3, t7w4 = inhaltLesen(set_var, "7")
t8w1, t8w2, t8w3, t8w4 = inhaltLesen(set_var, "8")

# Main part of program - set output
while True:
    if tast01.value and restart == 1:
        sendThePreCommand()

        keyboard.press(t1w1)
        keyboard.press(t1w2)
        time.sleep(0.05)
        keyboard.press(t1w3)
        keyboard.press(t1w4)
        time.sleep(0.1)

        print("Tast 1 - End Call")
        restart = 0
    elif tast02.value and restart == 1:
        sendThePreCommand()

        keyboard.press(t2w1)
        keyboard.press(t2w2)
        time.sleep(0.05)
        keyboard.press(t2w3)
        keyboard.press(t2w4)
        time.sleep(0.05)

        print("Tast 2 - Reject Call")
        restart = 0
    elif tast03.value and restart == 1:
        sendThePreCommand()

        keyboard.press(t3w1)
        keyboard.press(t3w2)
        time.sleep(0.05)
        keyboard.press(t3w3)
        keyboard.press(t3w4)
        time.sleep(0.05)

        print("Tast 3 - Toggle Screen Share")
        restart = 0
    elif tast04.value and restart == 1:
        sendThePreCommand()

        keyboard.press(t4w1)
        keyboard.press(t4w2)
        time.sleep(0.05)
        keyboard.press(t4w3)
        keyboard.press(t4w4)
        time.sleep(0.05)

        print("Tast 4 - Raise and Lower Hand")
        restart = 0
    elif tast05.value and restart == 1:
        sendThePreCommand()

        keyboard.press(t5w1)
        keyboard.press(t5w2)
        time.sleep(0.05)
        keyboard.press(t5w3)
        keyboard.press(t5w4)
        time.sleep(0.05)

        print("Tast 5 - Answer Video Call")
        restart = 0
    elif tast06.value and restart == 1:
        sendThePreCommand()

        keyboard.press(t6w1)
        keyboard.press(t6w2)
        time.sleep(0.05)
        keyboard.press(t6w3)
        keyboard.press(t6w4)
        time.sleep(0.05)

        print("Tast 6 - Answer Audio Call")
        restart = 0
    elif tast07.value and restart == 1:
        sendThePreCommand()

        keyboard.press(t7w1)
        keyboard.press(t7w2)
        time.sleep(0.05)
        keyboard.press(t7w3)
        keyboard.press(t7w4)
        time.sleep(0.05)

        print("Tast 7 - Toggle Camera")
        restart = 0
    elif tast08.value and restart == 1:
        sendThePreCommand()

        keyboard.press(t8w1)
        keyboard.press(t8w2)
        time.sleep(0.05)
        keyboard.press(t8w3)
        keyboard.press(t8w4)
        time.sleep(0.05)

        print("Tast 8 - Toggle Microphone")
        restart = 0
    else:
        keyboard.release_all()
        restart = 1

    time.sleep(0.01)
