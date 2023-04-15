## Code to Automatically open Zoom program to join a meeting and record meeting is required

#####################################

## Please read through the read me to understand how this works to implement it for your set up

#####################################

## You need the pyautogui library, best way to install is using 'pip install pyautogui'
## in essence this script uses your keyboard and mouse to automatically open zoom from start menu and click and enter meeting id
## some settings of zoom itself needs to be changed for this script to work correctly

## documentation for pyautogui

########################################################

## https://pyautogui.readthedocs.io/en/latest/index.html

########################################################

## Created by Rohit Kannachel
## Not liable if this causes any issues.
## thank you to Tanushi De Silva for giving me the inspiration to make this

# importing libraries
import pyautogui 
import time
import pyperclip
import subprocess
import datetime

#####   Joining Zoom Meeting   ###################

#######################################################################################
#Enter the meeting id as a string here *important that it is in string format
day = datetime.date.today().weekday()

if day == 0: # Num Mathe
    meet_id = '857 5093 6897'
    password = '912472'
    seconds_record = 5700
elif day == 1: # CN, ggf. Num Mathe
    """
    meet_id = '826 1336 7019'
    password = 'RNWS2021#'
    seconds_record = 9300
    """
    
    now = datetime.datetime.now()
    todayNoon = now.replace(hour=12, minute=0, second=0, microsecond=0)
    if now < todayNoon:
        meet_id = '826 1336 7019'
        password = 'RNWS2021#'
        seconds_record = 9300
    else:
        meet_id = '847 7389 1517'
        password = '443710'
        seconds_record = 5700
    

#esc clicked to ensure that the win key will open up correctly in the next step
pyautogui.press('esc',interval=0.1)

time.sleep(0.2)

#these lines are simulating starting up zoom by pressing windows key and typing zoom to open program
pyautogui.press('win',interval=0.1)
pyautogui.write('zoom')
pyautogui.press('enter',interval=0.5)

#time delay to factor for zoom app to load up, good buffer is like 10 sec but its case specific
time.sleep(5)

#this part simulates clicking join meeting, entering meeting id and pressing enter to join
##Make sure the joinButton.png file is located in the same folder as the python file or else it will not work
##this tells the script where to click to join the meeting

x,y = pyautogui.locateCenterOnScreen('joinButton.png')
pyautogui.click(x,y)

time.sleep(2)

## the interval of 1 second is important, if not there, then the meeting id will not be inputted
pyautogui.write(meet_id)
pyautogui.press('enter',interval=1)

time.sleep(3)
pyautogui.press('enter',interval=1)
#pyautogui.write(password)
for char in password:
    pyperclip.copy(char)
    pyautogui.hotkey('ctrl', 'v', interval=0.1)
pyautogui.press('enter',interval = 1)

##################################################################################


###### Recording meeting using Windows Game Bar  #############################


#######################################################################################


## this time buffer is added so that it accounts for the time taken to load into the meeting 
## a good buffer time is around 10-20 seconds before recording starts to ensure you're in the meeting

time.sleep(20)

## opening up windows game bar overlay
pyautogui.hotkey('win','g')
time.sleep(1)
## commencing screen recording
pyautogui.hotkey('win','alt','r')
time.sleep(1)
## closing windows game bar overlay
pyautogui.hotkey('win','g')


#### recording time amount
## however long you want, enter the time here in seconds, e.g. 30 minutes is 60*30 = 1800 seconds
## in windows game bar the default setting for time limit for recording is 2 hours,
## make sure to change this as you need

#time.sleep(seconds_record)
start = time.time()
while time.time() - start < seconds_record:
    joinBreakout = pyautogui.locateCenterOnScreen('joinBreakout.png')
    if joinBreakout is not None:
        pyautogui.click(joinBreakout[0],joinBreakout[1])
    time.sleep(15)


## ending screen recording
pyautogui.hotkey('win','alt','r')
time.sleep(2)
## By default, screen captures are sent to a folder called captures in "videos" in "this PC"

## closing Zoom
pyautogui.hotkey('alt','f4')
time.sleep(0.5)
pyautogui.hotkey('alt','f4')

subprocess.call(["shutdown", "/s"])

############################################################################################