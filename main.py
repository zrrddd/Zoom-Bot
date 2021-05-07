import subprocess
import pyautogui
import time
import pandas as pd
from datetime import datetime


def sign_in(m_id):
	# Open Zoom App
	subprocess.call(["/usr/bin/open", "/Applications/zoom.us.app"])

	# Join Meeting
	time.sleep(5)
	pyautogui.hotkey('command', 'j', interval=0.25)

	# Enter meeting ID
	pyautogui.write(m_id)
	pyautogui.press('enter')

	# Mute 
	time.sleep(5)
	pyautogui.hotkey('command', 'shiftleft', 'a', interval=0.25)


df = pd.read_excel("Zoom_meetings.xlsx")
have_start_meeting = False
have_end_meeting = False

while not have_start_meeting:
	# checking of the current time 
	now = datetime.now().strftime("%H:%M")
	for i in range(df.shape[0]):
		start_time = str(df['Start'][i])
		end_time = str(df['End'][i])
		if now == start_time[:-3]: # find the class to join 
			m_id = str(df['ID'][i])

			sign_in(m_id)
			time.sleep(30)
			print("Joined the meeting")
			have_start_meeting = True


while have_start_meeting and not have_end_meeting:
	now = datetime.now().strftime("%H:%M")
	if now == end_time[:-3]:
		time.sleep(2)
		# Quit the meeting
		pyautogui.hotkey('command', 'w', interval=0.25)
		time.sleep(1)
		pyautogui.press('enter')
		print("Quit the meeting")
		have_end_meeting = True

# Quit Zoom
time.sleep(2)
pyautogui.hotkey('command', 'q', interval=0.25)

# Lock My Screen 
time.sleep(2)
pyautogui.hotkey('command', 'ctrl', 'q', interval=0.25)



