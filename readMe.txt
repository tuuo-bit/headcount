headcount - Python GUI and Excel for Attendance Management

Taking attendance made fair, time efficient and easier by recognizing WiFi
SSID's in range rather than physical counting. This program is only intended
to be used a tool and by no means is a foolproof method for its shortcomings.

Only the following two lines from the sciprt require Admin privilges:
	...subprocess.run("netsh interface set interface name=\"Wi-Fi\" admin = disabled")
	...subprocess.run("netsh interface set interface name=\"Wi-Fi\" admin = enabled")
Turns ON and OFF the WiFi interface to scan for new networks in range.

Feature Description:
- Arrows: Switch between Home page and Main page
- Reload: refresh Main page values (SSIDs can be refreshed only with admin privileges)
- Summarize: displays attendance for date selected from drop-down menu
- Take Attendance: records attendance for the displays
- Remove: delete records of selected registered SSIDs from file
- Add: add selected SSIDs to record
- Register: register the new SSIDs with unique ID number

Python (3.12.3) Libraries Used:
- time
- subprocess
- pandas
- datetime
- os
- tkinter

Note:
- Make no changes to the "resource" files
- Excel files are modified in the same folder as the python file "gui.py"