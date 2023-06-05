Dim wshl
Dim intSleep

intSleep = 60 'seconds

set wshl = createObject("WScript.Shell")
Do
	WScript.Sleep(intSleep*1000)
	wShl.SendKeys("{SCROLLLOCK 2}")
Loop
			