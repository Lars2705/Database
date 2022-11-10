#pip install PyAduio
#pip install pySpeech
#pip install speech Recognition
#pip install pypiwin32
#conda install pywin32

import win32com.client as win32

speaker = win32.Dispatch("SAPI.SpVoice")

while 1:
    print("Enter the word you want to speak it ouy by computer: ")
    s = input()
    speaker.Speak(s)

