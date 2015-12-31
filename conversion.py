

from comtypes.client import CreateObject
from comtypes.gen import SpeechLib    
engine = CreateObject("SAPI.SpVoice")
stream = CreateObject("SAPI.SpFileStream")
infile = raw_input()
outfile = raw_input()
stream.Open(outfile, SpeechLib.SSFMCreateForWrite)
engine.AudioOutputStream = stream
f = open(infile, 'r')
theText = f.read()
f.close()
engine.speak(theText)
stream.Close()
