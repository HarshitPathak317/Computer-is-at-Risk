do
strText=("Your Computer is at Risk")
set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak StrText
loop