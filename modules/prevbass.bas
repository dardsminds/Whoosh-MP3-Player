Attribute VB_Name = "BassPreview"
Declare Sub BASS_StreamFree Lib "bass.dll" (ByVal Handle As Long)
Declare Function BASS_StreamGetLength Lib "bass.dll" (ByVal Handle As Long) As Long
Declare Function BASS_StreamPlay Lib "bass.dll" (ByVal Handle As Long, ByVal flush As Long, ByVal Flags As Long) As Long
Declare Function BASS_Start Lib "bass.dll" () As Long
Declare Function BASS_Stop Lib "bass.dll" () As Long
Declare Sub BASS_Free Lib "bass.dll" ()
Declare Function BASS_Init Lib "bass.dll" (ByVal device As Long, ByVal Freq As Long, ByVal Flags As Long, ByVal win As Long) As Long
Declare Function BASS_StreamCreateFile Lib "bass.dll" (ByVal mem As Long, ByVal f As Any, ByVal offset As Long, ByVal length As Long, ByVal Flags As Long) As Long

