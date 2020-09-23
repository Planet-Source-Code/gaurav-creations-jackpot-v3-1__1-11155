Attribute VB_Name = "Module1"
Global player1 As String
Global player2 As String
Global winner1 As Integer
Global winner2 As Integer
Public iHeight As Integer
Public iWidth As Integer

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function sndPlaySound Lib "winmm.dll" Alias _
"sndPlaySoundA" (ByVal lpszSoundName As String, ByVal _
uFlags As Long) As Long

