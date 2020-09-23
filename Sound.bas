Attribute VB_Name = "Sound"
'Call API
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Const SND_ASYNC = &H1

'wHaTSoUnD
Public Enum wHaTSoUnD
    Whip = 1
    Chop = 2
    Glsbrk = 3
    ArwHt = 4
    Drop = 5
    Slap = 6
End Enum

'The sub that plays the sounds
Public Sub PlaySnd(Soundtype As wHaTSoUnD)
    Select Case Soundtype
        Case Whip
            Call PlaySound(App.Path + "\whip.wav", 0, SND_ASYNC)
        Case Chop
            Call PlaySound(App.Path + "\chop.wav", 0, SND_ASYNC)
        Case Glsbrk
            Call PlaySound(App.Path + "\glassbreaking.wav", 0, SND_ASYNC)
        Case ArwHt
            Call PlaySound(App.Path + "\Arrowhit.wav", 0, SND_ASYNC)
        Case Drop
            Call PlaySound(App.Path + "\drop.wav", 0, SND_ASYNC)
        Case Slap
            Call PlaySound(App.Path + "\slap.wav", 0, SND_ASYNC)
    End Select
End Sub

'What a useful little module!
'PLEASE VOTE!! At http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=40630&lngWId=1
