Attribute VB_Name = "ModInit"
Option Explicit

'Stream / Music handles

'estacion  01
Dim Strm1 As Long
Dim Msc1 As Long

'estacion 02
Dim Strm2 As Long
Dim Msc2 As Long

Function CloseDevice(WLHandle1 As String, WLHandle2 As String) As String

' Stop digital output
MsgBox "Stopping all Bass Streams and Music"
BASS_Stop

' Free the first handles
Select Case WLHandle1
    Case "Stream"
        MsgBox "Handle1 stream passed"
        BASS_StreamFree Strm1   'stream
        
    Case "Music"
        MsgBox "Handle1 music passed"
        BASS_MusicFree Msc1     'music
        
    Case "None"
        MsgBox "Handle1 None passed"
        'xxx    NOTHING
        
    Case Else
        MsgBox "Handle1 x passed"
        'xxx    NOTHING
End Select

' Free the second handles
Select Case WLHandle2
    Case "Stream"
        MsgBox "Handle2 stream passed"
        BASS_StreamFree Strm2   'stream
        
    Case "Music"
        MsgBox "Handle2 music passed"
        BASS_MusicFree Msc2     'music
        
    Case "None"
        MsgBox "Handle2 none passed"
        'xxx    NOTHING
        
    Case Else
        MsgBox "Handle2 x passed"
        'xxx    NOTHING
End Select

' Close digital sound system
MsgBox "now will free Bass driver"
BASS_Free
CloseDevice = "Ok"

End Function

Function InitDevice(hWnd As Long, InitParm As String) As String

Dim ParmResult As String

' Check that BASS 0.8 was loaded
If BASS_GetStringVersion <> "0.8" Then
    DisplayMsg "BASS version 0.8 was not loaded"
    InitDevice = "NotOk"
    Exit Function
End If

'Check that is original RMMC Control
Desencriptar WWDefPass, InitParm, ParmResult
If ParmResult = "RMMultimediaControl" Then
    GoSub InitComp
Else
    DisplayMsg "Can't initialize RadioMaker Multimedia Control"
    InitDevice = "NotOk"
    Exit Function
End If

InitComp:
' Initialize digital sound - default device, 44100hz, stereo, 16 bits
If BASS_Init(-1, 44100, 0, hWnd) = BASSFALSE Then
    DisplayMsg "Can't initialize digital sound system"
    InitDevice = "NotOk"
    Exit Function
End If

' Start digital output
If BASS_Start = BASSFALSE Then
    DisplayMsg "Can't start digital output"
    InitDevice = "NotOk"
    Exit Function
End If

InitDevice = "Ok"
End Function

Function InitDevice3D(hWnd As Long, InitParm As String) As String

Dim ParmResult As String

' Check that BASS 0.8 was loaded
If BASS_GetStringVersion <> "0.8" Then
    DisplayMsg "BASS version 0.8 was not loaded"
    InitDevice3D = "NotOk"
    Exit Function
End If

'Check that is original RMMC Control
Desencriptar WWDefPass, InitParm, ParmResult
If ParmResult = "RMMultimediaControl" Then
    GoSub InitComp
Else
    DisplayMsg "Can't initialize RadioMaker Multimedia Control"
    InitDevice3D = "NotOk"
    Exit Function
End If

InitComp:
' Initialize output device - default device, 44100hz, stereo, 16 bits, with 3D funtionality
If BASS_Init(-1, 44100, BASS_DEVICE_3D, hWnd) = BASSFALSE Then
    DisplayMsg "Can't initialize digital 3D sound system"
    InitDevice3D = "NotOk"
    Exit Function
End If

' Use meters as distance unit, 2x real world rolloff, real doppler effect
BASS_Set3DFactors 1, 2, 1

' Start digital output
If BASS_Start = BASSFALSE Then
    DisplayMsg "Can't start digital output"
    InitDevice3D = "NotOk"
    Exit Function
End If

InitDevice3D = "Ok"
End Function

