Attribute VB_Name = "RMBass"
'////////////////////////////////////////////////////////
'*
'*  //////// MULTIMEDIA & FX module for Vb.6+ ////////
'*  ** this module depends on 100% of "modBass.bas" **
'*  ********* and is for Radiomaker 1.0 only *********
'*
'*     Copyright (c) 1987-2008 Only development Inc.
'*
'///////////////////////////////////////////////////////

Option Explicit

Public Const StrTime = 1    '(stream) result in time
Public Const StrByte = 2    '(stream) result in byte
Public Const MscRowCol = 1  '(music) result in row/col
Public Const MscByte = 2    '(music) result in byte

'/// Stream / Music file handle dims
'/// Estacion  01
Public Strm1 As Long

'/// Estacion 02
Public Strm2 As Long
Public StrmLen As Long
Public Wvol As Long

Public Function Stream01GetVolumen() As Long

If Stream01IsPlaying = True Then
    Stream01GetVolumen = BASS_ChannelGetAttributes(Strm1, -1, Wvol, -101)
    If Stream01GetVolumen = BASSFALSE Then
        Stream01GetVolumen = 99
    End If
End If

End Function

Public Function Stream02GetVolumen() As Long

If Stream02IsPlaying = True Then
    Stream02GetVolumen = BASS_ChannelGetAttributes(Strm2, -1, Wvol, -101)
    If Stream02GetVolumen = BASSFALSE Then
        Stream02GetVolumen = 99
    End If
End If

End Function

Public Function Stream01SetVolumen(Wvol As Long) As Boolean

If Stream01IsPlaying = True Then
    If Wvol < 0 Or Wvol > 100 Then
        Stream01SetVolumen = False
    Else
        If BASS_ChannelSetAttributes(Strm1, -1, Wvol, -101) = BASSFALSE Then
            Stream01SetVolumen = False
        Else
            Stream01SetVolumen = True
        End If
    End If
End If

End Function

Public Function Stream02SetVolumen(Wvol As Long) As Boolean

If Stream02IsPlaying = True Then
    If Wvol < 0 Or Wvol > 100 Then
        Stream02SetVolumen = False
    Else
        If BASS_ChannelSetAttributes(Strm2, -1, Wvol, -101) = BASSFALSE Then
            Stream02SetVolumen = False
        Else
            Stream02SetVolumen = True
        End If
    End If
End If

End Function

Function FileLenGetBytesPS() As Double

'Esta funcion es para uso interno de la libreria.
Dim flags As Long, bps As Long

 On Error GoTo None
 Call BASS_ChannelGetAttributes(StrmLen, bps, 0, 0)
 
  flags = BASS_ChannelGetLength(StrmLen)
  
 If Not (flags & BASS_SAMPLE_MONO) Then bps = bps * 2
 If Not (flags & BASS_SAMPLE_8BITS) Then bps = bps * 2
 
 FileLenGetBytesPS = bps

Exit Function
 
None:
 FileLenGetBytesPS = 0
End Function

Public Function RoundDown(IntDone As Long, IntMax As Long, MaxAmount As Long) As Long
    
    Dim d As Long
    
    On Error Resume Next
    
    d = Int(MaxAmount * IntDone / IntMax)
    
    RoundDown = CInt(d)
    
End Function

Public Function Percentage(IntDone As Long, IntMax As Long) As Long
    
    Dim d As Long
    
    On Error Resume Next
    
    d = Int(100 * IntDone / IntMax)
    
    Percentage = CInt(d)
    
End Function

Public Function Stream02GetBytesPS() As Double

'Funcion solo para uso interno de la libreria
 Dim flags As Long, bps As Long
 
 On Error GoTo None
 Call BASS_ChannelGetAttributes(Strm2, bps, 0, 0)
 
 Stream02GetBytesPS = bps
Exit Function
 
None:
 Stream02GetBytesPS = 0
End Function

Function Stream01GetBytesPS() As Double

'Esta funcion es para uso interno de la libreria.
Dim flags As Long, bps As Long

 On Error GoTo None
 Call BASS_ChannelGetAttributes(Strm1, bps, 0, 0)
  
 Stream01GetBytesPS = bps

Exit Function
 
None:
 Stream01GetBytesPS = 0
End Function

Function Stream02IsPlaying() As Boolean

If BASS_ChannelIsActive(Strm2) = BASSTRUE Then
    Stream02IsPlaying = True
Else
    Stream02IsPlaying = False
End If

End Function

Function Stream01IsPlaying() As Boolean

If BASS_ChannelIsActive(Strm1) = BASSTRUE Then
    Stream01IsPlaying = True
Else
    Stream01IsPlaying = False
End If

End Function

Sub Stream02SetPosition(ByVal WPosOrWseg As Single, ByVal WType As Long)

Dim Rst As Long
Dim RstS As Long

'CHEQUEOS
Select Case WType
    Case StrTime '=1
        If Stream02IsPlaying = True Then
            RstS = BASS_ChannelSeconds2Bytes(Strm2, WPosOrWseg)
            Rst = BASS_ChannelGetLength(Strm2)
            If RstS > Rst Then  'compare is Ok
                DisplayMsg GetComLng_ByID(LNGDef, "1001"), " Stream02SetPosition > Pos_NoConv: " & WPosOrWseg & " Byte_Pos_Conv: " & RstS & " Byte_File_Lng:" & Rst, err.Number, True   '...posicion espec. incorrecta"
            Else
                If BASS_ChannelSetPosition(Strm2, RstS) = BASSFALSE Then
                    DisplayMsg GetComLng_ByID(LNGDef, "1001"), " Stream02SetPosition > Pos_NoConv: " & WPosOrWseg & " Byte_Pos_Conv: " & RstS & " Byte_File_Lng:" & Rst, err.Number, True   '...posicion espec. incorrecta"
                End If
            End If
        End If

    Case StrByte
        If Stream02IsPlaying = True Then
            Rst = BASS_ChannelGetLength(Strm2)
            If WPosOrWseg > Rst Then  'compare is Ok
                DisplayMsg GetComLng_ByID(LNGDef, "1001"), " Stream02SetPosition > Pos_Conv: " & WPosOrWseg & " Byte_File_Lng:" & Rst, err.Number, True   '...posicion espec. incorrecta"
            Else
                If BASS_ChannelSetPosition(Strm2, WPosOrWseg) = BASSFALSE Then
                    DisplayMsg GetComLng_ByID(LNGDef, "1001"), " Stream02SetPosition > Pos_Conv: " & WPosOrWseg & " Byte_File_Lng:" & Rst, err.Number, True   '...posicion espec. incorrecta"
                End If
            End If
        End If

End Select

End Sub

Sub Stream01SetPosition(ByVal WPosOrWseg As Single, ByVal WType As Long)

Dim Rst As Long
Dim RstS As Long

'CHEQUEOS
Select Case WType
    Case StrTime '=1
        If Stream01IsPlaying = True Then
            RstS = BASS_ChannelSeconds2Bytes(Strm1, WPosOrWseg)
            Rst = BASS_ChannelGetLength(Strm1)
            If RstS > Rst Then  'compare is Ok
                DisplayMsg GetComLng_ByID(LNGDef, "1001"), " Stream01SetPosition > Pos_NoConv: " & WPosOrWseg & " Byte_Pos_Conv: " & RstS & " Byte_File_Lng:" & Rst, err.Number, True   '...posicion espec. incorrecta"
            Else
                If BASS_ChannelSetPosition(Strm1, RstS) = BASSFALSE Then
                    DisplayMsg GetComLng_ByID(LNGDef, "1001"), " Stream01SetPosition > Pos_NoConv: " & WPosOrWseg & " Byte_Pos_Conv: " & RstS & " Byte_File_Lng:" & Rst, err.Number, True   '...posicion espec. incorrecta"
                End If
            End If
        End If
        
    Case StrByte '=2
        If Stream01IsPlaying = True Then
            Rst = BASS_ChannelGetLength(Strm1)
            If WPosOrWseg > Rst Then  'compare is Ok
                DisplayMsg GetComLng_ByID(LNGDef, "1001"), " Stream01SetPosition > Pos_Conv: " & WPosOrWseg & " Byte_File_Lng:" & Rst, err.Number, True   '...posicion espec. incorrecta"
            Else
                If BASS_ChannelSetPosition(Strm1, WPosOrWseg) = BASSFALSE Then
                    DisplayMsg GetComLng_ByID(LNGDef, "1001"), " Stream01SetPosition > Pos_Conv: " & WPosOrWseg & " Byte_File_Lng:" & Rst, err.Number, True   '...posicion espec. incorrecta"
                End If
            End If
        End If
        
End Select

End Sub

Function Stream02GetLen(ByVal WTypeDisplay As Long) As Long

Dim SByte As Long
Dim STime As Long

SByte = BASS_ChannelGetLength(Strm2)
STime = CLng(BASS_ChannelBytes2Seconds(Strm2, SByte))

Select Case WTypeDisplay
    Case StrByte
        Stream02GetLen = SByte
    
    Case StrTime
        Stream02GetLen = STime

End Select

End Function

Function Stream01GetLen(ByVal WTypeDisplay As Long) As Long

Dim SByte As Long
Dim STime As Long

SByte = BASS_ChannelGetLength(Strm1)
STime = CLng(BASS_ChannelBytes2Seconds(Strm1, SByte))

Select Case WTypeDisplay
    Case StrByte
        Stream01GetLen = SByte
    
    Case StrTime
        Stream01GetLen = STime

End Select

End Function

Sub CloseDevice(WLHandle1 As String, WLHandle2 As String)

' Stop digital output
BASS_Stop

' Free the first handle
Select Case WLHandle1
    Case "Stream"
        BASS_StreamFree Strm1  'stream
    Case "Music"
        'BASS_MusicFree Msc1a     'music
    Case Else
        'xxx    NOTHING
End Select

' Free the second handle
Select Case WLHandle2
    Case "Stream"
        BASS_StreamFree Strm2   'stream
    Case "Music"
        'BASS_MusicFree Msc2a     'music
    Case Else
        'xxx    NOTHING
End Select

' Close digital sound system
BASS_Free

End Sub

Function InitDevice(ByVal Whwnd As Long) As String

Dim ParmResult As String
Dim bi As BASS_INFO
Dim Mode1 As Long, Mode2 As Long, Mode3 As Long
Dim msg As String, Msg0 As String, Msg3 As String, Msg4 As String
Dim Style, Title, Response

' Check that BASS 1.4 was loaded

' check the correct BASS was loaded
If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
    DisplayMsg GetComLng_ByID(LNGDef, "1002"), " Check Lib Version Error", err.Number, True  'no se puede iniciar bass.dll
    'Call MsgBox("An incorrect version of BASS.DLL was loaded", vbCritical)
    InitDevice = "NotOk"
    Exit Function
End If

InitComp:
' Initialize digital sound - default device, 44100hz, stereo, 16 bits
If BASS_Init(-1, 44100, 0, Whwnd, 0) = BASSFALSE Then
    DisplayMsg GetComLng_ByID(LNGDef, "1003"), " InitDevice Error", err.Number, True  'Cant access to Digital Audio System
    InitDevice = "NotOk"
    Exit Function
End If

' Start digital output
If BASS_Start = BASSFALSE Then
    DisplayMsg GetComLng_ByID(LNGDef, "1003"), " Start digital output Error", err.Number, True 'Cant access to Digital Audio System
    InitDevice = "NotOk"
    Exit Function
End If

' check for DX8 drivers.
'bi.Size = LenB(bi)      'LenB(..) returns a byte data
Call BASS_GetInfo(bi)
If (bi.dsver < 8) Then
    Msg0 = GetComLng_ByID(LNGDef, "1004")   'DirectX 8 not initialized
    Msg3 = " "
    Msg4 = GetComLng_ByID(LNGDef, "1005")   'Do you want to continue anyway?
    msg = Msg0 & Chr(13) & Msg3 & Chr(13) & Msg4
    Style = vbYesNo + vbCritical + vbDefaultButton2
    Title = GetComLng_ByID(LNGDef, "1004")  'DirectX 8 not initialized
    Response = MsgBox(msg, Style, Title)
    If Response = vbYes Then
        'Est12Control.LblFX.Caption = "NoFX"
    Else
        BASS_Free
        Unload MainForm
    End If
End If

InitDevice = "Ok"
End Function

Function Stream02GetPosition(ByVal WTypeDisplay As Long) As Long

Dim PosByte As Long
Dim PosTime As Long

If Stream02IsPlaying = True Then
    PosByte = BASS_ChannelGetPosition(Strm2)  'stream file position (Bytes)
    PosTime = CLng(BASS_ChannelBytes2Seconds(Strm2, PosByte))
    
    Select Case WTypeDisplay
        Case StrByte
            Stream02GetPosition = PosByte
            
        Case StrTime
            Stream02GetPosition = PosTime

    End Select
Else
    Stream02GetPosition = 0
End If

End Function

Function Stream01GetPosition(ByVal WTypeDisplay As Long) As Long

Dim PosByte As Long
Dim PosTime As Long

If Stream01IsPlaying = True Then
    PosByte = BASS_ChannelGetPosition(Strm1)  'get stream file position (Bytes)
    PosTime = CLng(BASS_ChannelBytes2Seconds(Strm1, PosByte)) 'convert byte 2 sec.
    
    Select Case WTypeDisplay
        Case StrByte
            Stream01GetPosition = PosByte
            
        Case StrTime
            Stream01GetPosition = PosTime
            
    End Select
Else
    Stream01GetPosition = 0
End If

End Function

Sub Stream02Restart()

If BASS_ChannelSetPosition(Strm2, 0) = BASSFALSE Then
    DisplayMsg GetComLng_ByID(LNGDef, "1001"), " Stream02Restart error", err.Number, True '...posicion espec. incorrecta
    Exit Sub
End If

End Sub

Sub Stream01Restart()

If BASS_ChannelSetPosition(Strm1, 0) = BASSFALSE Then
    DisplayMsg GetComLng_ByID(LNGDef, "1001"), " Stream01Restart error", err.Number, True  '...posicion espec. incorrecta
    Exit Sub
End If

End Sub

Function Stream02Load(WFileName As String, LastHandle As String) As String

'retorna NotOk si hay algo mal
'retorna Stream (new handle) si fue satisfactorio

Dim StreamHandle2 As Long
Dim Mode1 As Long, Mode2 As Long, Mode3 As Long, Mode4 As Long, Mode5 As Long

'verificamos si hay un handle anterior y lo eliminamos
If LastHandle = "Music" Then
    'BASS_MusicFree Msc2     'music
Else
    If LastHandle = "Stream" Then
        Stream02Clear
    Else
        Stream02Clear
    End If
End If

Mode4 = BASS_MP3_SETPOS
Mode5 = BASS_SAMPLE_FX

StreamHandle2 = BASS_StreamCreateFile(BASSFALSE, WFileName, 0, 0, Mode4 Or Mode5)

If StreamHandle2 = 0 Then
    DisplayMsg GetComLng_ByID(LNGDef, "1006"), " Stream02Load > Filename: " & WFileName & " LastHandle: " & LastHandle & " FileMode: " & Mode4 & " - " & Mode5, err.Number, True 'Cant load specific audio file
    Stream02Load = "999"
Else
    Strm2 = StreamHandle2
    Stream02Load = "Stream"
End If

End Function

Function Stream01Load(WFileName As String, LastHandle As String) As String

'retorna NotOk si hay algo mal
'retorna Stream (new handle) si fue satisfactorio

Dim StreamHandle1 As Long
Dim Mode1 As Long, Mode2 As Long, Mode3 As Long, Mode4 As Long, Mode5 As Long

'verificamos si hay un handle anterior y lo eliminamos
If LastHandle = "Music" Then
    'BASS_MusicFree Msc1     'music
Else
    If LastHandle = "Stream" Then
        Stream01Clear
    Else
        Stream01Clear
    End If
End If

'gets the config device data

Mode4 = BASS_MP3_SETPOS
Mode5 = BASS_SAMPLE_FX

StreamHandle1 = BASS_StreamCreateFile(BASSFALSE, WFileName, 0, 0, Mode4 Or Mode5)

If StreamHandle1 = 0 Then
    DisplayMsg GetComLng_ByID(LNGDef, "1006"), " Stream01Load > Filename: " & WFileName & " LastHandle: " & LastHandle & " FileMode: " & Mode4 & " - " & Mode5, err.Number, True 'Cant load specific audio file
    Stream01Load = "999"
Else
    Strm1 = StreamHandle1
    Stream01Load = "Stream"
End If

End Function

Sub Stream02Clear()

'remove the last sync
'Result = StreamRmvSync(2)

BASS_StreamFree Strm2

End Sub

Sub Stream01Clear()

'removes the last sync
'Result = StreamRmvSync(1)

BASS_StreamFree Strm1

End Sub

Sub Stream02Stop()

' Stop the stream
If BASS_ChannelStop(Strm2) = BASSFALSE Then
    DisplayMsg GetComLng_ByID(LNGDef, "1007"), " Stream02Stop error", err.Number, True ' Cant stop file
    Exit Sub
End If

End Sub

Sub Stream01Stop()

'Stop the stream
If BASS_ChannelStop(Strm1) = BASSFALSE Then
    DisplayMsg GetComLng_ByID(LNGDef, "1007"), " Stream01Stop error", err.Number, True ' Cant stop file
    Exit Sub
End If

End Sub

Sub Stream02Play(ByVal WFlagStrmSample As Long)

'Play stream, not flushed
Select Case WFlagStrmSample
    Case BASS_SAMPLE_LOOP
        If BASS_ChannelPlay(Strm2, BASSFALSE) = BASSFALSE Then
            DisplayMsg GetComLng_ByID(LNGDef, "1008") & " " & GetComLng_ByID(LNGDef, "1009"), " Stream02Play error", err.Number, True  'Cant play the file in mode: + LOOP.
        End If
    Case Else
        If BASS_ChannelPlay(Strm2, BASSFALSE) = BASSFALSE Then
            DisplayMsg GetComLng_ByID(LNGDef, "1006"), " Stream02Play error", err.Number, True 'Cant load specific audio file
        End If
End Select

End Sub

Sub Stream01Play(ByVal WFlagStrmSample As Long)

'Play stream
Select Case WFlagStrmSample
    Case BASS_SAMPLE_LOOP
         If BASS_ChannelPlay(Strm1, BASSFALSE) = BASSFALSE Then
            DisplayMsg GetComLng_ByID(LNGDef, "1008") & " " & GetComLng_ByID(LNGDef, "1009"), " Stream01Play error", err.Number, True  'Cant play the file in mode: + LOOP.
        End If
    Case Else
        'Call BASS_ChannelPlay(Strm1, BASSFALSE)
        If BASS_ChannelPlay(Strm1, BASSFALSE) = BASSFALSE Then
            DisplayMsg GetComLng_ByID(LNGDef, "1006"), " Stream01Play error", err.Number, True 'Cant load specific audio file
        End If
End Select

End Sub
