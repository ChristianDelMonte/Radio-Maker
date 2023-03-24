VERSION 5.00
Begin VB.UserControl RMMC 
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2160
   ScaleHeight     =   885
   ScaleWidth      =   2160
   ToolboxBitmap   =   "RMMC.ctx":0000
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      Picture         =   "RMMC.ctx":0312
      ScaleHeight     =   600
      ScaleWidth      =   2130
      TabIndex        =   0
      Top             =   0
      Width           =   2130
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MULTIMEDIA CONTROL"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   630
      Width           =   2175
   End
End
Attribute VB_Name = "RMMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'****************** RMMC ********************
'    RADIO MAKER MULTIMEDIA CONTROL OCX
'COPYRIGHT (C) 2001 CREACIONES DIGITALES INC.
'********************************************

Option Explicit

Const StrTime = 1
Const StrByte = 2
Const MscRowCol = 1
Const MscByte = 2

'Stream / Music handles
'estacion  01
Dim Strm1 As Long
Dim Msc1 As Long
'estacion 02
Dim Strm2 As Long
Dim Msc2 As Long

Public Function Stream02GetBytesPS() As Double

'Funcion solo para uso interno de la libreria
 Dim Flags As Long, BPS As Long
 
 On Error GoTo None
 Call BASS_ChannelGetAttributes(Strm2, BPS, 0, 0)
 
 Flags = BASS_ChannelGetFlags(Strm2)
  
 If Not (Flags & BASS_SAMPLE_MONO) Then BPS = BPS * 2
 If Not (Flags & BASS_SAMPLE_8BITS) Then BPS = BPS * 2
 
 Stream02GetBytesPS = BPS
Exit Function
 
None:
 Stream02GetBytesPS = 0
End Function

Function Stream01GetBytesPS() As Double

'Esta funcion es para uso interno de la libreria.
Dim Flags As Long, BPS As Long

 On Error GoTo None
 Call BASS_ChannelGetAttributes(Strm1, BPS, 0, 0)
 
 Flags = BASS_ChannelGetFlags(Strm1)
  
 If Not (Flags & BASS_SAMPLE_MONO) Then BPS = BPS * 2
 If Not (Flags & BASS_SAMPLE_8BITS) Then BPS = BPS * 2
 
 Stream01GetBytesPS = BPS

Exit Function
 
None:
 Stream01GetBytesPS = 0
End Function

Function Music02IsPlaying() As Boolean

If BASS_ChannelIsActive(Msc2) = BASSTRUE Then
    Music02IsPlaying = True
Else
    Music02IsPlaying = False
End If

End Function

Function Music01IsPlaying() As Boolean

If BASS_ChannelIsActive(Msc1) = BASSTRUE Then
    Music01IsPlaying = True
Else
    Music01IsPlaying = False
End If

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

Sub Music01SetPosition(WOrder As Long, WRow As Long)

Dim Rst As Long
Dim LOrder As Long
Dim LRow As Long

'CHEQUEOS
If BASS_ChannelIsActive(Msc1) = BASSTRUE Then
    Rst = BASS_MusicGetLength(Msc1)
    LOrder = GetLoWord(Rst)
    LRow = GetHiWord(Rst)
    
    If WOrder > LOrder Then
        'DisplayMsg "The play order can´t exceed the file order."
        DisplayMsg "No se puede reproducir el archivo desde la posición especificada."
        Exit Sub
    Else
        If BASS_ChannelSetPosition(Msc1, MakeLong(WOrder, WRow)) = BASSTRUE Then
            Exit Sub
        Else
            DisplayMsg "No se puede reproducir el archivo desde la posición especificada."
            Exit Sub
        End If
    End If
Else
    'nada que procesar
    Exit Sub
End If

End Sub

Sub Music02SetPosition(WOrder As Long, WRow As Long)

Dim Rst As Long
Dim LOrder As Long
Dim LRow As Long

'CHEQUEOS
If BASS_ChannelIsActive(Msc2) = BASSTRUE Then
    Rst = BASS_MusicGetLength(Msc2)
    LOrder = GetLoWord(Rst)
    LRow = GetHiWord(Rst)
    
    If WOrder > LOrder Then
        'DisplayMsg "The play order can´t exceed the file order."
        DisplayMsg "No se puede reproducir el archivo desde la posición especificada."
        Exit Sub
    Else
        If BASS_ChannelSetPosition(Msc2, MakeLong(WOrder, WRow)) = BASSTRUE Then
            Exit Sub
        Else
            DisplayMsg "No se puede reproducir el archivo desde la posición especificada."
            Exit Sub
        End If
    End If
Else
    'nada que procesar
    Exit Sub
End If

End Sub

Sub Stream02SetPosition(WPosOrWseg As Long, WType As Long)

Dim Rst As Long
Dim RstS As Long

'wtype contants
'Const StrTime = 1
'Const StrByte = 2

'CHEQUEOS
Select Case WType
    Case StrTime
        If BASS_ChannelIsActive(Strm2) = BASSTRUE Then
            RstS = (WPosOrWseg * 60000) * 3   'convert the seg into byte
            Rst = BASS_StreamGetLength(Strm2)   'get the real lenght (byte)
            If RstS > Rst Then  'compare is Ok
                DisplayMsg "No se puede reproducir el archivo desde la posición especificada."
                Exit Sub
            Else
                If BASS_ChannelSetPosition(Strm2, RstS) = BASSTRUE Then
                    Exit Sub
                Else
                    DisplayMsg "No se puede reproducir el archivo desde la posición especificada."
                    Exit Sub
                End If
            End If
        Else
            'nada que procesar
            Exit Sub
        End If

    Case StrByte
        If BASS_ChannelIsActive(Strm2) = BASSTRUE Then
            Rst = BASS_StreamGetLength(Strm2)   'get the real lenght (byte)
            If WPosOrWseg > Rst Then  'compare is Ok
                'DisplayMsg "The play position can´t exceed the file position."
                DisplayMsg "No se puede reproducir el archivo desde la posición especificada."
                Exit Sub
            Else
                If BASS_ChannelSetPosition(Strm2, WPosOrWseg) = BASSTRUE Then
                    Exit Sub
                Else
                    DisplayMsg "No se puede reproducir el archivo desde la posición especificada."
                    Exit Sub
                End If
            End If
        Else
            'nada que procesar
            Exit Sub
        End If
    
    Case Else
        'xxx
        'NOTHING TO DO
End Select

End Sub

Sub Stream01SetPosition(WPosOrWseg As Long, WType As Long)

Dim Rst As Long
Dim RstS As Long

'wtype contants
'Const StrTime = 1
'Const StrByte = 2

'CHEQUEOS
Select Case WType
    Case StrTime
        If BASS_ChannelIsActive(Strm1) = BASSTRUE Then
            RstS = (WPosOrWseg * 60000) * 3 'convert the seg into byte
            Rst = BASS_StreamGetLength(Strm1)   'get the real lenght (byte)
            If RstS > Rst Then  'compare is Ok
                DisplayMsg "No se puede reproducir el archivo desde la posición especificada."
                Exit Sub
            Else
                If BASS_ChannelSetPosition(Strm1, RstS) = BASSTRUE Then
                    Exit Sub
                Else
                    DisplayMsg "No se puede reproducir el archivo desde la posición especificada."
                    Exit Sub
                End If
            End If
        Else
            'nada que procesar
            Exit Sub
        End If

    Case StrByte
        If BASS_ChannelIsActive(Strm1) = BASSTRUE Then
            Rst = BASS_StreamGetLength(Strm1)   'get the real lenght (byte)
            If WPosOrWseg > Rst Then  'compare is Ok
                DisplayMsg "No se puede reproducir el archivo desde la posición especificada."
                Exit Sub
            Else
                If BASS_ChannelSetPosition(Strm1, WPosOrWseg) = BASSTRUE Then
                    Exit Sub
                Else
                    DisplayMsg "No se puede reproducir el archivo desde la posición especificada."
                    Exit Sub
                End If
            End If
        Else
            'nada que procesar
            Exit Sub
        End If
    
    Case Else
        'xxx
        'NOTHING TO DO
End Select

End Sub

Function Stream02GetLen(WTypeDisplay As Long) As String

'Const StrTime = 1
'Const StrByte = 2

Dim SLen As Long
Dim Seg As Long
Dim SSeg As Long
Dim BytesPS As Double

On Error GoTo None
SLen = BASS_StreamGetLength(Strm2)  'stream file lenght (Bytes)

Select Case WTypeDisplay
    Case StrByte
        Stream02GetLen = Trim(Str$(SLen))
    
    Case StrTime
        'Seg = (SLen / 60000) / 3    'segundos
        'SSeg = (SLen / 1000) Mod 60 'sentecimas de segundos
        'Stream02GetLen = Trim(Str$(Seg)) & "," & Trim(Str$(SSeg))

        BytesPS = Stream02GetBytesPS
        Stream02GetLen = Trim(Str$(SLen / BytesPS))

    Case Else
        Stream02GetLen = "0"
End Select
Exit Function

None:
Stream02GetLen = "0"
End Function

Function Stream01GetLen(WTypeDisplay As Long) As String

'Const StrTime = 1
'Const StrByte = 2

Dim SLen As Long
Dim Seg As Long
Dim SSeg As Long
Dim BytesPS As Double

On Error GoTo None
SLen = BASS_StreamGetLength(Strm1)  'stream file lenght (Bytes)

Select Case WTypeDisplay
    Case StrByte
        Stream01GetLen = Trim(Str$(SLen))
    
    Case StrTime
        'Seg = (SLen / 60000) / 3    'segundos
        'SSeg = (SLen / 1000) Mod 60   'sentecimas de segundos
        'Stream01GetLen = Trim(Str$(Seg)) & "," & Trim(Str$(SSeg))
        
        BytesPS = Stream01GetBytesPS
        Stream01GetLen = Trim(Str$(SLen / BytesPS))
        
    Case Else
        Stream01GetLen = "0"
End Select
Exit Function

None:
Stream01GetLen = "0"
End Function

Function Music02GetLen(WTypeDisplay As Long) As String

'Const MscRowCol = 1
'Const MscByte = 2

Dim MLen As Long
Dim LOrder As Long
Dim LRow As Long

On Error GoTo None
MLen = BASS_MusicGetLength(Msc2)  'music length (Order/Row)

Select Case WTypeDisplay
    Case MscByte
        Music02GetLen = Trim(Str$(MLen))
    
    Case MscRowCol
        LOrder = GetLoWord(MLen)
        LRow = GetHiWord(MLen)
        Music02GetLen = Trim(Str$(LOrder)) & "," & Trim(Str$(LRow))

    Case Else
        Music02GetLen = "0"
End Select
Exit Function

None:
Music02GetLen = "0"
End Function

Function Music01GetLen(WTypeDisplay As Long) As String

'Const MscRowCol = 1
'Const MscByte = 2

Dim MLen As Long
Dim LOrder As Long
Dim LRow As Long

On Error GoTo None
MLen = BASS_MusicGetLength(Msc1)  'music length (Order/Row)

Select Case WTypeDisplay
    Case MscByte
        Music01GetLen = Trim(Str$(MLen))
    
    Case MscRowCol
        LOrder = GetLoWord(MLen)
        LRow = GetHiWord(MLen)
        Music01GetLen = Trim(Str$(LOrder)) & "," & Trim(Str$(LRow))

    Case Else
        Music01GetLen = "0"
End Select
Exit Function

None:
Music01GetLen = "0"
End Function
Sub CloseDevice(WLHandle1 As String, WLHandle2 As String)

' Stop digital output
BASS_Stop

' Free the first handles
Select Case WLHandle1
    Case "Stream"
        BASS_StreamFree Strm1   'stream
    Case "Music"
        BASS_MusicFree Msc1     'music
    Case Else
        'xxx    NOTHING
End Select

' Free the second handles
Select Case WLHandle2
    Case "Stream"
        BASS_StreamFree Strm2   'stream
    Case "Music"
        BASS_MusicFree Msc2     'music
    Case Else
        'xxx    NOTHING
End Select

' Close digital sound system
BASS_Free

End Sub

Function InitDevice(hWnd As Long, WFlagDevice As Long, InitParm As String) As String

Dim ParmResult As String

' Check that BASS 0.8 was loaded
If BASS_GetStringVersion <> "1.1" Then
    DisplayMsg "No se puede inicializar la librería de audio."
    InitDevice = "NotOk"
    Exit Function
End If

'Check that is original RMMC Control
Desencriptar WWDefPass, InitParm, ParmResult
If ParmResult = "RMMultimediaControl" Then
    GoSub InitComp
Else
    DisplayMsg "No se puede inicializar RadioMaker Multimedia Control OCX"
    InitDevice = "NotOk"
    Exit Function
End If

InitComp:

'* Device setup flags *
'**********************
'Global Const BASS_DEVICE_8BITS = 1     'use 8 bit resolution, else 16 bit
'Global Const BASS_DEVICE_MONO = 2      'use mono, else stereo
'Global Const BASS_DEVICE_3D = 4        'enable 3D functionality
' If the BASS_DEVICE_3D flag is not specified when initilizing BASS,
' then the 3D flags (BASS_SAMPLE_3D and BASS_MUSIC_3D) are ignored when
' loading/creating a sample/stream/music.

' Initialize digital sound - default device, 44100hz, stereo, 16 bits
Select Case WFlagDevice
    Case BASS_DEVICE_8BITS     'use 8 bit resolution, else 16 bit
        If BASS_Init(-1, 44100, BASS_DEVICE_8BITS, hWnd) = BASSFALSE Then
            DisplayMsg "No se puede inicializar el sistema de sonido en modo 8bits"
            InitDevice = "NotOk"
            Exit Function
        End If

    Case BASS_DEVICE_MONO      'use mono, else stereo
        If BASS_Init(-1, 44100, BASS_DEVICE_MONO, hWnd) = BASSFALSE Then
            DisplayMsg "No se puede inicializar el sistema de sonido en modo MONO"
            InitDevice = "NotOk"
            Exit Function
        End If
    
    Case BASS_DEVICE_3D        'enable 3D functionality
        If BASS_Init(-1, 44100, BASS_DEVICE_3D, hWnd) = BASSFALSE Then
            DisplayMsg "No se puede inicializar el sistema de sonido en modo 3D"
            InitDevice = "NotOk"
            Exit Function
        End If
        BASS_Set3DFactors 1, 2, 1

    Case Else
        If BASS_Init(-1, 44100, 0, hWnd) = BASSFALSE Then
            DisplayMsg "No se puede inicializar el sistema de sonido"
            InitDevice = "NotOk"
            Exit Function
        End If
End Select

' Start digital output
If BASS_Start = BASSFALSE Then
    DisplayMsg "No se puede inicializar el sistema de sonido digital"
    InitDevice = "NotOk"
    Exit Function
End If

InitDevice = "Ok"
End Function

Sub Stream02SetPan(Wpan As Long)

If BASS_ChannelIsActive(Strm2) = BASSTRUE Then
    If Wpan < -100 Or Wpan > 100 Then
        DisplayMsg "Nivel de paneo inválido"
        Exit Sub
    End If
    If BASS_ChannelSetAttributes(Strm2, -1, -1, Wpan) = BASSFALSE Then
        DisplayMsg "No se puede cambiar el nivel de paneo"
        Exit Sub
    End If
Else
    Exit Sub
End If

End Sub

Sub Stream01SetPan(Wpan As Long)

If BASS_ChannelIsActive(Strm1) = BASSTRUE Then
    If Wpan < -100 Or Wpan > 100 Then
        DisplayMsg "Nivel de paneo inválido"
        Exit Sub
    End If

    If BASS_ChannelSetAttributes(Strm1, -1, -1, Wpan) = BASSFALSE Then
        DisplayMsg "No se puede cambiar el nivel de paneo"
        Exit Sub
    End If
Else
    Exit Sub
End If

End Sub

Sub Music02SetPan(Wpan As Long)

If BASS_ChannelIsActive(Msc2) = BASSTRUE Then
    If Wpan < -100 Or Wpan > 100 Then
        DisplayMsg "Nivel de paneo musical inválido"
        Exit Sub
    End If
    
    If BASS_ChannelSetAttributes(Msc2, -1, -1, Wpan) = BASSFALSE Then
        DisplayMsg "No se puede cambiar el nivel de paneo"
        Exit Sub
    End If
Else
    Exit Sub
End If

End Sub

Sub Music01SetPan(Wpan As Long)

If BASS_ChannelIsActive(Msc1) = BASSTRUE Then
    If Wpan < -100 Or Wpan > 100 Then
        DisplayMsg "Nivel de paneo musical inválido"
        Exit Sub
    End If
    
    If BASS_ChannelSetAttributes(Msc1, -1, -1, Wpan) = BASSFALSE Then
        DisplayMsg "No se puede cambiar el nivel de paneo"
        Exit Sub
    End If
Else
    Exit Sub
End If

End Sub

Sub Stream02SetVolume(WVol As Long)

If BASS_ChannelIsActive(Strm2) = BASSTRUE Then
    If WVol < 0 Or WVol > 100 Then
        DisplayMsg "Nivel de volumen inválido"
        Exit Sub
    End If
    If BASS_ChannelSetAttributes(Strm2, -1, WVol, -101) = BASSFALSE Then
        DisplayMsg "No se puede cambiar el nivel de volumen"
        Exit Sub
    End If
Else
    Exit Sub
End If

End Sub

Sub Stream01SetVolume(WVol As Long)

If BASS_ChannelIsActive(Strm1) = BASSTRUE Then
    If WVol < 0 Or WVol > 100 Then
        DisplayMsg "Nivel de volumen inválido"
        Exit Sub
    End If

    If BASS_ChannelSetAttributes(Strm1, -1, WVol, -101) = BASSFALSE Then
        DisplayMsg "No se puede cambiar el nivel de volumen"
        Exit Sub
    End If
Else
    Exit Sub
End If

End Sub

Sub Music02SetVolume(WVol As Long)

If BASS_ChannelIsActive(Msc2) = BASSTRUE Then
    If WVol < 0 Or WVol > 100 Then
        DisplayMsg "Nivel de volumen musical inválido"
        Exit Sub
    End If
    
    If BASS_ChannelSetAttributes(Msc2, -1, WVol, -101) = BASSFALSE Then
        DisplayMsg "No se puede cambiar el nivel de la musica"
        Exit Sub
    End If
Else
    Exit Sub
End If

End Sub

Sub Music01SetVolume(WVol As Long)

If BASS_ChannelIsActive(Msc1) = BASSTRUE Then
    If WVol < 0 Or WVol > 100 Then
        DisplayMsg "Nivel de volumen musical inválido"
        Exit Sub
    End If
    
    If BASS_ChannelSetAttributes(Msc1, -1, WVol, -101) = BASSFALSE Then
        DisplayMsg "No se puede cambiar el nivel de la musica"
        Exit Sub
    End If
Else
    Exit Sub
End If

End Sub

Sub DisplayMsg(Message As String)

'Display error dialogues
Dim ErrorNum As Long

ErrorNum = BASS_ErrorGetCode
MsgBox Message & vbCrLf & vbCrLf & " Error Codigo: " & ErrorNum & vbCrLf & BASS_GetErrorDescription(ErrorNum), vbCritical, "RMMC Error"

End Sub

Function Music02GetRIGHTLevel() As Long

Dim Level As Long
Dim RRRight As Long
Dim b

If BASS_ChannelIsActive(Msc2) = BASSTRUE Then
    Level = BASS_ChannelGetLevel(Msc2)  'music file level meter
    GoSub Process
Else
    GoSub ProcessTwo
End If

Process:
    b = 1
    If (b < 128) Then
        If GetHiWord(Level) >= b Then
            RRRight = GetHiWord(Level)
        Else
            RRRight = GetHiWord(Level)
            b = 2 * b - b / 2
        End If
    End If
    Music02GetRIGHTLevel = RRRight
Exit Function

ProcessTwo:
    Music02GetRIGHTLevel = 0

End Function

Function Music01GetRIGHTLevel() As Long

Dim Level As Long
Dim RRRight As Long
Dim b

If BASS_ChannelIsActive(Msc1) = BASSTRUE Then
    Level = BASS_ChannelGetLevel(Msc1)  'music file level meter
    GoSub Process
Else
    GoSub ProcessTwo
End If

Process:
    b = 1
    If (b < 128) Then
        If GetHiWord(Level) >= b Then
            RRRight = GetHiWord(Level)
        Else
            RRRight = GetHiWord(Level)
            b = 2 * b - b / 2
        End If
    End If
    Music01GetRIGHTLevel = RRRight
Exit Function

ProcessTwo:
    Music01GetRIGHTLevel = 0

End Function

Function Stream02GetRIGHTLevel() As Long

Dim Level As Long
Dim RRRight As Long
Dim b

If BASS_ChannelIsActive(Strm2) = BASSTRUE Then
    Level = BASS_ChannelGetLevel(Strm2)  'stream file level meter
    GoSub Process
Else
    GoSub ProcessTwo
End If

Process:
    b = 1
    If (b < 128) Then
        If GetHiWord(Level) >= b Then
            RRRight = GetHiWord(Level)
        Else
            RRRight = GetHiWord(Level)
            b = 2 * b - b / 2
        End If
    End If
    Stream02GetRIGHTLevel = RRRight
Exit Function

ProcessTwo:
    Stream02GetRIGHTLevel = 0

End Function

Function Stream01GetRIGHTLevel() As Long

Dim Level As Long
Dim RRRight As Long
Dim b

If BASS_ChannelIsActive(Strm1) = BASSTRUE Then
    Level = BASS_ChannelGetLevel(Strm1)  'stream file level meter
    GoSub Process
Else
    GoSub ProcessTwo
End If

Process:
    b = 1
    If (b < 128) Then
        If GetHiWord(Level) >= b Then
            RRRight = GetHiWord(Level)
        Else
            RRRight = GetHiWord(Level)
            b = 2 * b - b / 2
        End If
    End If
    Stream01GetRIGHTLevel = RRRight
Exit Function

ProcessTwo:
    Stream01GetRIGHTLevel = 0

End Function

Function Music02GetLEFTLevel() As Long

Dim Level As Long
Dim LLLeft As Long
Dim a

If BASS_ChannelIsActive(Msc2) = BASSTRUE Then
    Level = BASS_ChannelGetLevel(Msc2)  'Music file level meter
    GoSub Process
Else
    GoSub ProcessTwo
End If

Process:
    a = 93
    If (a > 0) Then
        If GetLoWord(Level) >= a Then
            LLLeft = GetLoWord(Level)
        Else
            LLLeft = GetLoWord(Level)
            a = a * 2 / 3
        End If
    End If
    Music02GetLEFTLevel = LLLeft
Exit Function

ProcessTwo:
    Music02GetLEFTLevel = 0

End Function

Function Music01GetLEFTLevel() As Long

Dim Level As Long
Dim LLLeft As Long
Dim a

If BASS_ChannelIsActive(Msc1) = BASSTRUE Then
    Level = BASS_ChannelGetLevel(Msc1)  'music file level meter
    GoSub Process
Else
    GoSub ProcessTwo
End If

Process:
    a = 93
    If (a > 0) Then
        If GetLoWord(Level) >= a Then
            LLLeft = GetLoWord(Level)
        Else
            LLLeft = GetLoWord(Level)
            a = a * 2 / 3
        End If
    End If
    Music01GetLEFTLevel = LLLeft
Exit Function

ProcessTwo:
    Music01GetLEFTLevel = 0

End Function

Function Stream02GetLEFTLevel() As Long

Dim Level As Long
Dim LLLeft As Long
Dim a

If BASS_ChannelIsActive(Strm2) = BASSTRUE Then
    Level = BASS_ChannelGetLevel(Strm2)  'stream file level meter
    GoSub Process
Else
    GoSub ProcessTwo
End If

Process:
    a = 93
    If (a > 0) Then
        If GetLoWord(Level) >= a Then
            LLLeft = GetLoWord(Level)
        Else
            LLLeft = GetLoWord(Level)
            a = a * 2 / 3
        End If
    End If
    Stream02GetLEFTLevel = LLLeft
Exit Function

ProcessTwo:
    Stream02GetLEFTLevel = 0

End Function

Function Music02GetPosition(WTypeDisplay As Long) As String

'Const MscRowCol = 1
'Const MscByte = 2

Dim Position As Long
Dim LOrder As Long
Dim LRow As Long

If BASS_ChannelIsActive(Msc2) = BASSTRUE Then
    Position = BASS_ChannelGetPosition(Msc2)  'music file position (Order/Row)
    GoSub Process
Else
    GoSub None
End If

Process:
Select Case WTypeDisplay
    Case MscByte
        Music02GetPosition = Trim(Str$(Position))
    
    Case MscRowCol
        LOrder = GetLoWord(Position)
        LRow = GetHiWord(Position)
        Music02GetPosition = Trim(Str$(LOrder)) & "," & Trim(Str$(LRow))

    Case Else
        Music02GetPosition = "0"
End Select
Exit Function

None:
Music02GetPosition = "0"
End Function

Function Music01GetPosition(WTypeDisplay As Long) As String

'Const MscRowCol = 1
'Const MscByte = 2

Dim Position As Long
Dim LOrder As Long
Dim LRow As Long

If BASS_ChannelIsActive(Msc1) = BASSTRUE Then
    Position = BASS_ChannelGetPosition(Msc1)  'music file position (Order/Row)
    GoSub Process
Else
    GoSub None
End If

Process:
Select Case WTypeDisplay
    Case MscByte
        Music01GetPosition = Trim(Str$(Position))
    
    Case MscRowCol
        LOrder = GetLoWord(Position)
        LRow = GetHiWord(Position)
        Music01GetPosition = Trim(Str$(LOrder)) & "," & Trim(Str$(LRow))

    Case Else
        Music01GetPosition = "0"
End Select
Exit Function

None:
Music01GetPosition = "0"
End Function

Function Stream02GetPosition(WTypeDisplay As Long) As String

'Const StrTime = 1
'Const StrByte = 2

Dim Position As Long
Dim Seg As Long
Dim SSeg As Long
Dim BytesPS As Double

If BASS_ChannelIsActive(Strm2) = BASSTRUE Then
    Position = BASS_ChannelGetPosition(Strm2)  'stream file position (Bytes)
    GoSub Process
Else
    GoSub None
End If

Process:
Select Case WTypeDisplay
    Case StrByte
        Stream02GetPosition = Trim(Str$(Position))
    
    Case StrTime
        'Seg = (Position / 60000) / 3    'segundos
        'SSeg = (Position / 1000) Mod 60 'sentecimas de segundos
        'Stream02GetPosition = Trim(Str$(Seg)) & "," & Trim(Str$(SSeg))

        On Error GoTo None
        BytesPS = Stream02GetBytesPS
        Stream02GetPosition = Trim(Str$(Position / BytesPS))

    Case Else
        Stream02GetPosition = "0"
End Select
Exit Function

None:
Stream02GetPosition = "0"
End Function

Sub Stream02Restart()

If BASS_ChannelSetPosition(Strm2, 0) = BASSFALSE Then
    DisplayMsg "No se puede reproducir el archivo desde su inicio."
    Exit Sub
End If

End Sub

Sub Stream01Restart()

If BASS_ChannelSetPosition(Strm1, 0) = BASSFALSE Then
    DisplayMsg "No se puede reproducir el archivo desde su inicio."
    Exit Sub
End If

End Sub

Sub Music02Restart()

' Play the music from the start
If BASS_MusicPlayEx(Msc2, 0, -1, BASSTRUE) = BASSFALSE Then
    DisplayMsg "No se puede reproducir el archivo desde su inicio."
    Exit Sub
End If

End Sub

Sub Music01Restart()

' Play the music from the start
If BASS_MusicPlayEx(Msc1, 0, -1, BASSTRUE) = BASSFALSE Then
    DisplayMsg "No se puede reproducir el archivo desde su inicio."
    Exit Sub
End If

End Sub

Function Music02Load(WFileName As String, WFlagMusic As Long, LastHandle As String) As String

'retorna NotOk si hay algo mal
'retorna Music (new handle) si fue satisfactorio

Dim ModHandle2 As Long

If WFileName = "" Or WFileName = " " Then
    Music02Load = "NotOk"
    Exit Function
End If

'verificamos si hay un handle anterior y lo eliminamos
If LastHandle = "Music" Then
    BASS_MusicFree Msc2     'music
Else
    If LastHandle = "Stream" Then
        BASS_StreamFree Strm2   'stream
    Else
        BASS_StreamFree Strm2   'stream
    End If
End If

Select Case WFlagMusic
    Case BASS_MUSIC_LOOP      ' loop music
        ModHandle2 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_LOOP)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle2 = 0 Then
            DisplayMsg "No se puede inicializar el archivo en modo LOOP"
            Music02Load = "NotOk"
            Exit Function
        Else
            Msc2 = ModHandle2
            Music02Load = "Music"
        End If

    Case BASS_MUSIC_FT2MOD    ' play .MOD as FastTracker 2 does
        ModHandle2 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_FT2MOD)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle2 = 0 Then
            DisplayMsg "No se puede cargar el módulo en modo FastTracker 2"
            Music02Load = "NotOk"
            Exit Function
        Else
            Msc2 = ModHandle2
            Music02Load = "Music"
        End If

    Case BASS_MUSIC_PT1MOD    ' play .MOD as ProTracker 1 does
        ModHandle2 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_PT1MOD)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle2 = 0 Then
            DisplayMsg "No se puede cargar el modulo en modo ProTracker 1"
            Music02Load = "NotOk"
            Exit Function
        Else
            Msc2 = ModHandle2
            Music02Load = "Music"
        End If

    Case BASS_MUSIC_MONO      ' force mono mixing (less CPU usage)
        ModHandle2 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_MONO)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle2 = 0 Then
            DisplayMsg "No se puede cargar el modulo en modo MONO"
            Music02Load = "NotOk"
            Exit Function
        Else
            Msc2 = ModHandle2
            Music02Load = "Music"
        End If

    Case BASS_MUSIC_3D       ' enable 3D functionality
        ModHandle2 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_3D)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle2 = 0 Then
            DisplayMsg "No se puede cargar el modulo en modo 3D"
            Music02Load = "NotOk"
            Exit Function
        Else
            Msc2 = ModHandle2
            Music02Load = "Music"
        End If

    Case BASS_MUSIC_POSRESET ' stop all notes when moving position
        ModHandle2 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_POSRESET)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle2 = 0 Then
            DisplayMsg "No se puede cargar el modulo en modo PosReset"
            Music02Load = "NotOk"
            Exit Function
        Else
            Msc2 = ModHandle2
            Music02Load = "Music"
        End If

    Case BASS_MUSIC_RAMPS      ' sensitive ramping
        ModHandle2 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_RAMPS)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle2 = 0 Then
            DisplayMsg "No se puede cargar el modulo en modo Sensitive Ramping"
            Music02Load = "NotOk"
            Exit Function
        Else
            Msc2 = ModHandle2
            Music02Load = "Music"
        End If

    Case BASS_MUSIC_RAMP       ' normal ramping
        ModHandle2 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_RAMP)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle2 = 0 Then
            DisplayMsg "No se puede cargar el modulo en modo Normal Ramping"
            Music02Load = "NotOk"
            Exit Function
        Else
            Msc2 = ModHandle2
            Music02Load = "Music"
        End If
        
    Case Else
        ModHandle2 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_RAMP)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle2 = 0 Then
            DisplayMsg "No se puede cargar el modulo seleccionado"
            Music02Load = "NotOk"
            Exit Function
        Else
            Msc2 = ModHandle2
            Music02Load = "Music"
        End If
End Select

End Function

Function Music01Load(WFileName As String, WFlagMusic As Long, LastHandle As String) As String

'retorna NotOk si hay algo mal
'retorna Music (new handle) si fue satisfactorio

Dim ModHandle1 As Long

If WFileName = "" Or WFileName = " " Then
    Music01Load = "NotOk"
    Exit Function
End If

'verificamos si hay un handle anterior y lo eliminamos
If LastHandle = "Music" Then
    BASS_MusicFree Msc1     'music
Else
    If LastHandle = "Stream" Then
        BASS_StreamFree Strm1   'stream
    Else
        BASS_StreamFree Strm1   'stream
    End If
End If

Select Case WFlagMusic
    Case BASS_MUSIC_LOOP      ' loop music
        ModHandle1 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_LOOP)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle1 = 0 Then
            DisplayMsg "No se puede cargar el modulo en modo Loop"
            Music01Load = "NotOk"
            Exit Function
        Else
            Msc1 = ModHandle1
            Music01Load = "Music"
        End If

    Case BASS_MUSIC_FT2MOD    ' play .MOD as FastTracker 2 does
        ModHandle1 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_FT2MOD)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle1 = 0 Then
            DisplayMsg "No se puede cargar el modulo en modo FastTracker 2"
            Music01Load = "NotOk"
            Exit Function
        Else
            Msc1 = ModHandle1
            Music01Load = "Music"
        End If

    Case BASS_MUSIC_PT1MOD    ' play .MOD as ProTracker 1 does
        ModHandle1 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_PT1MOD)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle1 = 0 Then
            DisplayMsg "No se puede cargar el modulo en modo ProTracker 1"
            Music01Load = "NotOk"
            Exit Function
        Else
            Msc1 = ModHandle1
            Music01Load = "Music"
        End If

    Case BASS_MUSIC_MONO      ' force mono mixing (less CPU usage)
        ModHandle1 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_MONO)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle1 = 0 Then
            DisplayMsg "No se puede cargar el modulo en modo MONO"
            Music01Load = "NotOk"
            Exit Function
        Else
            Msc1 = ModHandle1
            Music01Load = "Music"
        End If

    Case BASS_MUSIC_3D       ' enable 3D functionality
        ModHandle1 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_3D)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle1 = 0 Then
            DisplayMsg "No se puede cargar el modulo en modo 3D"
            Music01Load = "NotOk"
            Exit Function
        Else
            Msc1 = ModHandle1
            Music01Load = "Music"
        End If

    Case BASS_MUSIC_POSRESET ' stop all notes when moving position
        ModHandle1 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_POSRESET)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle1 = 0 Then
            DisplayMsg "No se puede cargar el modulo en modo PosReset"
            Music01Load = "NotOk"
            Exit Function
        Else
            Msc1 = ModHandle1
            Music01Load = "Music"
        End If

    Case BASS_MUSIC_RAMPS      ' sensitive ramping
        ModHandle1 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_RAMPS)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle1 = 0 Then
            DisplayMsg "No se puede cargar el modulo en modo Sensitive Ramping"
            Music01Load = "NotOk"
            Exit Function
        Else
            Msc1 = ModHandle1
            Music01Load = "Music"
        End If

    Case BASS_MUSIC_RAMP       ' normal ramping
        ModHandle1 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_RAMP)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle1 = 0 Then
            DisplayMsg "No se puede cargar el modulo en modo Ramping"
            Music01Load = "NotOk"
            Exit Function
        Else
            Msc1 = ModHandle1
            Music01Load = "Music"
        End If
        
    Case Else
        ModHandle1 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, BASS_MUSIC_RAMP)
        'BASS_MusicGetName (ModHandle1)
        If ModHandle1 = 0 Then
            DisplayMsg "No se puede cargar el modulo especificado"
            Music01Load = "NotOk"
            Exit Function
        Else
            Msc1 = ModHandle1
            Music01Load = "Music"
        End If
End Select

End Function

Function Stream02Load(WFileName As String, WStrmFlag As Long, LastHandle As String) As String

'retorna NotOk si hay algo mal
'retorna Stream (new handle) si fue satisfactorio

Dim StreamHandle2 As Long

If WFileName = "" Or WFileName = " " Then
    Stream02Load = "NotOk"
    Exit Function
End If

'verificamos si hay un handle anterior y lo eliminamos
If LastHandle = "Music" Then
    BASS_MusicFree Msc2     'music
Else
    If LastHandle = "Stream" Then
        BASS_StreamFree Strm2   'stream
    Else
        BASS_StreamFree Strm2   'stream
    End If
End If

'Flags para usar con BASS_StreamCreateFile
'Global Const BASS_SAMPLE_3D = 8                ' 3D functionality enabled
'Global Const BASS_MP3_HALFRATE = 65536         ' reduced quality MP3 (half sample rate)
'Global Const BASS_MP3_SETPOS = 131072          ' enable BASS_ChannelSetPosition on the MP3

Select Case WStrmFlag
    Case BASS_SAMPLE_3D                ' 3D functionality enabled
        StreamHandle2 = BASS_StreamCreateFile(BASSFALSE, WFileName, 0, 0, BASS_SAMPLE_3D)
        If StreamHandle2 = 0 Then
            DisplayMsg "No se puede cargar el archivo en modo 3D"
            Stream02Load = "NotOk"
            Exit Function
        Else
            Strm2 = StreamHandle2
            Stream02Load = "Stream"
        End If
    
    Case BASS_MP3_HALFRATE         ' reduced quality MP3 (half sample rate)
        StreamHandle2 = BASS_StreamCreateFile(BASSFALSE, WFileName, 0, 0, BASS_MP3_HALFRATE)
        If StreamHandle2 = 0 Then
            DisplayMsg "No se puede cargar el archivo Mp3 en modo Reduced Quality Stream"
            Stream02Load = "NotOk"
            Exit Function
        Else
            Strm2 = StreamHandle2
            Stream02Load = "Stream"
        End If
    
    Case BASS_MP3_SETPOS        'Enble BASS_ChannelSetPosition on the mp3
        StreamHandle2 = BASS_StreamCreateFile(BASSFALSE, WFileName, 0, 0, BASS_MP3_SETPOS)
        If StreamHandle2 = 0 Then
            DisplayMsg "No se puede cargar el archivo de audio seleccionado"
            Stream02Load = "NotOk"
            Exit Function
        Else
            Strm2 = StreamHandle2
            Stream02Load = "Stream"
        End If
    
    Case Else           'nothing, load on normal mode
        StreamHandle2 = BASS_StreamCreateFile(BASSFALSE, WFileName, 0, 0, 0)
        If StreamHandle2 = 0 Then
            DisplayMsg "No se puede cargar el archivo de audio seleccionado"
            Stream02Load = "NotOk"
            Exit Function
        Else
            Strm2 = StreamHandle2
            Stream02Load = "Stream"
        End If
    
End Select

End Function

Function Stream01Load(WFileName As String, WStrmFlag As Long, LastHandle As String) As String

'retorna NotOk si hay algo mal
'retorna Stream (new handle) si fue satisfactorio

Dim StreamHandle1 As Long

If WFileName = "" Or WFileName = " " Then
    Stream01Load = "NotOk"
    Exit Function
End If

'verificamos si hay un handle anterior y lo eliminamos
If LastHandle = "Music" Then
    BASS_MusicFree Msc1     'music
Else
    If LastHandle = "Stream" Then
        BASS_StreamFree Strm1   'stream
    Else
        BASS_StreamFree Strm1   'stream
    End If
End If

Select Case WStrmFlag
    Case BASS_SAMPLE_3D                ' 3D functionality enabled
        StreamHandle1 = BASS_StreamCreateFile(BASSFALSE, WFileName, 0, 0, BASS_SAMPLE_3D)
        If StreamHandle1 = 0 Then
            DisplayMsg "No se puede cargar el archivo en modo 3D"
            Stream01Load = "NotOk"
            Exit Function
        Else
            Strm1 = StreamHandle1
            Stream01Load = "Stream"
        End If
    
    Case BASS_MP3_HALFRATE         ' reduced quality MP3 (half sample rate)
        StreamHandle1 = BASS_StreamCreateFile(BASSFALSE, WFileName, 0, 0, BASS_MP3_HALFRATE)
        If StreamHandle1 = 0 Then
            DisplayMsg "No se puede cargar el archivo Mp3 en modo Reduced Quality Stream"
            Stream01Load = "NotOk"
            Exit Function
        Else
            Strm1 = StreamHandle1
            Stream01Load = "Stream"
        End If
    
    Case BASS_MP3_SETPOS         ' Enable BASS_ChannelSetPosition on the mp3
        StreamHandle1 = BASS_StreamCreateFile(BASSFALSE, WFileName, 0, 0, BASS_MP3_SETPOS)
        If StreamHandle1 = 0 Then
            DisplayMsg "No se puede cargar el archivo de audio especificado"
            Stream01Load = "NotOk"
            Exit Function
        Else
            Strm1 = StreamHandle1
            Stream01Load = "Stream"
        End If
    
    Case Else                   'nothing, normal mode
        StreamHandle1 = BASS_StreamCreateFile(BASSFALSE, WFileName, 0, 0, 0)
        If StreamHandle1 = 0 Then
            DisplayMsg "No se puede cargar el archivo de audio especificado"
            Stream01Load = "NotOk"
            Exit Function
        Else
            Strm1 = StreamHandle1
            Stream01Load = "Stream"
        End If
End Select

End Function

Sub Music02Clear()

BASS_MusicFree Msc2

End Sub

Sub Music01Clear()

BASS_MusicFree Msc1

End Sub

Sub Music02Stop()

If BASS_ChannelStop(Msc2) = BASSFALSE Then
    DisplayMsg "No se puede detener la reproduccion del modulo"
    Exit Sub
End If

End Sub

Sub Music01Stop()

If BASS_ChannelStop(Msc1) = BASSFALSE Then
    DisplayMsg "No se puede detener la reproduccion del modulo"
    Exit Sub
End If

End Sub

Sub Music02Play()

If BASS_MusicPlay(Msc2) = BASSFALSE Then
    DisplayMsg "No se puede reproducir el modulo"
    Exit Sub
End If

End Sub

Sub Music01Play()

If BASS_MusicPlay(Msc1) = BASSFALSE Then
    DisplayMsg "No se puede reproducir el modulo"
    Exit Sub
End If

End Sub

Sub Stream02Clear()

BASS_StreamFree Strm2

End Sub

Sub Stream01Clear()

BASS_StreamFree Strm1

End Sub

Sub Stream02Stop()

' Stop the stream
If BASS_ChannelStop(Strm2) = BASSFALSE Then
    DisplayMsg "No se puede detener la reproduccion del archivo"
    Exit Sub
End If

End Sub

Sub Stream01Stop()

'Stop the stream
If BASS_ChannelStop(Strm1) = BASSFALSE Then
    DisplayMsg "No se puede detener la reproduccion del archivo"
    Exit Sub
End If

End Sub

Sub Stream02Play(WFlagStrmSample As Long)

'Play stream, not flushed
Select Case WFlagStrmSample
    Case BASS_SAMPLE_8BITS
        If BASS_StreamPlay(Strm2, BASSFALSE, BASS_SAMPLE_8BITS) = BASSFALSE Then
            DisplayMsg "No se puede reproducir el archivo en modo 8 bits"
            Exit Sub
        End If
    Case BASS_SAMPLE_MONO
        If BASS_StreamPlay(Strm2, BASSFALSE, BASS_SAMPLE_MONO) = BASSFALSE Then
            DisplayMsg "No se puede reproducir el archivo en modo Mono"
            Exit Sub
        End If
    Case BASS_SAMPLE_LOOP
        If BASS_StreamPlay(Strm2, BASSFALSE, BASS_SAMPLE_LOOP) = BASSFALSE Then
            DisplayMsg "No se puede reproducir el archivo en modo Loop"
            Exit Sub
        End If
    Case BASS_SAMPLE_3D
        If BASS_StreamPlay(Strm2, BASSFALSE, BASS_SAMPLE_3D) = BASSFALSE Then
            DisplayMsg "No se puede reproducir el archivo en modo 3d"
            Exit Sub
        End If
    Case Else
        If BASS_StreamPlay(Strm2, BASSFALSE, 0) = BASSFALSE Then
            DisplayMsg "No se puede reproducir el arhivo"
            Exit Sub
        End If
End Select

End Sub

Sub Stream01Play(WFlagStrmSample As Long)

'Play stream, not flushed
Select Case WFlagStrmSample
    Case BASS_SAMPLE_8BITS
        If BASS_StreamPlay(Strm1, BASSFALSE, BASS_SAMPLE_8BITS) = BASSFALSE Then
            DisplayMsg "No se puede reproducir el archivo en modo 8 bits"
            Exit Sub
        End If
    Case BASS_SAMPLE_MONO
        If BASS_StreamPlay(Strm1, BASSFALSE, BASS_SAMPLE_MONO) = BASSFALSE Then
            DisplayMsg "No se puede reproducir el archivo en modo Mono"
            Exit Sub
        End If
    Case BASS_SAMPLE_LOOP
        If BASS_StreamPlay(Strm1, BASSFALSE, BASS_SAMPLE_LOOP) = BASSFALSE Then
            DisplayMsg "No se puede reproducir el archivo en modo Loop"
            Exit Sub
        End If
    Case BASS_SAMPLE_3D
        If BASS_StreamPlay(Strm1, BASSFALSE, BASS_SAMPLE_3D) = BASSFALSE Then
            DisplayMsg "No se puede reproducir el archivo en modo 3d"
            Exit Sub
        End If
    Case Else
        If BASS_StreamPlay(Strm1, BASSFALSE, 0) = BASSFALSE Then
            DisplayMsg "No se puede reproducir el arhivo"
            Exit Sub
        End If
End Select

End Sub

Function Stream01GetPosition(WTypeDisplay As Long) As String

'Const StrTime = 1
'Const StrByte = 2

Dim Position As Long
Dim Seg As Long
Dim SSeg As Long
Dim BytesPS As Double

If BASS_ChannelIsActive(Strm1) = BASSTRUE Then
    Position = BASS_ChannelGetPosition(Strm1)  'stream file position (Bytes)
    GoSub Process
Else
    GoSub None
End If

Process:
Select Case WTypeDisplay
    Case StrByte
        Stream01GetPosition = Trim(Str$(Position))
    
    Case StrTime
        'convertimos los bytes en
        'Seg = (Position / 60000) / 3    'segundos
        'SSeg = (Position / 1000) Mod 60  'sentecimas de segundos
        'Stream01GetPosition = Trim(Str$(Seg)) & "," & Trim(Str$(SSeg))

        On Error GoTo None
        BytesPS = Stream01GetBytesPS
        Stream01GetPosition = Trim(Str$(Position / BytesPS))

    Case Else
        Stream01GetPosition = "0"
End Select
Exit Function

None:
Stream01GetPosition = "0"
End Function

Function Stream01GetLEFTLevel() As Long

Dim Level As Long
Dim LLLeft As Long
Dim a

If BASS_ChannelIsActive(Strm1) = BASSTRUE Then
    Level = BASS_ChannelGetLevel(Strm1)  'stream file level meter
    GoSub Process
Else
    GoSub ProcessTwo
End If

Process:
    a = 93
    If (a > 0) Then
        If GetLoWord(Level) >= a Then
            LLLeft = GetLoWord(Level)
        Else
            LLLeft = GetLoWord(Level)
            a = a * 2 / 3
        End If
    End If
    Stream01GetLEFTLevel = LLLeft
Exit Function

ProcessTwo:
    Stream01GetLEFTLevel = 0

End Function
