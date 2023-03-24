Attribute VB_Name = "NetShowVoice"
'********************* RM100 *********************
'    RADIO MAKER IMPORT/EXPORT FILE MODULE
'COPYRIGHT (C) 1987-2002 ONLY development inc.
'*************************************************

'---------------------------------------------------------------
'Modulo para procesar archivos Region/PlayList de NetShow Player
'y convertirlos en información CUE para su posterior proceso en
'Radio Maker.
'---------------------------------------------------------------
'Tambien funciona con exportaciones de tipo Windows Media
'Script file (.txt) desde Sony Sound Forge.
'---------------------------------------------------------------

Option Explicit

Public Function GetRegion(WFileName As String, WNumReg As Long, Dtype As Long) As String

Dim Wname As String
Dim TimeIn As Single, TimeOut As String
Dim Result As String
Dim StartM As String, EndM As String, Data As String
Dim ntexto As String

On Error GoTo NoGetAudio
Open WFileName For Input As #35

Do Until EOF(35)
Line Input #35, Data
    If Trim(Data) = "start_region_table" Or Trim(Data) = "end_region_table" Then
        'xxxxx nada
    Else
        If CLng(Right$(Data, 3)) = WNumReg Then
            Select Case Dtype
                Case 1  'start marker
                    StartM = Left$(Data, 8)    '=00:00:00.0 hh/mm/ss/s
                    TimeIn = ConvMinToSec(StartM)
                        GetRegion = Str$(TimeIn) & "," & Mid$(Data, 10, 1)
                    Close #35
                    Exit Do
                    Exit Function
                Case 2  'end marker
                    EndM = Mid$(Data, 12, 8)   '=00:00:00.0 hh/mm/ss/s
                    TimeOut = Str$(ConvMinToSec(EndM))
                        GetRegion = TimeOut & "," & Mid$(Data, 21, 1)
                    Close #35
                    Exit Do
                    Exit Function
            End Select
        End If
    End If
Loop
Close #35
Exit Function

NoGetAudio:
DisplayMsg GetComLng_ByID(LNGDef, "1010"), " Modulo NetShowVoice > error in function GetRegion. FileName: " & WFileName & " RegNum: " & WNumReg & " Dtipo: " & Dtype, err.Number, True   'RMVoice: File data audio region not found
Close #35
GetRegion = "999"
End Function
