Attribute VB_Name = "CABMaker"
'/////////////////////////////////////////////
'*
'*  CabMaker File managger module
'*  Copyright (c) 1987-2001 Only development.
'*                all rights reserved.
'*
'/////////////////////////////////////////////

Option Explicit

'This structure will describe our binary file's
'size and number of contained files
Private Type FILEHEADER
    FileType As String * 10     'the ident file
    intNumFiles As Integer      'How many files are inside?
    lngFileSize As Long         'How big is this file? (Used to check integrity)
End Type

'This structure will describe each file contained
'in our binary file
Private Type INFOHEADER
    lngFileSize As Long         'How big is this chunk of stored data?
    lngFileStart As Long        'Where does the chunk start?
    strFileName As String * 255  'What's the name of the file? (zlib compressed)
End Type

Private Type INFODATA
    BtFileData() As Byte
End Type

Sub ComLine(WFN As String)

'this is the command Line processor (NEW) (Very, very good... jeje)
Dim WFExt As String
Dim Result As String
Dim Strip As String
Dim StartF As Integer
Dim EndF As Integer

'chequeos necesarios
If WFN = "" Or WFN = " " Then Exit Sub

StartF = Len(WFN)
EndF = StartF - 2
Strip = Mid$(WFN, 2, EndF)

WFExt = LCase(StripExtFromFile(Strip))  'NEW!!!
'WFExt = Right$(Strip, 3)
'WFExt = LCase(WFExt)

If WFExt = "cbm" Then
    FrmCab.T1View.ListItems.Clear
    Result = OpenCabFile(Strip)
    If Result = "NotOk" Then
        End
    Else
        FrmCab.LblFile.Caption = Strip
        FrmCab.StBar1.Panels(1).Text = "Archivo: " & LCase(Strip)
        FrmCab.cmdExtract.Enabled = True
        FrmCab.CmdExtractAll.Enabled = True
        FrmCab.CmdSave.Enabled = False
        FrmCab.Caption = "CABMaker v 1.0 - " & Strip
    End If
End If

End Sub

'///////////////////////////////////////////////
'* WCabFileName:   Cabinet Filename
'* ExPath:         Path to extract the files
'* TempPath:       Temporal decompresion path
'* CantFiles:      0=Extract all files in CAB
'*                 1=Extract selected file only
'* OUT:            Ok or NotOk
'///////////////////////////////////////////////

Function ExtractCabFile(WCABFileName As String, ExPath As String, TempPath As String, CantFiles As Long) As String

Dim i As Integer
Dim x As Integer

Dim intSampleFile As Integer
Dim intBinaryFile As Integer
Dim bytSampleData() As Byte
Dim FileHead As FILEHEADER
Dim InfoHead() As INFOHEADER
Dim FileToExtract As String

Dim Src() As String   'file extracted to decompress
Dim Dst() As String   'destination file decompressed
Dim FileDcm() As String
                    
    'Set up the error handler
    'On Local Error GoTo ErrOut

    FrmCab.Prog1.Min = 0
    FrmCab.Prog1.Visible = True
    FrmCab.MousePointer = 11

    If WCABFileName = "" Or WCABFileName = " " Then
        ExtractCabFile = "NotOk"
        Exit Function
    End If
    
    'Open the binary file
    intBinaryFile = FreeFile
    Open WCABFileName For Binary Access Read Lock Write As intBinaryFile
    
    'Extract the FILEHEADER
    Get intBinaryFile, 1, FileHead

    'Check the file for validity
    If FileHead.FileType = "CabMakerFq" Then
        'xxxx all ok
    Else
        MsgBox "Formato de archivo CabMaker no válido.", vbOKOnly, "Archivo Inválido"
        Close intBinaryFile
        ExtractCabFile = "NotOk"
        FrmCab.Prog1.Visible = False
        FrmCab.MousePointer = 0
        Exit Function
    End If
    
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    ReDim Src(FileHead.intNumFiles - 1)
    ReDim Dst(FileHead.intNumFiles - 1)
    ReDim FileDcm(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get intBinaryFile, , InfoHead
    
    FrmCab.Prog1.Max = UBound(InfoHead)
    
    'check for file to extract
    Select Case CantFiles
        Case 0  '*********************************** extract all files in cab
            For i = 0 To UBound(InfoHead)
                'Resize the byte data array
                ReDim bytSampleData(InfoHead(i).lngFileSize - 1)
                'Get the data
                x = i + 1
                Get intBinaryFile, (InfoHead(i).lngFileStart + (10 * x)), bytSampleData
                'Open a new file and store the data
                intSampleFile = FreeFile
                Open TempPath & "\" & Trim(Desencriptar(DefPass & Trim(Str(i)), InfoHead(i).strFileName)) For Binary Access Write Lock Write As intSampleFile
                Put intSampleFile, 1, bytSampleData
                Close intSampleFile
                'lets decompress the file extracted
                FileDcm(i) = Trim(Desencriptar(DefPass & Trim(Str(i)), InfoHead(i).strFileName))
                Src(i) = TempPath & "\" & Trim(FileDcm(i))
                Dst(i) = Trim(ExPath) & "\" & Trim(FileDcm(i))
                'decompress the compressed file (zlib)
                UnCompressFile Src(i), Dst(i)
                FrmCab.Prog1.value = i
                DoEvents
            Next
        Case 1  '********************************** extract selected file only
            For i = 0 To UBound(InfoHead)
                'Resize the byte data array
                ReDim bytSampleData(InfoHead(i).lngFileSize - 1)
                'Get the data
                x = i + 1
                Get intBinaryFile, (InfoHead(i).lngFileStart + (10 * x)), bytSampleData
                'compare the filename
                FileToExtract = FrmCab.T1View.SelectedItem.ListSubItems(2).Text
                If LCase(Trim(InfoHead(i).strFileName)) = LCase(Trim(FileToExtract)) Then
                    'Open a new file and store the data
                    intSampleFile = FreeFile
                    Open TempPath & "\" & Trim(Desencriptar(DefPass & Trim(Str(i)), InfoHead(i).strFileName)) For Binary Access Write Lock Write As intSampleFile
                    Put intSampleFile, 1, bytSampleData
                    Close intSampleFile
                    'lets decompress the file extracted
                    FileDcm(i) = Trim(Desencriptar(DefPass & Trim(Str(i)), InfoHead(i).strFileName))
                    Src(i) = TempPath & "\" & Trim(FileDcm(i))
                    Dst(i) = Trim(ExPath) & "\" & Trim(FileDcm(i))
                    'decompress the compressed file (zlib)
                    UnCompressFile Src(i), Dst(i)
                    FrmCab.Prog1.value = i
                    DoEvents
                    Exit For
                Else
                    FrmCab.Prog1.value = i
                    DoEvents
                    'xxxx continue the search
                End If
            Next
    End Select
    
    'Close the binary file
    Close intBinaryFile
    'close the sample binary file
    Close intSampleFile
    FrmCab.Prog1.Visible = False
    FrmCab.MousePointer = 0
    ExtractCabFile = "Ok"
    Exit Function

ErrOut:

    'Display an error message if it didn't work
    MsgBox "No se puede decodificar el archivo cbm.", vbOKOnly, "Error"
    ExtractCabFile = "NotOk"
    FrmCab.Prog1.Visible = False
    FrmCab.MousePointer = 0
    
End Function

'///////////////////////////////////////////////
'* WCabFileName:   Cabinet Filename to open
'* OUT:            Ok or NotOk
'///////////////////////////////////////////////

Function OpenCabFile(WCABFileName As String) As String

Dim i As Integer
Dim x As Integer

Dim intSampleFile As Integer
Dim intBinaryFile As Integer
Dim bytSampleData() As Byte
Dim FileHead As FILEHEADER
Dim InfoHead() As INFOHEADER
    
    'Set up the error handler
    On Local Error GoTo ErrOut

    FrmCab.Prog1.Min = 0
    FrmCab.Prog1.Visible = True
    FrmCab.MousePointer = 11

    'Open the binary file
    intBinaryFile = FreeFile
    Open WCABFileName For Binary Access Read Lock Write As intBinaryFile
    
    'Extract the FILEHEADER
    Get intBinaryFile, 1, FileHead
    
    'Check the file for validity
    If FileHead.FileType = "CabMakerFq" Then
        'xxxx all ok
    Else
        MsgBox "Formato de archivo CabMaker no válido.", vbOKOnly, "Archivo Inválido"
        Close intBinaryFile
        OpenCabFile = "NotOk"
        Exit Function
    End If
    
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get intBinaryFile, , InfoHead
    FrmCab.Prog1.Max = UBound(InfoHead)

    'Extract all of the files from the binary file
    For i = 0 To UBound(InfoHead)
        'Resize the byte data array
        ReDim bytSampleData(InfoHead(i).lngFileSize - 1)
        'Get the data
        x = i + 1
        Get intBinaryFile, (InfoHead(i).lngFileStart + (10 * x)), bytSampleData
        'gets the file name
        Dim CabIntFName As String
        CabIntFName = Trim(Desencriptar(DefPass & Trim(Str(i)), InfoHead(i).strFileName))
        Call FrmCab.DeployFile(App.Path, CabIntFName)
        FrmCab.Prog1.value = i
        DoEvents
    Next
    
    'Close the binary file
    Close intBinaryFile
    FrmCab.Prog1.Visible = False
    FrmCab.MousePointer = 0
    OpenCabFile = "Ok"
    'Exit before we hit the error handler
    Exit Function

ErrOut:

    'Display an error message if it didn't work
    MsgBox "No de puede decodificar el archivo cbm.", vbOKOnly, "Error"
    OpenCabFile = "NotOk"
    FrmCab.Prog1.Visible = False
    FrmCab.MousePointer = 0

End Function

'///////////////////////////////////////////////
'* WCabFileName:   Cabinet Filename to save
'* OUT:            Ok or NotOk
'///////////////////////////////////////////////

Function SaveCabFile(WCABFileName As String) As String

'dimensions
Dim intBinaryFile As Integer
Dim FileHead As FILEHEADER
Dim InfoHead() As INFOHEADER
Dim lngFileStart As Long

'new
Dim FileName As String
Dim FilePath As String
Dim NumFiles As Integer
Dim Fl As Integer
Dim File() As Integer
Dim FileN() As String * 255
Dim bytSampleData() As Byte
Dim Data As Integer
Dim DataFile() As INFODATA
Dim NIndex
Dim FSize As Long
Dim i As Integer
Dim x As Integer

    'Set up the error handler
    'On Error GoTo ErrOut

    'Find some free file numbers to use and open the files
    NumFiles = FrmCab.T1View.ListItems.Count - 1
    ReDim File(0 To NumFiles)
    ReDim FileN(0 To NumFiles)
    ReDim DataFile(0 To NumFiles)
    ReDim LenData(0 To NumFiles)
    
    'Set up the file header
    FileHead.intNumFiles = NumFiles + 1
    
    FrmCab.Prog1.Max = NumFiles
    FrmCab.Prog1.Min = 0
    FrmCab.Prog1.Visible = True
    FrmCab.MousePointer = 11
    
    For Fl = 0 To NumFiles
        NIndex = Fl + 1
        FrmCab.T1View.ListItems.Item(NIndex).Selected = True
        FilePath = FrmCab.T1View.SelectedItem.SubItems(1)
        FileName = FrmCab.T1View.SelectedItem.SubItems(2)
        File(Fl) = FreeFile
        FileN(Fl) = FileName
        Open FilePath & FileName For Binary Access Read Lock Write As File(Fl)
        'Find out how large the files are and
        'resize the data arrays appropriately
        ReDim DataFile(Fl).BtFileData(LOF(File(Fl)) - 1)
        'Get the data from the files
        Get File(Fl), 1, DataFile(Fl).BtFileData
        'setup the total file size
        FileHead.lngFileSize = FileHead.lngFileSize + (UBound(DataFile(Fl).BtFileData) + 1)
        'Close the files
        Close File(Fl)
        FrmCab.Prog1.value = Fl
        DoEvents
    Next Fl
    
    FileHead.lngFileSize = (FileHead.lngFileSize + 16) + (FileHead.intNumFiles * 263) + (FileHead.intNumFiles + 10)
    '**************************************************
    'NOTE:
    'The '16' added to lngFileSize represents the size
    'of the FILEHEADER structure - 16bytes. You can
    'determine the size of a structure by examining
    'the data types it uses. A LONG uses 4 bytes, an
    'INTEGER uses 2, a STRING uses 1byte per character,
    'etc. The INFOHEADER structure takes up 24 bytes
    'for each entry, hence we multiply 24 by the number
    'of entries (intNumFiles).
    '**************************************************

    FrmCab.Prog1.value = 0

    'Set up the info headers
    ReDim InfoHead(FileHead.intNumFiles - 1)
    lngFileStart = (16) + (FileHead.intNumFiles * 263) + 1
    InfoHead(0).lngFileStart = lngFileStart
    InfoHead(0).lngFileSize = (UBound(DataFile(0).BtFileData) + 1)
    InfoHead(0).strFileName = Encriptar(DefPass & "0", FileN(0))
        
    For Fl = 1 To NumFiles
        InfoHead(Fl).lngFileSize = (UBound(DataFile(Fl).BtFileData) + 1)
            lngFileStart = lngFileStart + (InfoHead(Fl - 1).lngFileSize)
        InfoHead(Fl).lngFileStart = lngFileStart
        InfoHead(Fl).strFileName = Encriptar(DefPass & Trim(Str(Fl)), FileN(Fl))
        FrmCab.Prog1.value = Fl
        DoEvents
    Next Fl
    
    'create the CabMaker File and save the new data
    intBinaryFile = FreeFile
    Open WCABFileName For Binary Access Write Lock Write As intBinaryFile
    
    'Store the data in the file
    FileHead.FileType = "CabMakerFq"
    Put intBinaryFile, 1, FileHead
    Put intBinaryFile, , InfoHead
    For Fl = 0 To NumFiles
        If Fl > NumFiles Then Exit For
        Put intBinaryFile, , DataFile(Fl)
        FrmCab.Prog1.value = Fl
        DoEvents
    Next Fl
    
    'Close the file
    Close intBinaryFile
    FrmCab.Prog1.Visible = False
    FrmCab.MousePointer = 0
    
    SaveCabFile = "Ok"
    'Exit before we hit the error handler
    Exit Function
    
ErrOut:

    'Display an error message if it didn't work
    MsgBox "No se puede codificar el archivo cbm.", vbOKOnly, "Error"
    SaveCabFile = "NotOk"
    FrmCab.Prog1.Visible = False
    FrmCab.MousePointer = 0
    
End Function
