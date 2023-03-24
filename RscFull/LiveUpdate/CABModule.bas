Attribute VB_Name = "CABMaker"
'/////////////////////////////////////////////
'*
'*  CabMaker File managger module
'*  Copyright (c) 1987-2007 Only development.
'*                all rights reserved.
'*  Christian A. Del Monte
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

'---
Public Const DefPass = "STPMaker1"
'---


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

    FrmDecompress.Prog1.Min = 0
    FrmDecompress.Prog1.Visible = True
    FrmDecompress.MousePointer = 11

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
        MsgBox "Formato de archivo no válido.", vbOKOnly, "Archivo Inválido"
        Close intBinaryFile
        ExtractCabFile = "NotOk"
        FrmDecompress.Prog1.Visible = False
        FrmDecompress.MousePointer = 0
        Exit Function
    End If
    
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    ReDim Src(FileHead.intNumFiles - 1)
    ReDim Dst(FileHead.intNumFiles - 1)
    ReDim FileDcm(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get intBinaryFile, , InfoHead
    
    FrmDecompress.Prog1.Max = UBound(InfoHead)
    
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
                FrmDecompress.Prog1.value = i
                DoEvents
            Next
    End Select
    
    'Close the binary file
    Close intBinaryFile
    'close the sample binary file
    Close intSampleFile
    FrmDecompress.Prog1.Visible = False
    FrmDecompress.MousePointer = 0
    ExtractCabFile = "Ok"
    Exit Function

ErrOut:

    'Display an error message if it didn't work
    MsgBox "No se puede decodificar el archivo cbm.", vbOKOnly, "Error"
    ExtractCabFile = "NotOk"
    FrmDecompress.Prog1.Visible = False
    FrmDecompress.MousePointer = 0
    
End Function
