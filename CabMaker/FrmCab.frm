VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmCab 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CABMaker v 1.0"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10350
   ForeColor       =   &H8000000F&
   Icon            =   "FrmCab.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "<< &Agregar Directorio"
      Height          =   315
      Left            =   8610
      TabIndex        =   38
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton CmdSaveIni 
      Height          =   375
      Left            =   6690
      Picture         =   "FrmCab.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Guardar datos en archivo INI y DAT"
      Top             =   3690
      Width           =   420
   End
   Begin VB.CheckBox ChkDat 
      Caption         =   "Crear archivo SETUP.DAT"
      Height          =   210
      Left            =   4200
      TabIndex        =   17
      Top             =   3885
      Width           =   2355
   End
   Begin VB.CheckBox CHKini 
      Caption         =   "Crear archivo SETUP.INI"
      Height          =   210
      Left            =   4200
      TabIndex        =   16
      Top             =   3660
      Width           =   2355
   End
   Begin VB.CommandButton CmdExtractAll 
      Caption         =   "Extraer &Todo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2790
      TabIndex        =   15
      Top             =   3690
      Width           =   1185
   End
   Begin MSComctlLib.ProgressBar Prog1 
      Height          =   180
      Left            =   3495
      TabIndex        =   12
      Top             =   4230
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   14
      Top             =   4185
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Text            =   "Archivo:"
            TextSave        =   "Archivo:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   8820
            MinWidth        =   8820
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1306
            MinWidth        =   1306
            TextSave        =   "09:44 p.m."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "07/01/2003"
         EndProperty
      EndProperty
      Enabled         =   0   'False
   End
   Begin VB.PictureBox PicForm 
      AutoSize        =   -1  'True
      Height          =   3450
      Left            =   45
      ScaleHeight     =   3390
      ScaleWidth      =   1800
      TabIndex        =   13
      Top             =   45
      Width           =   1860
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   39
         Top             =   30
         Width           =   285
      End
   End
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   2610
      Top             =   4725
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   4995
      TabIndex        =   8
      Top             =   675
      Width           =   2445
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   4995
      TabIndex        =   7
      Top             =   270
      Width           =   2490
   End
   Begin VB.FileListBox File1 
      DragIcon        =   "FrmCab.frx":040C
      Height          =   2430
      Left            =   7560
      TabIndex        =   6
      Top             =   1095
      Width           =   2760
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   9405
      TabIndex        =   5
      Top             =   3690
      Width           =   915
   End
   Begin VB.CommandButton CmdSave 
      Height          =   375
      Left            =   1035
      Picture         =   "FrmCab.frx":084E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guardar datos en archivo CBM"
      Top             =   3690
      Width           =   420
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   45
      Picture         =   "FrmCab.frx":0950
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Nuevo archivo CBM"
      Top             =   3690
      Width           =   420
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   540
      Picture         =   "FrmCab.frx":0A52
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Abrir archivo CBM"
      Top             =   3690
      Width           =   420
   End
   Begin MSComctlLib.ListView T1View 
      Height          =   3210
      Left            =   2025
      TabIndex        =   1
      Top             =   270
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   5662
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Dir"
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Archivos"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "DRM"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DRS"
         Object.Width           =   882
      EndProperty
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "&Extraer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1530
      TabIndex        =   0
      Top             =   3690
      Width           =   1185
   End
   Begin VB.Label Label13 
      Height          =   225
      Left            =   1290
      TabIndex        =   37
      Top             =   6945
      Width           =   915
   End
   Begin VB.Label Label12 
      Height          =   225
      Left            =   2340
      TabIndex        =   36
      Top             =   6945
      Width           =   915
   End
   Begin VB.Label Label11 
      Height          =   225
      Left            =   3360
      TabIndex        =   35
      Top             =   6960
      Width           =   915
   End
   Begin VB.Label Label10 
      Height          =   225
      Left            =   4365
      TabIndex        =   34
      Top             =   6960
      Width           =   915
   End
   Begin VB.Label Label9 
      Height          =   225
      Index           =   8
      Left            =   3285
      TabIndex        =   32
      Top             =   6375
      Width           =   915
   End
   Begin VB.Label Label9 
      Height          =   225
      Index           =   7
      Left            =   3285
      TabIndex        =   31
      Top             =   6090
      Width           =   915
   End
   Begin VB.Label Label9 
      Height          =   225
      Index           =   6
      Left            =   3285
      TabIndex        =   30
      Top             =   5805
      Width           =   915
   End
   Begin VB.Label Label9 
      Height          =   225
      Index           =   5
      Left            =   2295
      TabIndex        =   29
      Top             =   6360
      Width           =   915
   End
   Begin VB.Label Label9 
      Height          =   225
      Index           =   4
      Left            =   2295
      TabIndex        =   28
      Top             =   6075
      Width           =   915
   End
   Begin VB.Label Label9 
      Height          =   225
      Index           =   3
      Left            =   2295
      TabIndex        =   27
      Top             =   5790
      Width           =   915
   End
   Begin VB.Label Label9 
      Height          =   225
      Index           =   2
      Left            =   1275
      TabIndex        =   26
      Top             =   6345
      Width           =   915
   End
   Begin VB.Label Label9 
      Height          =   225
      Index           =   1
      Left            =   1275
      TabIndex        =   25
      Top             =   6060
      Width           =   915
   End
   Begin VB.Label Label9 
      Height          =   225
      Index           =   0
      Left            =   1260
      TabIndex        =   24
      Top             =   5775
      Width           =   915
   End
   Begin VB.Label Label8 
      Height          =   225
      Left            =   6420
      TabIndex        =   23
      Top             =   5430
      Width           =   915
   End
   Begin VB.Label Label7 
      Height          =   225
      Left            =   5400
      TabIndex        =   22
      Top             =   5430
      Width           =   915
   End
   Begin VB.Label Label6 
      Height          =   225
      Left            =   4380
      TabIndex        =   21
      Top             =   5415
      Width           =   915
   End
   Begin VB.Label Label5 
      Height          =   225
      Left            =   3330
      TabIndex        =   20
      Top             =   5415
      Width           =   915
   End
   Begin VB.Label Label4 
      Height          =   225
      Left            =   2310
      TabIndex        =   19
      Top             =   5415
      Width           =   915
   End
   Begin VB.Label Label3 
      Height          =   225
      Left            =   1290
      TabIndex        =   18
      Top             =   5400
      Width           =   915
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   30
      X2              =   10290
      Y1              =   3615
      Y2              =   3615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   45
      X2              =   10305
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label LblFile 
      BackColor       =   &H00000000&
      Height          =   240
      Left            =   3195
      TabIndex        =   11
      Top             =   4725
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de archivos a compactar"
      Height          =   240
      Left            =   2025
      TabIndex        =   10
      Top             =   45
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione los archivos que desea compactar y arrástrelos hacia la lista de archivos."
      Height          =   630
      Left            =   7560
      TabIndex        =   9
      Top             =   270
      Width           =   2745
   End
End
Attribute VB_Name = "FrmCab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Function DeployFile(FilePath As String, FileName As String)

Dim itmX As ListItem
Dim Onum
Dim NNum
Dim TxtKey As String
Dim NewKey As String
Dim Completo As String

Onum = T1View.ListItems.Count
NNum = Onum + 1
TxtKey = "r"
NewKey = TxtKey & NNum
  
Completo = FilePath & "\" & FileName

If T1View.ListItems.Count >= 120 Then
    Exit Function
Else
    Set itmX = T1View.ListItems.Add(NNum, NewKey, Completo) 'path & file
    itmX.SubItems(1) = FilePath & "\"     'file path
    itmX.SubItems(2) = FileName           'file name
    itmX.SubItems(3) = "0"
    itmX.SubItems(4) = "0"
End If

Label14.Caption = T1View.ListItems.Count

End Function


Sub WriteDataFile(WCABFileName As String)

Dim i As Integer
Dim x As Integer
Dim numfiles As Integer
Dim FileName As String

    '/////////////////////////////////////////////////////////////////////
    'escribimos el archivo setup.dat
    numfiles = FrmCab.T1View.ListItems.Count - 1
    AppData.CABFileName = StripFileFromDir(WCABFileName)

    If FrmCab.ChkDat.value = 1 Then
        For i = 0 To numfiles
            x = i + 1
            FrmCab.T1View.ListItems.Item(x).Selected = True
            FileName = FrmCab.T1View.SelectedItem.SubItems(2)
            AppData.FileName = FileName
            AppData.Destination = CInt(Trim(FrmCab.T1View.SelectedItem.SubItems(3)))
            AppData.DestNum = CInt(Trim(FrmCab.T1View.SelectedItem.SubItems(4)))
            AppData.Id = x
            'write the data to the file
            WriteAppFile AppData, StripDirFromFile(WCABFileName)
            'continue
            'FrmCab.Prog1.value = x
            'DoEvents
        Next i
    End If

    '////////////////////////////////////////////////////////////////////
    'escribimos el archivo setup.ini
    If FrmCab.CHKini.value = 1 Then
        IniData.AppTitle = FrmCab.Label3.Caption
        IniData.AppVersion = FrmCab.Label4.Caption
        IniData.AppCompany = FrmCab.Label5.Caption
        IniData.FrmTitle = FrmCab.Label6.Caption
        IniData.AppComment = FrmCab.Label7.Caption
        IniData.AppDefDir = FrmCab.Label8.Caption
        IniData.APPReadmeDesc = FrmCab.Label10.Caption
        IniData.APPReadmeName = FrmCab.Label11.Caption
        IniData.APPEXEDesc = FrmCab.Label12.Caption
        IniData.APPEXEName = FrmCab.Label13.Caption
        For i = 0 To 8
            IniData.AppDefSubDir.Dr(i) = FrmCab.Label9(i).Caption
        Next i
        IniData.Id = 1
        'write the data to the file
        WriteIniFile IniData, StripDirFromFile(WCABFileName)
        'continue
    End If

End Sub

Private Sub CHKini_Click()

FrmINI.Show

End Sub

Sub cmdExtract_Click()

Dim Folder As String
Dim Result As String
Dim TmpPath As String
Dim FName As String

TmpPath = App.Path & "\Temp"
FName = Trim(LblFile.Caption)

'select the folder to extract
Folder = BrowseForFolder("Seleccione el directorio de extracción")
If Folder = "" Or Folder = " " Then
    MsgBox "Oups!!"
    Exit Sub
End If
    
Folder = Left$(Folder, Len(Folder) - 1)

'extract the selected file from cab to temp folder
Result = ExtractCabFile(FName, Trim(Folder), TmpPath, 1)

If Result = "NotOk" Then
    MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
    Exit Sub
End If

End Sub

Private Sub CmdExtractAll_Click()

Dim Folder As String
Dim Result As String

'select the folder to extract
Folder = BrowseForFolder("Seleccione el directorio de extracción")
Folder = Left$(Folder, Len(Folder) - 1)

'extract all files from cab
Result = ExtractCabFile(Trim(LblFile.Caption), Trim(Folder), App.Path & "\Temp", 0)

If Result = "NotOk" Then
    MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
    Exit Sub
End If

End Sub


Private Sub CmdSave_Click()

Dim ConvertTX As String
Dim Result As String

On Error Resume Next
Cmd1.InitDir = App.Path
Cmd1.Filter = "Archivo CabMaker (*.cbm)|*.cbm|Archivo CabMaker"
Cmd1.DialogTitle = "CabFileMaker - Guardar un archivo"
Cmd1.CancelError = True
Cmd1.ShowSave

If Err.Number = 32755 Then Exit Sub

ConvertTX = Cmd1.FileName

'Lets open the file for save the new data
Result = SaveCabFile(ConvertTX)
If Result = "NotOK" Then
    MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
    Exit Sub
End If

LblFile.Caption = ConvertTX
StBar1.Panels(1).Text = "Archivo: " & LCase(ConvertTX)
CmdSave.Enabled = False
cmdExtract.Enabled = True
CmdExtractAll.Enabled = True
CmdSaveIni.Enabled = True

FrmCab.Caption = "ONLY CABMaker V." & App.Major & "." & App.Minor & " - Archivo: " & ConvertTX

End Sub

Private Sub CmdSaveIni_Click()

If Trim(LblFile.Caption = "") Or Trim(LblFile.Caption = " ") Then
    MsgBox "No se ha especificado el nombre del archivo CBM.", vbCritical, App.ProductName
    MsgBox "Guarde primero los datos del archivo CBM", vbCritical, App.ProductName
    Exit Sub
End If

Call WriteDataFile(Trim(LblFile.Caption))

End Sub

Private Sub Command1_Click()

Dim ConvertTX As String
Dim Result As String

On Error Resume Next
Cmd1.InitDir = App.Path
Cmd1.Filter = "Archivo CabMaker (*.cbm)|*.cbm|Archivo CabMaker"
Cmd1.DialogTitle = "CabFileMaker - Abrir un archivo"
Cmd1.CancelError = True
Cmd1.ShowOpen

If Err.Number = 32755 Then Exit Sub

ConvertTX = Cmd1.FileName

'Lets open the file for read the info
T1View.ListItems.Clear
Result = OpenCabFile(ConvertTX)
If Result = "NotOK" Then
    MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
    Exit Sub
End If

LblFile.Caption = ConvertTX
StBar1.Panels(1).Text = "Archivo: " & LCase(ConvertTX)
cmdExtract.Enabled = True
CmdExtractAll.Enabled = True
CmdSave.Enabled = False
CmdSaveIni.Enabled = False

FrmCab.Caption = "ONLY CABMaker V." & App.Major & "." & App.Minor & " - Archivo: " & ConvertTX

End Sub

Private Sub Command2_Click()

T1View.ListItems.Clear
LblFile.Caption = ""
StBar1.Panels(1).Text = "Archivo: "
cmdExtract.Enabled = False
CmdExtractAll.Enabled = False
CmdSave.Enabled = False
CmdSaveIni.Enabled = False

FrmCab.Caption = "ONLY CABMaker V." & App.Major & "." & App.Minor

End Sub

Private Sub Command3_Click()

Dim FPath As String, FName As String, TmpPath As String
Dim FileA As String, FileB As String, FileExt As String
Dim Result As String, z As Integer, numfiles As Integer


On Error Resume Next

    numfiles = File1.ListCount - 1

If numfiles >= 120 Then
    Prog1.Max = 120
    Prog1.Min = 0
    Prog1.Visible = True
    MousePointer = 11
    For z = 1 To 120
        TmpPath = App.Path & "\Temp"
        FPath = Dir1.Path
        FName = File1.List(z)
        'sets the file new name and path
        FileA = FPath & "\" & FName
        FileB = TmpPath & "\" & FName
        'lets compress the file (widt zlib)
        CompressFile FileA, FileB, 9
        'lets deploy the file into the file list box
        DeployFile TmpPath, FName
        Prog1.value = z
        DoEvents
    Next z
Else
    Prog1.Max = File1.ListCount - 1
    Prog1.Min = 0
    Prog1.Visible = True
    MousePointer = 11
    For z = 1 To File1.ListCount - 1
        TmpPath = App.Path & "\Temp"
        FPath = Dir1.Path
        FName = File1.List(z)
        'sets the file new name and path
        FileA = FPath & "\" & FName
        FileB = TmpPath & "\" & FName
        'lets compress the file (widt zlib)
        CompressFile FileA, FileB, 9
        'lets deploy the file into the file list box
        DeployFile TmpPath, FName
        Prog1.value = z
        DoEvents
    Next z
End If

CmdSave.Enabled = True
Prog1.Visible = False
FrmCab.MousePointer = 0

End Sub

Private Sub Command4_Click()

Dim TmpPath As String

On Error Resume Next
'lets clear the temp directory fisrt
TmpPath = App.Path & "\Temp"
Kill TmpPath & "\*.*"

End

End Sub


Private Sub Dir1_Change()

On Error Resume Next
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path

End Sub


Private Sub Drive1_Change()

Dir1.Path = Drive1.Drive

End Sub


Private Sub File1_Click()

File1.Path = Dir1.Path

End Sub

Private Sub File1_DblClick()

Dim WFN As String
Dim WFExt As String
Dim Result As String
Dim Strip As String
Dim StartF As Integer
Dim EndF As Integer

WFN = Dir1.Path & "\" & File1.FileName

'chequeos necesarios
If WFN = "" Or WFN = " " Then Exit Sub

WFExt = LCase(StripExtFromFile(WFN))

If WFExt = "cbm" Then
    T1View.ListItems.Clear
    Result = OpenCabFile(WFN)
    If Result = "NotOk" Then
        'xxx nothing to do
    Else
        FrmCab.LblFile.Caption = WFN
        FrmCab.StBar1.Panels(1).Text = "Archivo: " & LCase(WFN)
        FrmCab.cmdExtract.Enabled = True
        FrmCab.CmdExtractAll.Enabled = True
        FrmCab.CmdSave.Enabled = False
        FrmCab.Caption = "CABMaker v 1.0 - " & WFN
    End If
End If

End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

File1.DragIcon = File1.DragIcon
File1.Drag

End Sub

Private Sub Form_Initialize()

'ChDrive App.Path
'ChDir App.Path

End Sub

Private Sub Form_Load()

cmdExtract.Enabled = False
CmdSave.Enabled = False

PicForm.Picture = LoadResPicture("CBTM", 0)

'chequeamos la linea de comandos
Call ComLine(Command$)

On Error Resume Next
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path

Me.Caption = "ONLY CabMaker. V." & App.Major & "." & App.Minor
End Sub

Private Sub PicForm_Click()

About.Show

End Sub

Private Sub T1View_DblClick()

If T1View.ListItems.Count >= 1 Then
    FrmDirs.Show
End If

'If T1View.ListItems.Count >= 1 Then
'    If Trim(LblFile.Caption) = "" Or Trim(LblFile.Caption) = " " Then
'        Exit Sub
'    Else
'        Call cmdExtract_Click
'    End If
'End If

End Sub

Private Sub T1View_DragDrop(Source As Control, x As Single, Y As Single)

Dim FPath As String
Dim FName As String
Dim TmpPath As String
Dim FileA As String
Dim FileB As String
Dim FileExt As String
Dim Result As String

TmpPath = App.Path & "\Temp"
FPath = Dir1.Path
FName = File1.FileName
FileExt = StripExtFromFile(File1.FileName)

Select Case FileExt
    Case "cbm", "CBM", "Cbm", "cBm", "cbM"
        T1View.ListItems.Clear
        Result = OpenCabFile(FPath & "\" & FName)
        If Result = "NotOk" Then
            'xxx nothing to do
        Else
            FrmCab.LblFile.Caption = FPath & "\" & FName
            FrmCab.StBar1.Panels(1).Text = "Archivo: " & LCase(FPath & "\" & FName)
            FrmCab.cmdExtract.Enabled = True
            FrmCab.CmdExtractAll.Enabled = True
            FrmCab.CmdSave.Enabled = False
            FrmCab.Caption = "CABMaker v 1.0 - " & FPath & "\" & FName
        End If
    Case Else
        'sets the file new name and path
        FileA = FPath & "\" & FName
        FileB = TmpPath & "\" & FName
        'lets compress the file (widt zlib)
        CompressFile FileA, FileB, 9
        'lets deploy the file into the file list box
        DeployFile TmpPath, FName
        CmdSave.Enabled = True
End Select

End Sub
