VERSION 5.00
Begin VB.Form FrmCopy 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2535
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5640
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   3420
      Top             =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00404040&
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   2145
      TabIndex        =   5
      Top             =   2025
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H80000010&
      Height          =   825
      Left            =   135
      TabIndex        =   2
      Top             =   450
      Width           =   5370
      Begin VB.Label LblDst 
         Height          =   195
         Left            =   780
         TabIndex        =   8
         Top             =   495
         Width           =   4515
      End
      Begin VB.Label LblOrg 
         Height          =   195
         Left            =   780
         TabIndex        =   7
         Top             =   225
         Width           =   4515
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Destino:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   495
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Origen:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.PictureBox ProgessAll 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   180
      Picture         =   "FrmCopy.frx":0000
      ScaleHeight     =   300
      ScaleWidth      =   5250
      TabIndex        =   0
      Top             =   1395
      Width           =   5250
      Begin VB.PictureBox Progress 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   0
         Picture         =   "FrmCopy.frx":2734
         ScaleHeight     =   300
         ScaleWidth      =   5250
         TabIndex        =   1
         Top             =   0
         Width           =   5250
      End
   End
   Begin VB.Label Lb2 
      BackColor       =   &H000000FF&
      Caption         =   "Label5"
      Height          =   225
      Left            =   765
      TabIndex        =   10
      Top             =   2070
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Lb1 
      BackColor       =   &H000000FF&
      Caption         =   "Label4"
      Height          =   225
      Left            =   165
      TabIndex        =   9
      Top             =   2070
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      Height          =   390
      Left            =   135
      Top             =   1350
      Width           =   5355
   End
   Begin VB.Label LblTITLE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "xxx XXX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   5280
   End
End
Attribute VB_Name = "FrmCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Sub StartCopy()

'File dims
Dim i As Integer, UltimoRg As Integer
Dim CabName As String, FName As String
Dim FDest As Integer, FDest2 As Integer
Dim TMPath As String, ExPath As String, NewPath As String
Dim result As Boolean
Dim Num As Integer
'Status dims
Dim StatMax As Integer, StatMin As Integer, StatInterval As Integer

'disable the timer
Timer1.Interval = 0
Timer1.Enabled = False

'lets start the verifications
TMPath = Lb2.Caption: ExPath = Lb1.Caption

If Right$(TMPath, 1) = "\" Then
    TMPath = TMPath
Else
    TMPath = TMPath & "\"
End If
If Right$(ExPath, 1) = "\" Then
    ExPath = ExPath
Else
    ExPath = ExPath & "\"
End If

On Error Resume Next
MkDir TMPath
MkDir ExPath

'////////////////////////////////////////////////////////////////////
'cargamos los datos del archivo ini y creamos los directorios
Dim APData As StpIniFile
Dim Dirs(0 To 8) As String

APData = ReadIniFile(App.Path, 1)
'----
For i = 0 To 8
    Dirs(i) = Trim(Desencriptar(DefPass, APData.AppDefSubDir.Dr(i)))
    If Left$(Dirs(i), 10) = "My Sub Dir" Then
        'xxxx
    Else
        NewPath = ExPath & Dirs(i)
        MkDir NewPath
    End If
Next i
'////////////////////////////////////////////////////////////////////

'comenzamos la instalacion de la aplicacion
UltimoRg = GetAppFileLastReg(App.Path)

StatMin = 1
StatMax = 5250
StatInterval = StatMax / UltimoRg

CabName = Trim(Desencriptar(DefPass, ReadAppFile(App.Path, 1).CABFileName))

For i = 1 To UltimoRg
    FName = Trim(Desencriptar(DefPass, ReadAppFile(App.Path, i).FileName))
    NewPath = ""
    Select Case CInt(ReadAppFile(App.Path, i).Destination)
        Case STP_None, 0
            NewPath = ExPath
        Case STP_AppDir, 1
            NewPath = ExPath
        Case STP_WinDir, 2
            NewPath = sGetWinDir
        Case STP_WinSysDir, 3
            NewPath = sGetWinSysDir
        Case STP_AppSubDir, 4
            Num = CInt(ReadAppFile(App.Path, i).DestNum)
            NewPath = ExPath & Dirs(Num - 1)
    End Select
    'mas chequeos
    If Right$(NewPath, 1) = "\" Then
        NewPath = NewPath
    Else
        NewPath = NewPath & "\"
    End If
    'instalamos el archivo....
      LblOrg.Caption = CabName & ": " & FName     'origen
      LblDst.Caption = NewPath & FName       'destino
    result = ExtractCabFile(App.Path & "\" & CabName, NewPath, TMPath, 1, FName)
      If Progress.Width >= StatMax Then           'progreso de instalacion
        Progress.Width = Progress.Width
      Else
        Progress.Width = Progress.Width + StatInterval
      End If
    DoEvents
Next i

'removemos los datos del directorio temporal
On Error Resume Next
Kill TMPath & "\*.*"

frmShortCut.Show

End Sub


Private Sub Command1_Click()

Dim Msg, Msg0, Msg1, Msg2, Msg3, Msg4
Dim Style, Response, Title

Msg0 = "Esta seguro de que desea cancelar el proceso de"
Msg1 = "instalación de Radio Maker?."
Msg2 = "Si cancela puede que no funcione correctamente."
Msg3 = " "
Msg4 = "¿Desea cancelar de todas maneras?"
Msg = Msg0 & Chr(13) & Msg1 & Chr(13) & Msg2 & Chr(13) & Msg3 & Chr(13) & Msg4
Style = vbYesNo + vbCritical + vbDefaultButton2
Title = App.ProductName
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
    End
Else
    'xxxxx nothing to do.
End If

End Sub

Private Sub Form_Load()

Dim APData As StpIniFile
Dim APName As String, APVersion As String, APDir As String
Dim APCompany As String, APTitle As String

'//////////////////////////////////////////////////////////
'cargamos los datos del archivo ini
APData = ReadIniFile(App.Path, 1)
'----
APName = Trim(Desencriptar(DefPass, APData.AppTitle))
APVersion = Trim(Desencriptar(DefPass, APData.AppVersion))
LblTITLE.Caption = "Instalando " & APName
'//////////////////////////////////////////////////////////

Progress.Width = 1
Lb1.Caption = Trim(SetupInit.Text1.Text)    'path de instalacion (EXPath)
Lb2.Caption = Trim(SetupInit.Text2.Text)    'path temporal (TMPath)

Unload SetupInit

Timer1.Enabled = True
Timer1.Interval = 1000

FrmCopy.MousePointer = 11

End Sub

Private Sub Timer1_Timer()

Call StartCopy

End Sub
