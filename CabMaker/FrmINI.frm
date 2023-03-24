VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmINI 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Opciones INI"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4815
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   5250
      Top             =   4005
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Directorios"
      Height          =   3945
      Left            =   90
      TabIndex        =   13
      Top             =   4320
      Width           =   4650
      Begin VB.TextBox Text7 
         Height          =   285
         Index           =   8
         Left            =   1425
         TabIndex        =   32
         Text            =   "My Sub Dir 9"
         Top             =   3510
         Width           =   3090
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Index           =   7
         Left            =   1425
         TabIndex        =   30
         Text            =   "My Sub Dir 8"
         Top             =   3165
         Width           =   3090
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Index           =   6
         Left            =   1425
         TabIndex        =   28
         Text            =   "My Sub Dir 7"
         Top             =   2820
         Width           =   3090
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Index           =   5
         Left            =   1425
         TabIndex        =   26
         Text            =   "My Sub Dir 6"
         Top             =   2475
         Width           =   3090
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Index           =   4
         Left            =   1425
         TabIndex        =   24
         Text            =   "My Sub Dir 5"
         Top             =   2130
         Width           =   3090
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Index           =   3
         Left            =   1425
         TabIndex        =   22
         Text            =   "My Sub Dir 4"
         Top             =   1785
         Width           =   3090
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Index           =   2
         Left            =   1425
         TabIndex        =   20
         Text            =   "My Sub Dir 3"
         Top             =   1440
         Width           =   3090
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Index           =   1
         Left            =   1425
         TabIndex        =   18
         Text            =   "My Sub Dir 2"
         Top             =   1080
         Width           =   3090
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Index           =   0
         Left            =   1425
         TabIndex        =   16
         Text            =   "My Sub Dir 1"
         Top             =   735
         Width           =   3090
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1425
         TabIndex        =   14
         Text            =   "My APP Dir"
         Top             =   375
         Width           =   3090
      End
      Begin VB.Label Label7 
         Caption         =   "Subdirectorio 9:"
         Height          =   225
         Index           =   8
         Left            =   150
         TabIndex        =   33
         Top             =   3540
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "Subdirectorio 8:"
         Height          =   225
         Index           =   7
         Left            =   150
         TabIndex        =   31
         Top             =   3195
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "Subdirectorio 7:"
         Height          =   225
         Index           =   6
         Left            =   150
         TabIndex        =   29
         Top             =   2850
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "Subdirectorio 6:"
         Height          =   225
         Index           =   5
         Left            =   150
         TabIndex        =   27
         Top             =   2505
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "Subdirectorio 5:"
         Height          =   225
         Index           =   4
         Left            =   150
         TabIndex        =   25
         Top             =   2160
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "Subdirectorio 4:"
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   23
         Top             =   1815
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "Subdirectorio 3:"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   21
         Top             =   1470
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "Subdirectorio 2:"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   19
         Top             =   1110
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "Subdirectorio 1:"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   17
         Top             =   765
         Width           =   1125
      End
      Begin VB.Label Label6 
         Caption         =   "Maestro:"
         Height          =   225
         Left            =   645
         TabIndex        =   15
         Top             =   405
         Width           =   675
      End
   End
   Begin VB.CommandButton CmdAccept 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   8520
      Width           =   1185
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3540
      TabIndex        =   11
      Top             =   8520
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de aplicación"
      Height          =   3990
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   4650
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   240
         Left            =   1875
         TabIndex        =   43
         ToolTipText     =   "examinar..."
         Top             =   3630
         Width           =   285
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   240
         Left            =   1890
         TabIndex        =   42
         ToolTipText     =   "examinar..."
         Top             =   3015
         Width           =   285
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   165
         MaxLength       =   100
         TabIndex        =   37
         Text            =   "My APP.EXE"
         Top             =   2985
         Width           =   2040
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   2430
         MaxLength       =   50
         TabIndex        =   36
         Text            =   "Runme NOW"
         Top             =   2985
         Width           =   2040
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   135
         MaxLength       =   100
         TabIndex        =   35
         Text            =   "My Readme"
         Top             =   3600
         Width           =   2040
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2430
         MaxLength       =   100
         TabIndex        =   34
         Text            =   "Readme NOW"
         Top             =   3600
         Width           =   2040
      End
      Begin VB.TextBox Text5 
         Height          =   795
         Left            =   150
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Text            =   "FrmINI.frx":0000
         Top             =   1830
         Width           =   4335
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2430
         MaxLength       =   100
         TabIndex        =   7
         Text            =   "My APP - Ver 1.0"
         Top             =   1260
         Width           =   2040
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   150
         MaxLength       =   100
         TabIndex        =   5
         Text            =   "My Company"
         Top             =   1260
         Width           =   2040
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2430
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "Ver 1.0"
         Top             =   645
         Width           =   2040
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   165
         MaxLength       =   100
         TabIndex        =   1
         Text            =   "My APP"
         Top             =   645
         Width           =   2040
      End
      Begin VB.Label Label11 
         Caption         =   "Archivo EXE principal:"
         Height          =   195
         Left            =   165
         TabIndex        =   41
         Top             =   2775
         Width           =   1740
      End
      Begin VB.Label Label10 
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   2430
         TabIndex        =   40
         Top             =   2775
         Width           =   1740
      End
      Begin VB.Label Label9 
         Caption         =   "Archivo Readme principal:"
         Height          =   195
         Left            =   150
         TabIndex        =   39
         Top             =   3390
         Width           =   1875
      End
      Begin VB.Label Label8 
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   2430
         TabIndex        =   38
         Top             =   3390
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Comentarios:"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   1635
         Width           =   1740
      End
      Begin VB.Label Label4 
         Caption         =   "Título de ventanas:"
         Height          =   195
         Left            =   2430
         TabIndex        =   8
         Top             =   1050
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre de la Compañía:"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   1050
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Versión:"
         Height          =   195
         Left            =   2430
         TabIndex        =   4
         Top             =   435
         Width           =   1740
      End
      Begin VB.Label Label1 
         Caption         =   "Título de la aplicación:"
         Height          =   195
         Left            =   165
         TabIndex        =   2
         Top             =   435
         Width           =   1740
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   105
      X2              =   4725
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   90
      X2              =   4725
      Y1              =   8415
      Y2              =   8415
   End
End
Attribute VB_Name = "FrmINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAccept_Click()

Dim i As Integer

FrmCab.Enabled = True

FrmCab.Label3.Caption = Text1.Text
FrmCab.Label4.Caption = Text2.Text
FrmCab.Label5.Caption = Text3.Text
FrmCab.Label6.Caption = Text4.Text
FrmCab.Label7.Caption = Text5.Text
FrmCab.Label8.Caption = Text6.Text

For i = 0 To 8
    FrmCab.Label9(i).Caption = Text7(i).Text
Next i

FrmCab.Label10.Caption = Text8.Text
FrmCab.Label11.Caption = Text9.Text
FrmCab.Label12.Caption = Text10.Text
FrmCab.Label13.Caption = Text11.Text

Unload Me

End Sub

Private Sub CmdCancel_Click()

FrmCab.Enabled = True
FrmCab.CHKini.value = 0

Unload Me

End Sub

Private Sub Command1_Click()

Dim ConvertTX As String
Dim Result As String

On Error Resume Next
Cmd1.InitDir = App.Path
Cmd1.Filter = "Archivos de programa (*.exe)|*.exe|Archivos de programa"
Cmd1.DialogTitle = "Seleccione archivo de programa"
Cmd1.CancelError = True
Cmd1.ShowOpen

If Err.Number = 32755 Then Exit Sub

ConvertTX = Cmd1.FileName
Text11.Text = ConvertTX

End Sub

Private Sub Command2_Click()

Dim ConvertTX As String
Dim Result As String

On Error Resume Next
Cmd1.InitDir = App.Path
Cmd1.Filter = "Archivos de texto (*.txt)|*.txt|Archivos de texto"
Cmd1.DialogTitle = "Seleccione archivo de texto"
Cmd1.CancelError = True
Cmd1.ShowOpen

If Err.Number = 32755 Then Exit Sub

ConvertTX = Cmd1.FileName
Text9.Text = ConvertTX

End Sub


Private Sub Form_Load()

Dim i As Integer

If FrmCab.Label3.Caption = "" Or FrmCab.Label3.Caption = " " Then
    FrmCab.Enabled = False
Else
    Text1.Text = FrmCab.Label3.Caption
    Text2.Text = FrmCab.Label4.Caption
    Text3.Text = FrmCab.Label5.Caption
    Text4.Text = FrmCab.Label6.Caption
    Text5.Text = FrmCab.Label7.Caption
    Text6.Text = FrmCab.Label8.Caption
    For i = 0 To 8
        Text7(i).Text = FrmCab.Label9(i).Caption
    Next i
    Text8.Text = FrmCab.Label10.Caption
    Text9.Text = FrmCab.Label11.Caption
    Text10.Text = FrmCab.Label12.Caption
    Text11.Text = FrmCab.Label13.Caption
    FrmCab.Enabled = False
End If

End Sub

