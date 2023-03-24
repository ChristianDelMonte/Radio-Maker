VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FRMState 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar a base de datos"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Command3"
      Height          =   405
      Left            =   4620
      TabIndex        =   8
      Top             =   3720
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command3"
      Height          =   405
      Left            =   3150
      TabIndex        =   7
      Top             =   3720
      Width           =   1365
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   405
      Left            =   180
      TabIndex        =   6
      Top             =   3720
      Width           =   1365
   End
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   1830
      Top             =   3660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   5745
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      MaxLength       =   255
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3120
      Width           =   5745
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2490
      Width           =   2865
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Estados / Provincias ya cargados:"
      Height          =   225
      Left            =   150
      TabIndex        =   5
      Top             =   210
      Width           =   3285
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "URL de coneccion climatica:"
      Height          =   225
      Left            =   150
      TabIndex        =   2
      Top             =   2880
      Width           =   2835
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de Estado / Provincia:"
      Height          =   225
      Left            =   150
      TabIndex        =   0
      Top             =   2220
      Width           =   2835
   End
End
Attribute VB_Name = "FRMState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Recargarlista()

Dim i As Integer, Data As String

List1.Clear

On Error Resume Next
LastReg = GetStateLastReg

    Text1.Text = GetComLng_ByID(LNGDef, "144")
    Text2.Text = GetComLng_ByID(LNGDef, "145")

For i = 1 To LastReg
    StateDatabase = GetStateData(i)
    Data = Trim(StateDatabase.State_Desc)
    List1.AddItem Data
Next i

End Sub

Private Sub Command1_Click()

'guardamos los datos
StateDatabase.State_Desc = Trim(Text1.Text)
StateDatabase.State_URL = Trim(Text2.Text)
FileState = SaveStateData(-1, StateDatabase)
If FileState = True Then
    'recargar lista
    Call Recargarlista
Else
    MsgBox "error"
End If

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Command3_Click()

On Error Resume Next
Cmd1.InitDir = App.Path
Cmd1.Filter = GetComLng_ByID(LNGDef, "147")
Cmd1.DialogTitle = GetComLng_ByID(LNGDef, "146")
Cmd1.CancelError = True
Cmd1.ShowOpen

If err.Number = 32755 Then Exit Sub

FileState = ImportTextFile(Cmd1.filename)

End Sub

Private Sub Form_Load()

Me.Caption = GetComLng_ByID(LNGDef, "137")
Label3.Caption = GetComLng_ByID(LNGDef, "138")
Label1.Caption = GetComLng_ByID(LNGDef, "139")
Label2.Caption = GetComLng_ByID(LNGDef, "140")
Command3.Caption = GetComLng_ByID(LNGDef, "141")
Command1.Caption = GetComLng_ByID(LNGDef, "142")
Command2.Caption = GetComLng_ByID(LNGDef, "143")

Call Recargarlista

End Sub

Private Sub List1_Click()

Text1.Text = List1.Text
Text2.Text = SearchURL_from_StateName(List1.Text)

End Sub

Private Sub Text1_GotFocus()

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub Text2_GotFocus()

Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)

End Sub
