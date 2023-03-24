VERSION 5.00
Begin VB.Form NewLng 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RMLocalize - Setear nuevo lenguaje"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3090
      TabIndex        =   3
      Top             =   1680
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   210
      TabIndex        =   2
      Top             =   1680
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   210
      MaxLength       =   50
      TabIndex        =   0
      Top             =   990
      Width           =   4155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lenguaje:"
      Height          =   255
      Left            =   210
      TabIndex        =   1
      Top             =   660
      Width           =   1395
   End
End
Attribute VB_Name = "NewLng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim ultreg As Integer, tst As Boolean

ultreg = GetCFGLastReg
ultreg = ultreg + 1

ConfigData.Id = ultreg
ConfigData.Lenguaje = Trim(UCase(Text1.Text))
ConfigData.LNG_Predet = 0

tst = SaveCFGData(ConfigData, ultreg, False)
If tst = True Then
    Unload Me
Else
    MsgBox "Upss!! error. Plaese... Report This!!!"
    Unload Me
End If

End Sub

Private Sub Command2_Click()

Unload Me

End Sub
