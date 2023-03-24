VERSION 5.00
Begin VB.Form FrmWarning 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ONLY development - Live Update"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   5115
      Begin VB.CommandButton Command1 
         Caption         =   "&Continuar >>"
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   1320
         Width           =   1125
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "<< C&ancelar"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmWarning.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   4935
      End
   End
End
Attribute VB_Name = "FrmWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()

MsgBox "Ha cancelado la operación de actualización.", vbInformation
Unload Me

End Sub

Private Sub Command1_Click()

FrmDecompress.Show
Unload Me

End Sub
