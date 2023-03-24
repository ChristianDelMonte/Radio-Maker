VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acerca CabMaker"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   3645
      TabIndex        =   0
      Top             =   2475
      Width           =   960
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "url: http://www.creaciones-digitales.com02.com"
      Height          =   195
      Left            =   135
      TabIndex        =   7
      Top             =   1620
      Width           =   3480
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "e-mail: creaciones-digitales@com02.com"
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   1395
      Width           =   2985
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed by: Christian A. Del Monte."
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   1170
      Width           =   4425
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2002 - ONLY development software inc."
      Height          =   240
      Left            =   90
      TabIndex        =   4
      Top             =   2070
      Width           =   4515
   End
   Begin VB.Line Line1 
      X1              =   4590
      X2              =   90
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "V 1.0a"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3510
      TabIndex        =   3
      Top             =   720
      Width           =   555
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CabMaker"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   810
      TabIndex        =   2
      Top             =   405
      Width           =   2715
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ONLY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   855
      TabIndex        =   1
      Top             =   180
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "About.frx":0000
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

FrmCab.Enabled = True
Unload Me

End Sub

Private Sub Form_Load()

FrmCab.Enabled = False

End Sub
