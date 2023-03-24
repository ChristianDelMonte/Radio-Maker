VERSION 5.00
Begin VB.Form DlgAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acerca RMMC"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3615
   ForeColor       =   &H00C0C0C0&
   Icon            =   "DlgAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   120
      Picture         =   "DlgAbout.frx":030A
      ScaleHeight     =   600
      ScaleWidth      =   2130
      TabIndex        =   1
      Top             =   120
      Width           =   2130
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Christian Adrian Del Monte Inc."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3480
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2001 Creaciones Digitales Inc."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Versión 1.0a"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Multimedia Control"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
End
Attribute VB_Name = "DlgAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
