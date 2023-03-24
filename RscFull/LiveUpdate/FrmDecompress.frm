VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDecompress 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ONLY Development - Live Update"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   4905
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1035
      End
      Begin VB.CommandButton CmdNow 
         Caption         =   "<< &Actualizar Ahora >>"
         Height          =   375
         Left            =   2940
         TabIndex        =   3
         Top             =   1320
         Width           =   1845
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmDecompress.frx":0000
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
         Width           =   4725
      End
   End
   Begin MSComctlLib.ProgressBar Prog1 
      Height          =   225
      Left            =   90
      TabIndex        =   0
      Top             =   1950
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
   End
End
Attribute VB_Name = "FrmDecompress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()

MsgBox "Ha cancelado la operación de actualización.", vbInformation
Unload Me

End Sub

Private Sub CmdNow_Click()

Dim Folder As String, Result As String
Dim TmpPath As String, FName As String

'///abrimos un archivo con el detalle de la version actual del programa
On Error GoTo er
Open App.Path & "\VerChk.dat" For Input As #31
Input #31, myVer, myPath
Close #31

'///comenzamos con los seteos antes de la extracción
TmpPath = App.Path & "\Temp"
FName = App.Path & "\Update\RM100_update.cbm"
Folder = myPath

'extract the selected file from cab to temp folder
Result = ExtractCabFile(FName, Trim(Folder), TmpPath, 0)

If Result = "NotOk" Then
    MsgBox "Ha ocurrido un Error. Operacion de actualización Abortada.", vbCritical
    Unload Me
Else
    MsgBox "Actualización Completada con éxito.", vbInformation
    Unload Me
End If
End
Exit Sub

er:
MsgBox "Ha ocurrido un Error. Operacion de actualización Abortada.", vbCritical
Unload Me

End Sub

