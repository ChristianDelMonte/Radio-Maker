VERSION 5.00
Begin VB.Form FrmDirs 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Opciones de directorio"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3105
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1095
      Width           =   2850
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1860
      TabIndex        =   4
      Top             =   1755
      Width           =   1185
   End
   Begin VB.CommandButton CmdAccept 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1755
      Width           =   1185
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   420
      Width           =   2850
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000015&
      X1              =   60
      X2              =   3045
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   60
      X2              =   3045
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SubDirectorios:"
      Height          =   210
      Left            =   150
      TabIndex        =   2
      Top             =   870
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione directorio:"
      Height          =   225
      Left            =   135
      TabIndex        =   1
      Top             =   195
      Width           =   1650
   End
End
Attribute VB_Name = "FrmDirs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAccept_Click()

Dim DataA(0 To 4) As String
Dim DataKa As String
Dim NIndex As Integer
Dim Response
Dim itmX As ListItem

NIndex = CInt(Trim(FrmCab.T1View.SelectedItem.Index))    'numero de index
DataKa = Trim(FrmCab.T1View.SelectedItem.Key)            'key

'seleccionamos el item
FrmCab.T1View.ListItems.Item(NIndex).Selected = True

'seteamos los nuevos datos
'HOST FILE
DataA(0) = Trim(FrmCab.T1View.SelectedItem.Text)    'file & path
DataA(1) = Trim(FrmCab.T1View.SelectedItem.ListSubItems(1).Text)    'path only
DataA(2) = Trim(FrmCab.T1View.SelectedItem.ListSubItems(2).Text)    'file only
DataA(3) = Trim(FrmCab.T1View.SelectedItem.ListSubItems(3).Text)    'Directorio Type
DataA(4) = Trim(FrmCab.T1View.SelectedItem.ListSubItems(4).Text)    'Sub Dir type

'chequeos necesarios
FrmCab.T1View.ListItems.Remove (NIndex)

'ponemos los nuevos datos
Set itmX = FrmCab.T1View.ListItems.Add(NIndex, DataKa, DataA(0)) 'path & file
itmX.SubItems(1) = DataA(1)
itmX.SubItems(2) = DataA(2)

Select Case Combo1.Text
    Case "APP directory"
        itmX.SubItems(3) = 1
        
    Case "Windows directory"
        itmX.SubItems(3) = 2
        
    Case "Windows\System directory"
        itmX.SubItems(3) = 3
        
    Case "APP SUB directory"
        itmX.SubItems(3) = 4
        
    Case Else
        itmX.SubItems(3) = 0
End Select

Select Case Combo2.ListIndex
    Case 0
        itmX.SubItems(4) = 1
    Case 1
        itmX.SubItems(4) = 2
    Case 2
        itmX.SubItems(4) = 3
    Case 3
        itmX.SubItems(4) = 4
    Case 4
        itmX.SubItems(4) = 5
    Case 5
        itmX.SubItems(4) = 6
    Case 6
        itmX.SubItems(4) = 7
    Case 7
        itmX.SubItems(4) = 8
    Case 8
        itmX.SubItems(4) = 9
    Case Else
        itmX.SubItems(4) = 0
End Select

'una vez finalizado. seleccionamos el item
FrmCab.T1View.ListItems.Item(NIndex).Selected = True
'y... actualizamos las horas de lanzamiento
FrmCab.Enabled = True
Unload Me

End Sub

Private Sub CmdCancel_Click()

FrmCab.Enabled = True
Unload Me

End Sub

Private Sub Combo1_Change()

Select Case Combo1.Text
    Case "APP SUB directory"
        Combo2.Enabled = True
        
    Case Else
        Combo2.Enabled = False
        
End Select

End Sub

Private Sub Combo1_Click()

Select Case Combo1.Text
    Case "APP SUB directory"
        Combo2.Enabled = True
        
    Case Else
        Combo2.Enabled = False
        
End Select

End Sub

Private Sub Form_Load()

'Public Const STP_None = 0
'Public Const STP_AppDir = 1
'Public Const STP_WinDir = 2
'Public Const STP_WinSysDir = 3
'Public Const STP_AppSubDir = 4

'get the index
'LIDX.Caption = FrmCab.T1View.SelectedItem.Index

Combo1.Text = "Ninguno"
Combo1.AddItem "Ninguno", 0
Combo1.AddItem "APP directory", 1
Combo1.AddItem "Windows directory", 2
Combo1.AddItem "Windows\System directory", 3
Combo1.AddItem "APP SUB directory", 4

Combo2.Enabled = False
If FrmCab.Label3.Caption = "" Or FrmCab.Label3.Caption = " " Then
    Combo2.Text = "No especificado!!"
    FrmCab.Enabled = False
Else
    Combo2.Text = FrmCab.Label9(0).Caption
    For i = 0 To 8
        Combo2.AddItem FrmCab.Label9(i).Caption, i
    Next i
    FrmCab.Enabled = False
End If

End Sub

