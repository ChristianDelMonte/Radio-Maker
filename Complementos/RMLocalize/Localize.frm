VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form LocalizeForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RMLocalize - Principal"
   ClientHeight    =   7215
   ClientLeft      =   165
   ClientTop       =   315
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "&Nuevo..."
      Height          =   315
      Left            =   9120
      TabIndex        =   23
      Top             =   1500
      Width           =   915
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   1500
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Abrir archivo..."
      Height          =   465
      Left            =   120
      TabIndex        =   20
      Top             =   6600
      Width           =   1635
   End
   Begin VB.CommandButton Command6 
      Caption         =   "N&uevo"
      Height          =   315
      Left            =   6810
      TabIndex        =   19
      Top             =   3150
      Width           =   975
   End
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   3360
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Nuevo..."
      Height          =   315
      Left            =   8310
      TabIndex        =   17
      Top             =   2340
      Width           =   1005
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Salir"
      Height          =   465
      Left            =   8850
      TabIndex        =   14
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Guardar cambios"
      Height          =   465
      Left            =   5520
      TabIndex        =   13
      Top             =   6600
      Width           =   1665
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Importar..."
      Height          =   465
      Left            =   1890
      TabIndex        =   12
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   945
      Left            =   5520
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Localize.frx":0000
      Top             =   5280
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Height          =   825
      Left            =   5520
      MaxLength       =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "Localize.frx":0006
      Top             =   3960
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   5520
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3150
      Width           =   1185
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5490
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   2340
      Width           =   2715
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1500
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      Caption         =   "Listado de componentes"
      Height          =   3945
      Left            =   120
      TabIndex        =   1
      Top             =   1290
      Width           =   5145
      Begin VB.ListBox List1 
         Height          =   3375
         ItemData        =   "Localize.frx":000C
         Left            =   150
         List            =   "Localize.frx":000E
         TabIndex        =   18
         Top             =   390
         Width           =   4845
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lenguaje Predeterminado"
      Height          =   825
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   5145
      Begin VB.CommandButton Command1 
         Caption         =   "Establecer predeterminado"
         Height          =   345
         Left            =   2610
         TabIndex        =   16
         Top             =   360
         Width           =   2385
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   150
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   360
         Width           =   2325
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de archivo:"
      Height          =   255
      Left            =   6840
      TabIndex        =   22
      Top             =   1260
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Comentarios:"
      Height          =   255
      Left            =   5520
      TabIndex        =   11
      Top             =   5040
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Contenido:"
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   3720
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ID. de Sistema:"
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   2910
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lenguaje:"
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   2100
      Width           =   2685
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ID. General:"
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   1260
      Width           =   1185
   End
End
Attribute VB_Name = "LocalizeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SetLenguaje()

Dim numreg As Integer

numreg = GetIDFromCFGConfig(Combo2.Text)

Me.Caption = GetLNGData(numreg, 102)
Frame2.Caption = GetLNGData(numreg, 103)
Frame1.Caption = GetLNGData(numreg, 104)
Label2.Caption = GetLNGData(numreg, 105)
Label1.Caption = GetLNGData(numreg, 106)
Label3.Caption = GetLNGData(numreg, 107)
Label4.Caption = GetLNGData(numreg, 108)
Label5.Caption = GetLNGData(numreg, 109)
Label6.Caption = GetLNGData(numreg, 110)
Command6.Caption = GetLNGData(numreg, 111)
Command5.Caption = GetLNGData(numreg, 112)
Command8.Caption = GetLNGData(numreg, 113)
Command1.Caption = GetLNGData(numreg, 114)
Command7.Caption = GetLNGData(numreg, 115)
Command2.Caption = GetLNGData(numreg, 116)
Command3.Caption = GetLNGData(numreg, 117)
Command4.Caption = GetLNGData(numreg, 118)

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

End Sub

Private Sub Command1_Click()

'establecemos un idioma como predeterminado y seteamos los otros como no predeterminados (no puede haber dos iguales)
Dim numreg As Integer, ultreg As Integer, x As Integer, z As Integer, tst As Boolean, NewItem As String

numreg = GetIDFromCFGConfig(Combo2.Text)
ultreg = GetCFGLastReg

For x = 1 To ultreg
    ConfigData = GetCFGConfigList(x)
    If Trim(ConfigData.Lenguaje) = Trim(Combo2.Text) Then
        ConfigData.LNG_Predet = 1
        tst = SaveCFGData(ConfigData, x, True)
    Else
        tst = SaveCFGData(ConfigData, x, False)
    End If
    If x >= ultreg Then
        Exit For
    End If
Next

'ahora volvemos a cargar los datos
List1.Clear

numreg = GetIDFromCFGConfig(Combo2.Text)

'rellenamos la lista con los item ya nacionalizados segun el idioma por defecto
'primero verificamos el lenguaje predeterminado
If Combo2.Text = GetLNGData(numreg, 119) Then
    'MsgBox "Antes de importar datos de una lista debe establecer un idioma predeterminado. Los datos importados seran tomados por defecto para dicho idioma."
    Exit Sub
Else
    ultreg = GetLNGLastReg
    For z = 1 To ultreg
        LenguajeData = GetLNGDataList(numreg, z)
        If Trim(LenguajeData.PRG_ID) <> 0 Then
            NewItem = Trim(LenguajeData.PRG_ID) & " - " & Trim(LenguajeData.PRG_Desc)
            List1.AddItem NewItem
        End If
    Next z
End If

Call SetLenguaje

End Sub

Private Sub Command2_Click()

Dim numreg As Integer

numreg = GetIDFromCFGConfig(Combo2.Text)

If Combo2.Text = GetLNGData(numreg, 119) Then
    MsgBox = GetLNGData(numreg, 120)
    Exit Sub
End If

Cmd1.ShowOpen
If Cmd1.FileName <> "" Then
    ImportTextFile Cmd1.FileName, numreg
End If

End Sub

Private Sub Command3_Click()

Dim ultreg As Integer

'extraemos el id del lenguaje en cuestion
numreg = GetIDFromCFGConfig(Combo1.Text)

End Sub

Private Sub Command4_Click()

End

End Sub

Private Sub Command5_Click()

NewLng.Show , Me

End Sub

Private Sub Command6_Click()

Dim ultreg As Integer
ultreg = GetLNGLastReg
ultreg = ultreg + 1
Text1.Text = ultreg

Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

End Sub

Private Sub Form_Load()

Dim ultreg As Integer, z As Integer
Dim numreg As Integer, NewItem As String

'cargamos los datos de idiomas delarchivo de configuracion
ultreg = GetCFGLastReg

numreg = GetIDFromCFGConfig(Combo2.Text)

Combo1.Text = GetLNGData(numreg, 119)
Combo2.Text = GetLNGData(numreg, 119)

For z = 1 To ultreg
    ConfigData = GetCFGConfigList(z)
    Combo1.AddItem Trim(ConfigData.Lenguaje)
    Combo2.AddItem Trim(ConfigData.Lenguaje)
    If ConfigData.LNG_Predet = 1 Then
        Combo1.Text = Trim(ConfigData.Lenguaje)
        Combo2.Text = Trim(ConfigData.Lenguaje)
    End If
Next z

'rellenamos la lista con los item ya nacionalizados segun el idioma por defecto
'primero verificamos el lenguaje predeterminado
If Combo2.Text = GetLNGData(numreg, 119) Then
    'MsgBox "Antes de importar datos de una lista debe establecer un idioma predeterminado. Los datos importados seran tomados por defecto para dicho idioma."
    Exit Sub
Else
    ultreg = GetLNGLastReg
    For z = 1 To ultreg
        LenguajeData = GetLNGDataList(numreg, z)
        If Trim(LenguajeData.PRG_ID) <> 0 Then
            NewItem = Trim(LenguajeData.PRG_ID) & " - " & Trim(LenguajeData.PRG_Desc)
            List1.AddItem NewItem
        End If
    Next z
End If

Call SetLenguaje
Text5.Text = LNG_File

End Sub

Private Sub List1_Click()

Dim numreg As Integer, ultreg As Integer, progid As Integer, oldid As String

numreg = GetIDFromCFGConfig(Combo2.Text)

If Combo2.Text = GetLNGData(numreg, 119) Then
    'MsgBox "Antes de importar datos de una lista debe establecer un idioma predeterminado. Los datos importados seran tomados por defecto para dicho idioma."
    Exit Sub
Else
    ultreg = GetLNGLastReg
    oldid = Left$(List1.Text, 4)
    progid = CInt(Trim(oldid))
    LenguajeData = GetLNGData2(numreg, progid)
    If LenguajeData.Id = -1 Then
        MsgBox "oups"
    Else
        Text1.Text = LenguajeData.Id
        Combo1.Text = Combo2.Text
        Text2.Text = Trim(LenguajeData.PRG_ID)
        Text3.Text = Trim(LenguajeData.PRG_Desc)
        Text4.Text = Trim(LenguajeData.LNG_Comm)
    End If
End If

End Sub
