VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ConfigForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuraciones generales"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   345
      Left            =   2970
      TabIndex        =   27
      Top             =   1350
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   435
      Left            =   6900
      TabIndex        =   26
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   5700
      TabIndex        =   25
      Top             =   4260
      Width           =   1065
   End
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   4170
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Caption         =   "Seleccion de voces"
      Height          =   765
      Left            =   4170
      TabIndex        =   21
      Top             =   150
      Width           =   3825
      Begin VB.OptionButton Option5 
         Caption         =   "Personalizada"
         Height          =   225
         Left            =   2400
         TabIndex        =   24
         Top             =   360
         Width           =   1395
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Femenina"
         Height          =   225
         Left            =   1200
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Masculina"
         Height          =   225
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Archivos de audio personalizados"
      Enabled         =   0   'False
      Height          =   2895
      Left            =   4170
      TabIndex        =   12
      Top             =   1140
      Width           =   3825
      Begin VB.CommandButton Command7 
         Caption         =   "Command4"
         Height          =   285
         Left            =   3270
         TabIndex        =   31
         Top             =   2460
         Width           =   435
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command4"
         Height          =   285
         Left            =   3270
         TabIndex        =   30
         Top             =   1860
         Width           =   435
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command4"
         Height          =   285
         Left            =   3270
         TabIndex        =   29
         Top             =   1260
         Width           =   435
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   285
         Left            =   3270
         TabIndex        =   28
         Top             =   660
         Width           =   435
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   150
         MaxLength       =   255
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   2460
         Width           =   3075
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   150
         MaxLength       =   255
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   1860
         Width           =   3075
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   150
         MaxLength       =   255
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   1260
         Width           =   3075
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   150
         MaxLength       =   255
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   660
         Width           =   3075
      End
      Begin VB.Label Label9 
         Caption         =   "Humedad:"
         Height          =   255
         Left            =   150
         TabIndex        =   19
         Top             =   2220
         Width           =   1395
      End
      Begin VB.Label Label8 
         Caption         =   "Temperatura:"
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   1620
         Width           =   1395
      End
      Begin VB.Label Label7 
         Caption         =   "Minutos:"
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Label Label6 
         Caption         =   "Hora:"
         Height          =   255
         Left            =   150
         TabIndex        =   13
         Top             =   420
         Width           =   1395
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tiempo de actualizacion on-line"
      Height          =   1665
      Left            =   90
      TabIndex        =   7
      Top             =   2970
      Width           =   3855
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "60"
         Top             =   1230
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "segundos."
         Height          =   255
         Left            =   3060
         TabIndex        =   11
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Tiempo de refresco:"
         Height          =   285
         Left            =   150
         TabIndex        =   9
         Top             =   1260
         Width           =   1545
      End
      Begin VB.Label Label3 
         Caption         =   "Recuerde que debe poseer una coneccion a internet permanente para poder actualizar los datos climaticos."
         Height          =   585
         Left            =   150
         TabIndex        =   8
         Top             =   330
         Width           =   3525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Visualizacion de temperatura"
      Height          =   855
      Left            =   90
      TabIndex        =   4
      Top             =   1860
      Width           =   3855
      Begin VB.OptionButton Option2 
         Caption         =   "Grados Centigrados"
         Height          =   285
         Left            =   1950
         TabIndex        =   6
         Top             =   390
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Grados Farenhit"
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   390
         Width           =   1695
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1350
      Width           =   2805
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   390
      Width           =   3915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Estado / Provincia para la extraccion de datos climaticos"
      Height          =   435
      Left            =   90
      TabIndex        =   2
      Top             =   900
      Width           =   3825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione idioma"
      Height          =   225
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   3855
   End
End
Attribute VB_Name = "ConfigForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

FRMState.Show

End Sub

Private Sub Command2_Click()
    
'seteamos los datos a guardar en configuracion
Select Case Combo1.Text
    Case "ESPAÑOL": ConfigData.Lng_Id = 1
    Case "INGLES": ConfigData.Lng_Id = 2
    Case "FRANCES": ConfigData.Lng_Id = 3
    Case "ITALIANO": ConfigData.Lng_Id = 4
    Case "PORTUGUES": ConfigData.Lng_Id = 5
    Case Else: ConfigData.Lng_Id = 1
End Select

ConfigData.State_Id = Trim(Combo2.Text)

If Option1.value = True Then
    ConfigData.Temp_Mode = 1
Else
    ConfigData.Temp_Mode = 2
End If

If Option3.value = True Then
    ConfigData.Voice_Id = 1
Else
    If Option4.value = True Then
        ConfigData.Voice_Id = 2
    Else
        ConfigData.Voice_Id = 3
    End If
End If

ConfigData.Temp_RTime = CInt(Trim(Text1.Text))
ConfigData.HVoicePath = Trim(Text2.Text)
ConfigData.MVoicePath = Trim(Text3.Text)
ConfigData.TmpVoicePath = Trim(Text4.Text)
ConfigData.HumVoicePath = Trim(Text5.Text)

FileState = SaveConfigData(ConfigData)

If Command2.Enabled = True Then
    Command2.Enabled = False
End If

End Sub

Private Sub Command3_Click()

Call Command2_Click
Call MainForm.RecargarConfig
Unload Me

End Sub

Private Sub Command4_Click()

On Error Resume Next
Cmd1.InitDir = App.Path
Cmd1.Filter = GetComLng_ByID(LNGDef, "148")
Cmd1.DialogTitle = GetComLng_ByID(LNGDef, "149")
Cmd1.CancelError = True
Cmd1.ShowOpen

If err.Number = 32755 Then Exit Sub

Text2.Text = Cmd1.filename

End Sub

Private Sub Command5_Click()

On Error Resume Next
Cmd1.InitDir = App.Path
Cmd1.Filter = GetComLng_ByID(LNGDef, "148")
Cmd1.DialogTitle = GetComLng_ByID(LNGDef, "149")
Cmd1.CancelError = True
Cmd1.ShowOpen

If err.Number = 32755 Then Exit Sub

Text3.Text = Cmd1.filename

End Sub

Private Sub Command6_Click()

On Error Resume Next
Cmd1.InitDir = App.Path
Cmd1.Filter = GetComLng_ByID(LNGDef, "148")
Cmd1.DialogTitle = GetComLng_ByID(LNGDef, "149")
Cmd1.CancelError = True
Cmd1.ShowOpen

If err.Number = 32755 Then Exit Sub

Text4.Text = Cmd1.filename

End Sub

Private Sub Command7_Click()

On Error Resume Next
Cmd1.InitDir = App.Path
Cmd1.Filter = GetComLng_ByID(LNGDef, "148")
Cmd1.DialogTitle = GetComLng_ByID(LNGDef, "149")
Cmd1.CancelError = True
Cmd1.ShowOpen

If err.Number = 32755 Then Exit Sub

Text5.Text = Cmd1.filename

End Sub

Private Sub Form_Load()

Dim i As Integer, NewData As String

ConfigData = GetConfigData

Me.Caption = GetComLng_ByID(LNGDef, "1166")
Label1.Caption = GetComLng_ByID(LNGDef, "116")
Label2.Caption = GetComLng_ByID(LNGDef, "117")
Frame1.Caption = GetComLng_ByID(LNGDef, "118")
Option1.Caption = GetComLng_ByID(LNGDef, "119")
Option2.Caption = GetComLng_ByID(LNGDef, "120")
Frame2.Caption = GetComLng_ByID(LNGDef, "121")
Label3.Caption = GetComLng_ByID(LNGDef, "122")
Label4.Caption = GetComLng_ByID(LNGDef, "123")
Label5.Caption = GetComLng_ByID(LNGDef, "124")
Frame4.Caption = GetComLng_ByID(LNGDef, "125")
Option3.Caption = GetComLng_ByID(LNGDef, "126")
Option4.Caption = GetComLng_ByID(LNGDef, "127")
Option5.Caption = GetComLng_ByID(LNGDef, "128")
Frame3.Caption = GetComLng_ByID(LNGDef, "129")
Label6.Caption = GetComLng_ByID(LNGDef, "130")
Label7.Caption = GetComLng_ByID(LNGDef, "131")
Label8.Caption = GetComLng_ByID(LNGDef, "132")
Label8.Caption = GetComLng_ByID(LNGDef, "133")
Command4.Caption = GetComLng_ByID(LNGDef, "1331")
Command5.Caption = GetComLng_ByID(LNGDef, "1331")
Command6.Caption = GetComLng_ByID(LNGDef, "1331")
Command7.Caption = GetComLng_ByID(LNGDef, "1331")
Command2.Caption = GetComLng_ByID(LNGDef, "134")
Command3.Caption = GetComLng_ByID(LNGDef, "135")
Command1.Caption = GetComLng_ByID(LNGDef, "136")

Select Case ConfigData.Lng_Id
    Case 1: Combo1.Text = "ESPAÑOL"
    Case 2: Combo1.Text = "INGLES"
    Case 3: Combo1.Text = "FRANCES"
    Case 4: Combo1.Text = "ITALIANO"
    Case 5: Combo1.Text = "PORTUGUES"
    Case Else: Combo1.Text = "ESPAÑOL"
End Select

If GetLNGFilename("ESPAÑOL") <> "." Then Combo1.AddItem "ESPAÑOL"
If GetLNGFilename("INGLES") <> "." Then Combo1.AddItem "INGLES"
If GetLNGFilename("FRANCES") <> "." Then Combo1.AddItem "FRANCES"
If GetLNGFilename("ITALIANO") <> "." Then Combo1.AddItem "ITALIANO"
If GetLNGFilename("PORTUGUES") <> "." Then Combo1.AddItem "PORTUGUES"

Combo2.Text = Trim(ConfigData.State_Id)

'datos del statedatabase
LastReg = GetStateLastReg
For i = 1 To LastReg
    NewData = Trim(GetStateData(i).State_Desc)
    Combo2.AddItem NewData
Next i
    
If ConfigData.Temp_Mode = 1 Then
    Option1.value = True
Else
    Option2.value = True
End If

If ConfigData.Voice_Id = 1 Then
    Option3.value = True
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Frame3.Enabled = False
    If ConfigData.Lng_Id = 1 Then   'VOZ MASCULINA EN ESPAÑOL
        Text2.Text = App.Path & "\Data\" & FnameHoraESP
        Text3.Text = App.Path & "\Data\" & FnameMinutoESP
        Text4.Text = App.Path & "\Data\" & FnameTempESP
        Text5.Text = App.Path & "\Data\" & FnameHumeESP
    Else    'VOZ MASCULINA EN INGLES
        Text2.Text = App.Path & "\Data\" & FnameHoraEN
        Text3.Text = App.Path & "\Data\" & FnameMinutoEN
        Text4.Text = App.Path & "\Data\" & FnameTempEN
        Text5.Text = App.Path & "\Data\" & FnameHumeEN
    End If
Else
    If ConfigData.Voice_Id = 2 Then
        Option4.value = True
        Command4.Enabled = False
        Command5.Enabled = False
        Command6.Enabled = False
        Command7.Enabled = False
        Frame3.Enabled = False
        If ConfigData.Lng_Id = 1 Then   'VOZ FEMENINA EN ESPAÑOL
            Text2.Text = "none"
            Text3.Text = "none"
            Text4.Text = "none"
            Text5.Text = "none"
        Else    'VOZ FEMENINA EN INGLES
            Text2.Text = "none"
            Text3.Text = "none"
            Text4.Text = "none"
            Text5.Text = "none"
        End If
    Else
        Option5.value = True
        Command4.Enabled = True
        Command5.Enabled = True
        Command6.Enabled = True
        Command7.Enabled = True
        Frame3.Enabled = True
        'VOZ PERSONALIZADA
        Text2.Text = Trim(ConfigData.HVoicePath)
        Text3.Text = Trim(ConfigData.MVoicePath)
        Text4.Text = Trim(ConfigData.TmpVoicePath)
        Text5.Text = Trim(ConfigData.HumVoicePath)
    End If
End If

Text1.Text = Trim(ConfigData.Temp_RTime)

End Sub

Private Sub Option3_Click()

Frame3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False

End Sub

Private Sub Option4_Click()

Frame3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False

End Sub

Private Sub Option5_Click()

Frame3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True

End Sub
