VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Begin VB.Form MainForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RMVoice - Ventana de estado"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   330
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Command3"
      Height          =   405
      Left            =   6330
      TabIndex        =   32
      Top             =   4710
      Width           =   1365
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command3"
      Height          =   405
      Left            =   4890
      TabIndex        =   31
      Top             =   4710
      Width           =   1365
   End
   Begin VB.Timer HuTimer 
      Left            =   4860
      Top             =   5760
   End
   Begin VB.Timer TTimer 
      Left            =   4440
      Top             =   5760
   End
   Begin VB.Timer ClimaTimer 
      Left            =   5880
      Top             =   5760
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   5520
      TabIndex        =   24
      Top             =   4260
      Width           =   2175
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6540
      Top             =   5610
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   5400
      Top             =   5760
   End
   Begin VB.Frame Frame5 
      Caption         =   "Configuraciones generales"
      Height          =   2295
      Left            =   2880
      TabIndex        =   10
      Top             =   1320
      Width           =   4815
      Begin VB.Label LblClima 
         Caption         =   "Español"
         Height          =   255
         Left            =   2040
         TabIndex        =   28
         Top             =   1740
         Width           =   2655
      End
      Begin VB.Label Label9 
         Caption         =   "Datos Climaticos para:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1740
         Width           =   1815
      End
      Begin VB.Label LblVoice 
         Caption         =   "Español"
         Height          =   255
         Left            =   2040
         TabIndex        =   26
         Top             =   1980
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo de voz:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1980
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Idioma Actual:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label LblLng 
         Caption         =   "Español"
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   1500
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "Archivo de Humedad:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label LBLHuName 
         Caption         =   "no especificado"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Archivo de Temperatura:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label LBLTName 
         Caption         =   "no especificado"
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label LBLHName 
         Caption         =   "ESPAÑOL"
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Archivo de minutos:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label LBLMName 
         Caption         =   "ESPAÑOL"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Archivo de hora:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Humedad"
      Height          =   1095
      Left            =   5400
      TabIndex        =   7
      Top             =   120
      Width           =   2295
      Begin VB.Label LblB 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "actualizando..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   30
         Top             =   870
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "199%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Temperatura"
      Height          =   1095
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   1935
      Begin VB.Label LblA 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "actualizando..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   29
         Top             =   870
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "-20º"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandos "
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   2655
      Begin VB.ListBox LSTCommand 
         BackColor       =   &H00000000&
         ForeColor       =   &H80000018&
         Height          =   1620
         ItemData        =   "FrmMain_Voice.frx":0000
         Left            =   120
         List            =   "FrmMain_Voice.frx":0002
         TabIndex        =   21
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hora Actual"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3135
      Begin VB.Label LBLTime 
         Alignment       =   2  'Center
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Timer MTimer 
      Left            =   3900
      Top             =   5760
   End
   Begin VB.Timer HTimer 
      Left            =   3480
      Top             =   5760
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PB"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   3810
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ONLY RMVoice Plug-In for ONLY Radiomaker. "
      Height          =   1275
      Left            =   120
      TabIndex        =   23
      Top             =   3810
      Width           =   4515
   End
   Begin VB.Label LblStrm 
      Caption         =   "1"
      Height          =   255
      Left            =   1440
      TabIndex        =   22
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label LblEnd 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label LblStart 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Data As String, RPlug As String, NewLng As String, cond As String
Dim lnga As Long, lngB As Long, lngC As Long, lngD As Long, temp1 As Long
Dim NewVol As Long
Dim Result As String
Dim EndPos As Long, Poscnv As Long, Rst As Long

Function RecargarConfig()

Dim Nuevaduracion As Long
On Error Resume Next

'cargamos los datos de la configuracion
ConfigData = GetConfigData
If ConfigData.Id = -1 Then  'no hay config file. CANNOT CONTINUE
    MsgBox "ERROR en archivo de configuración. No se puede continuar.", vbCritical
    MsgBox "ERROR in config file. Cannot continue.", vbCritical
    Unload MainForm
    Exit Function
End If

LBLTime.Caption = time$

'comenzamos seteando el lenguaje del sistema
Select Case ConfigData.Lng_Id
    Case 1: LNGDef = "ESPAÑOL"
    Case 2: LNGDef = "INGLES"
    Case 3: LNGDef = "FRANCES"
    Case 4: LNGDef = "ITALIANO"
    Case 5: LNGDef = "PORTUGUES"
    Case Else: LNGDef = "ESPAÑOL"
End Select

    Me.Caption = GetComLng_ByID(LNGDef, "100")
    
    Frame1.Caption = GetComLng_ByID(LNGDef, "101")
    Frame2.Caption = GetComLng_ByID(LNGDef, "104")
    Frame3.Caption = GetComLng_ByID(LNGDef, "102")
    Frame4.Caption = GetComLng_ByID(LNGDef, "103")
    Frame5.Caption = GetComLng_ByID(LNGDef, "105")
    
    Label3.Caption = GetComLng_ByID(LNGDef, "106")
    Label5.Caption = GetComLng_ByID(LNGDef, "107")
    Label8.Caption = GetComLng_ByID(LNGDef, "108")
    Label10.Caption = GetComLng_ByID(LNGDef, "109")
    Label12.Caption = GetComLng_ByID(LNGDef, "110")
    Label9.Caption = GetComLng_ByID(LNGDef, "111")
    Label6.Caption = GetComLng_ByID(LNGDef, "112")
    Label4.Caption = GetComLng_ByID(LNGDef, "115")
    
    Command4.Caption = GetComLng_ByID(LNGDef, "113") 'boton config
    Command2.Caption = GetComLng_ByID(LNGDef, "114") 'boton cerrar
    Command3.Caption = GetComLng_ByID(LNGDef, "102") 'boton actualizar datos climaticos 'no visible
    Command1.Caption = GetComLng_ByID(LNGDef, "150") 'boton prueba de audio 'no visible
    
    Label1.Caption = "--"
    Label2.Caption = "--"
    
    LblA.Caption = GetComLng_ByID(LNGDef, "1144")   'refresh label in red
    LblB.Caption = GetComLng_ByID(LNGDef, "1144")

If ConfigData.Lng_Id = 1 Then   'español
    Select Case ConfigData.Voice_Id
        Case 1 'masculina en español
            LBLHName.Caption = FnameHoraESP
            LBLMName.Caption = FnameMinutoESP
            LBLTName.Caption = FnameTempESP
            LBLHuName.Caption = FnameHumeESP
            LblLng.Caption = LNGDef
            LblVoice.Caption = GetComLng_ByID(LNGDef, "126")
        Case 2 'femenina en español
            LBLHName.Caption = "..." & Right$(Trim(ConfigData.HVoicePath), 25)
            LBLMName.Caption = "..." & Right$(Trim(ConfigData.MVoicePath), 25)
            LBLTName.Caption = "..." & Right$(Trim(ConfigData.TmpVoicePath), 25)
            LBLHuName.Caption = "..." & Right$(Trim(ConfigData.HumVoicePath), 25)
            LblLng.Caption = LNGDef
            LblVoice.Caption = GetComLng_ByID(LNGDef, "127")
        Case 3 'voz personalizada
            LBLHName.Caption = "..." & Right$(Trim(ConfigData.HVoicePath), 25)
            LBLMName.Caption = "..." & Right$(Trim(ConfigData.MVoicePath), 25)
            LBLTName.Caption = "..." & Right$(Trim(ConfigData.TmpVoicePath), 25)
            LBLHuName.Caption = "..." & Right$(Trim(ConfigData.HumVoicePath), 25)
            LblLng.Caption = LNGDef
            LblVoice.Caption = GetComLng_ByID(LNGDef, "128")
    End Select
Else
    If ConfigData.Lng_Id = 2 Then ' ingles
        Select Case ConfigData.Voice_Id
            Case 1 'masculina en ingles
                LBLHName.Caption = FnameHoraEN
                LBLMName.Caption = FnameMinutoEN
                LBLTName.Caption = FnameTempEN
                LBLHuName.Caption = FnameHumeEN
                LblLng.Caption = LNGDef
                LblVoice.Caption = GetComLng_ByID(LNGDef, "126")
            Case 2 'femenina en ingles
                LBLHName.Caption = "..." & Right$(Trim(ConfigData.HVoicePath), 25)
                LBLMName.Caption = "..." & Right$(Trim(ConfigData.MVoicePath), 25)
                LBLTName.Caption = "..." & Right$(Trim(ConfigData.TmpVoicePath), 25)
                LBLHuName.Caption = "..." & Right$(Trim(ConfigData.HumVoicePath), 25)
                LblLng.Caption = LNGDef
                LblVoice.Caption = GetComLng_ByID(LNGDef, "127")
            Case 3 'voz personalizada
                LBLHName.Caption = "..." & Right$(Trim(ConfigData.HVoicePath), 25)
                LBLMName.Caption = "..." & Right$(Trim(ConfigData.MVoicePath), 25)
                LBLTName.Caption = "..." & Right$(Trim(ConfigData.TmpVoicePath), 25)
                LBLHuName.Caption = "..." & Right$(Trim(ConfigData.HumVoicePath), 25)
                LblLng.Caption = LNGDef
                LblVoice.Caption = GetComLng_ByID(LNGDef, "128")
        End Select
    Else    'otro idioma no ingles ni español
        Select Case ConfigData.Voice_Id
            Case 1 'masculina en ingles
                LBLHName.Caption = FnameHoraEN
                LBLMName.Caption = FnameMinutoEN
                LBLTName.Caption = FnameTempEN
                LBLHuName.Caption = FnameHumeEN
                LblLng.Caption = LNGDef
                LblVoice.Caption = GetComLng_ByID(LNGDef, "126")
            Case 2 'femenina en ingles
                LBLHName.Caption = "..." & Right$(Trim(ConfigData.HVoicePath), 25)
                LBLMName.Caption = "..." & Right$(Trim(ConfigData.MVoicePath), 25)
                LBLTName.Caption = "..." & Right$(Trim(ConfigData.TmpVoicePath), 25)
                LBLHuName.Caption = "..." & Right$(Trim(ConfigData.HumVoicePath), 25)
                LblLng.Caption = LNGDef
                LblVoice.Caption = GetComLng_ByID(LNGDef, "127")
            Case 3 'voz personalizada
                LBLHName.Caption = "..." & Right$(Trim(ConfigData.HVoicePath), 25)
                LBLMName.Caption = "..." & Right$(Trim(ConfigData.MVoicePath), 25)
                LBLTName.Caption = "..." & Right$(Trim(ConfigData.TmpVoicePath), 25)
                LBLHuName.Caption = "..." & Right$(Trim(ConfigData.HumVoicePath), 25)
                LblLng.Caption = LNGDef
                LblVoice.Caption = GetComLng_ByID(LNGDef, "128")
        End Select
    End If
End If

LblClima.Caption = Trim(ConfigData.State_Id)

Call ActualizarClima
DoEvents

'iniciamos el timer de extraccion climatica
ClimaTimer.Enabled = True
Nuevaduracion = CLng(ConfigData.Temp_RTime)
Nuevaduracion = Nuevaduracion * 1000

ClimaTimer.Interval = Nuevaduracion

'inicializamos el dispositivo
Result = InitDevice(1)
Debug.Print "init device:" & Result

Timer1.Enabled = True
Timer1.Interval = 1000

End Function

Private Function ActualizarClima()

Dim WLugar As String, Wurl As String

LblA.Visible = True
LblB.Visible = True

'funcion para actualizar la informacion climatica
ConfigData = GetConfigData
WLugar = Trim(ConfigData.State_Id)  'extraemos el lugar de origen
Wurl = Trim(SearchURL_from_StateName(WLugar))

On Error GoTo er
Data = Inet1.OpenURL(Wurl)
Parse (Data)
LblA.Visible = False
LblB.Visible = False

Exit Function

er:
Label1.Caption = "N/A"
Label2.Caption = "N/A"
LblA.Visible = False
LblB.Visible = False

DisplayMsg "ERROR in module ActualizarClima -> Form: MainForm. ???? Este error suele deberse generalmente a problemas de coneccion a la base de datos climatica on-line. Intente mas tarde.", " Function_data: Site:" & WLugar & " URL:" & Wurl, err.Number, False
End Function

Private Function Parse(Data As String)
   
    lngD = InStr(1, Data, "&humid=") + 7 'this is basically looking For the start of &humid= in the data file, as u may guess it will be the humidity
    lngC = InStr(1, Data, "&cond=") + 6
    lnga = InStr(1, Data, "temp=") 'looking For the temp...
    'cond = Mid(Data, lngC, 5) 'This gets the "cloudy,"rainy" text etc; I wanted To change it so...
    
    'If cond = "cloud" Then 'if it says "cloud" replace it With the world "Cloudy :-("
    '    cond = "nublado"
    'End If
    'If cond = "clear" Then 'if it says "clear" replace it With the word "Sunny :-)"
    '    cond = "despejado"
    'End If

    temp1 = CLng(Mid(Data, lnga + 5, 2))
    ConfigData = GetConfigData
    If ConfigData.Temp_Mode = 1 Then
        temp1 = temp1
    Else
        temp1 = (temp1 - 32) * 5 / 9
    End If
    Label1.Caption = temp1 & "º"""
    Label2.Caption = Mid(Data, lngD, 2) & "%"

'C = (f - 32) * 5 / 9 conversion de farenhit a centigrados

End Function

Private Sub ClimaTimer_Timer()

Call ActualizarClima

End Sub

Private Sub Command1_Click()

If Stream01IsPlaying = BASSTRUE Then
    'LblVol.Caption = Str$(Stream01GetVolumen)
    LblStrm.Caption = "strm1"
    NewVol = 50
    Stream01SetVolumen (NewVol)
Else
    'LblVol.Caption = Str$(Stream02GetVolumen)
    LblStrm.Caption = "strm2"
    NewVol = 50
    Stream02SetVolumen (NewVol)
End If

Call InitHora

End Sub

Private Sub Command2_Click()

CloseDevice "Stream", "Stream"
Unload MainForm

End Sub

Private Sub Command3_Click()

'Call ActualizarClima
'Label1.Caption = "-14°"

Call InitTemperatura

End Sub

Private Sub Command4_Click()

ConfigForm.Show , Me

End Sub

Private Sub Form_Load()

Call RecargarConfig
DoEvents

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

CloseDevice "Stream", "Stream"

End Sub

Private Sub Form_Terminate()

CloseDevice "Stream", "Stream"

End Sub

Private Sub Form_Unload(Cancel As Integer)

CloseDevice "Stream", "Stream"

End Sub

Private Sub HTimer_Timer()

If MainForm.LblStrm.Caption = "1" Then
    Result = Stream01GetPosition(2)
Else
    Result = Stream02GetPosition(2)
End If

'EndPos = CLng(Trim(LblEnd.Caption))
Poscnv = CLng(Trim(Result))

'avisamos que el plugin esta en uso
Pstate = True

'verificamos si la reproduccion de la seccion debe finalizar
If Poscnv >= EndPos Then
    If MainForm.LblStrm.Caption = "1" Then
        Stream01Stop
    Else
        Stream02Stop
    End If
    HTimer.Interval = 0
    HTimer.Enabled = False
    'LSTCommand.AddItem "reproducir minutos..."
    Call InitMinutos

End If

End Sub

Private Sub HuTimer_Timer()

If MainForm.LblStrm.Caption = "1" Then
    Result = Stream01GetPosition(2)
Else
    Result = Stream02GetPosition(2)
End If

EndPos = CLng(Trim(LblEnd.Caption))
Poscnv = CLng(Trim(Result))

'avisamos que el plugin esta en uso
Pstate = True

'verificamos si la reproduccion de la posicion debe finalizar
If Poscnv >= EndPos Then
    If MainForm.LblStrm.Caption = "1" Then
        Stream01Stop
    Else
        Stream02Stop
    End If
    HuTimer.Interval = 0
    HuTimer.Enabled = False
    Pstate = False
End If

End Sub

Private Sub MTimer_Timer()

If MainForm.LblStrm.Caption = "1" Then
    Result = Stream01GetPosition(2)
Else
    Result = Stream02GetPosition(2)
End If

EndPos = CLng(Trim(LblEnd.Caption))
Poscnv = CLng(Trim(Result))

'avisamos que el plugin esta en uso
Pstate = True

'verificamos si la reproduccion de la posicion debe finalizar
If Poscnv >= EndPos Then
    If MainForm.LblStrm.Caption = "1" Then
        Stream01Stop
    Else
        Stream02Stop
    End If
    MTimer.Interval = 0
    MTimer.Enabled = False
    Pstate = False
End If

End Sub

Private Sub Timer1_Timer()

LBLTime.Caption = time$

End Sub

Private Sub TTimer_Timer()

If MainForm.LblStrm.Caption = "1" Then
    Result = Stream01GetPosition(2)
Else
    Result = Stream02GetPosition(2)
End If

EndPos = CLng(Trim(LblEnd.Caption))
Poscnv = CLng(Trim(Result))

'avisamos que el plugin esta en uso
Pstate = True

'verificamos si la reproduccion de la seccion debe finalizar
If Poscnv >= EndPos Then
    If MainForm.LblStrm.Caption = "1" Then
        Stream01Stop
    Else
        Stream02Stop
    End If
    TTimer.Interval = 0
    TTimer.Enabled = False
    Call InitHumedad
End If

End Sub
