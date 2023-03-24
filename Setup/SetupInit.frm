VERSION 5.00
Begin VB.Form SetupInit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "XXX"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7665
   ForeColor       =   &H8000000F&
   Icon            =   "SetupInit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2505
      TabIndex        =   13
      Text            =   "C:\TEMP"
      Top             =   2040
      Width           =   3930
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Cambiar..."
      Height          =   330
      Left            =   6510
      TabIndex        =   12
      Top             =   2040
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione el tipo de Instalación que desea realizar"
      Height          =   1275
      Left            =   2520
      TabIndex        =   7
      Top             =   2490
      Width           =   5010
      Begin VB.CheckBox Check3 
         Caption         =   "Completa"
         Height          =   195
         Left            =   3735
         TabIndex        =   10
         Top             =   315
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Media"
         Height          =   195
         Left            =   2070
         TabIndex        =   9
         Top             =   315
         Width           =   780
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mínima"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   315
         Width           =   915
      End
      Begin VB.Label LblDesc 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   600
         Left            =   135
         TabIndex        =   11
         Top             =   585
         Width           =   4740
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cambiar..."
      Height          =   330
      Left            =   6525
      TabIndex        =   6
      Top             =   1080
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Text            =   "C:\RM100"
      Top             =   1080
      Width           =   3930
   End
   Begin VB.CommandButton Command2 
      Caption         =   "C&ancelar"
      Height          =   375
      Left            =   6255
      TabIndex        =   2
      Top             =   3960
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Continuar >>"
      Height          =   375
      Left            =   4815
      TabIndex        =   1
      Top             =   3960
      Width           =   1320
   End
   Begin VB.PictureBox PicSetup 
      AutoSize        =   -1  'True
      Height          =   4305
      Left            =   90
      ScaleHeight     =   4245
      ScaleWidth      =   2250
      TabIndex        =   0
      Top             =   45
      Width           =   2310
   End
   Begin VB.Label APNName 
      Caption         =   "Label4"
      Height          =   210
      Left            =   2610
      TabIndex        =   15
      Top             =   4125
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione el directorio temporal donde se extraerán los datos para su posterior instalación."
      Height          =   465
      Left            =   2505
      TabIndex        =   14
      Top             =   1530
      Width           =   5055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   2475
      X2              =   7560
      Y1              =   3855
      Y2              =   3855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "XX"
      Height          =   390
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "XX"
      Height          =   405
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   4875
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   2475
      X2              =   7560
      Y1              =   3825
      Y2              =   3825
   End
End
Attribute VB_Name = "SetupInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

If Check1.value = 1 Then
    Check2.value = 0
    Check3.value = 0
    LblDesc.Caption = "Realiza la instalación mínima de " & APNName.Caption & ". Esto Incluye solamente el Programa Principal y los drivers necesarios para su funcionamiento."
End If

End Sub

Private Sub Check2_Click()

If Check2.value = 1 Then
    Check1.value = 0
    Check3.value = 0
    LblDesc.Caption = "Realiza la instalación media de " & APNName.Caption & ". Esto Incluye el Programa Principal, drivers y plug-ins necesarios para su funcionamiento."
End If

End Sub

Private Sub Check3_Click()

If Check3.value = 1 Then
    Check2.value = 0
    Check1.value = 0
    LblDesc.Caption = "Realiza la instalación full de " & APNName.Caption & ". Esto Incluye el Programa Principal, drivers, plug-ins, temas de ejemplo y documentación."
End If

End Sub

Private Sub Command1_Click()

'chequeos necesarios
If Text1.Text = "" Or Text1.Text = " " Then
    MsgBox "La carpeta de Instalación no es correcta.", vbCritical, App.ProductName
    Exit Sub
End If
If Text2.Text = "" Or Text2.Text = " " Then
    MsgBox "La carpeta de instalación temporal no es correcta.", vbCritical, App.ProductName
    Exit Sub
End If

'continuamos
Dim Msg, Msg0, Msg1, Msg2, Msg3, Msg4
Dim Style, Title, Response

Msg0 = "Todos los archivos que se encuentren dentro de"
Msg1 = "la carpeta seleccionada como Temporal serán"
Msg2 = "eliminados una vez finalizada la instalación."
Msg3 = " "
Msg4 = "¿Desea continuar con la instalación?"
Msg = Msg0 & Chr(13) & Msg1 & Chr(13) & Msg2 & Chr(13) & Msg3 & Chr(13) & Msg4
Style = vbYesNo + vbCritical + vbDefaultButton2
Title = App.ProductName
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
    FrmCopy.Show
Else
    Exit Sub
End If

End Sub

Private Sub Command2_Click()

End

End Sub

Private Sub Command3_Click()

Dim Result As String
Result = BrowseForFolder("Seleccione la nueva carpeta.")
Text1.Text = Result

End Sub

Private Sub Command4_Click()

Dim Result As String
Result = BrowseForFolder("Seleccione la carpeta temporal.")
Text2.Text = Result

End Sub

Private Sub Form_Load()

Dim APData As StpIniFile
Dim APName As String, APVersion As String, APDir As String
Dim APCompany As String, APTitle As String

PicSetup.Picture = LoadResPicture("RM_SETUP1", 0)

'cargamos los datos del archivo ini
APData = ReadIniFile(App.Path, 1)

APName = Trim(Desencriptar(DefPass, APData.AppTitle)): APNName.Caption = APName
If APName = "" Or APName = " " Then
    MsgBox "No se encontró el archivo de datos SETUP.INI", vbCritical, App.ProductName
    End
End If

APVersion = Trim(Desencriptar(DefPass, APData.AppVersion))
APDir = Trim(Desencriptar(DefPass, APData.AppDefDir))
APCompany = Trim(Desencriptar(DefPass, APData.AppCompany))
APTitle = Trim(Desencriptar(DefPass, APData.FrmTitle))

SetupInit.Caption = APTitle & " - Programa de instalación."
Label2.Caption = "Bienvenidos al programa de Instalación de " & APName & " " & APVersion
Label3.Caption = "El programa se instalará por defecto en el siguiente directorio. Si desea instalar " & APName & " en otra carpeta haga click en Cambiar."
SetupInit.Text1.Text = "C:\" & Trim(Desencriptar(DefPass, APData.AppDefDir))
LblDesc.Caption = "Realiza la instalación full de " & APName & ". Esto Incluye el Programa Principal, drivers, plug-ins, temas de ejemplo y documentación."

End Sub
