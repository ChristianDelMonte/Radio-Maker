VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmLiveUpd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ONLY development - Live Update"
   ClientHeight    =   6165
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   6120
   ControlBox      =   0   'False
   Icon            =   "liveu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pc1 
      AutoSize        =   -1  'True
      Height          =   1560
      Left            =   60
      Picture         =   "liveu.frx":030A
      ScaleHeight     =   1500
      ScaleWidth      =   5940
      TabIndex        =   6
      Top             =   30
      Width           =   6000
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6780
      Top             =   1620
   End
   Begin VB.Timer Timer2 
      Left            =   6300
      Top             =   1620
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   4890
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Max             =   3
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5790
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Live Update Control Panel"
      Height          =   4065
      Left            =   60
      TabIndex        =   0
      Top             =   1650
      Width           =   5985
      Begin VB.CommandButton Command2 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   4620
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3570
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<< &Buscar >>"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3570
         Width           =   1965
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   180
         TabIndex        =   8
         Top             =   2820
         Width           =   5595
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   180
         TabIndex        =   7
         Top             =   2610
         Width           =   5595
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<< Detenido >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   120
         TabIndex        =   3
         Top             =   510
         Width           =   5745
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6660
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmLiveUpd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim TransferSuccess As Boolean
    
    Label1.Caption = "Buscando actualizaciones..."
    UpdateTime = 0
    Timer2.Interval = 1000
    Command1.Enabled = False
    Command2.Enabled = False
    ProgressBar1.value = 1
    
    status$ = "Chequeando..."
    
    TransferSuccess = GetInternetFile(Inet1, "http://www.interbandas.com.ar/liveupdate/RM100_UpdateVer.dat", App.Path & "\Update")
    'TransferSuccess = GetInternetFile(Inet1, "http://www.liveupdate.interbandas.com.ar/RM100_UpdateVer.dat", App.Path & "\Update")
    If TransferSuccess = False Then
        ProgressBar1.value = 3
        Timer2.Interval = 500
        Exit Sub
    End If
       
    ProgressBar1.value = 2
    status$ = "Listo."
    
        Open App.Path & "\Update\RM100_UpdateVer.dat" For Input As #22
            Input #22, updatever$
        Close #22
      
    If updatever$ > myVer Then
        Label1.Caption = "NUEVO UPDATE ENCONTRADO: << descargando archivo >>"
        Label2.Caption = "Versión Actual: " & myVer
        Label3.Caption = "Versión Nueva: " + updatever
        Timer2.Interval = 500
    Else
        Label1.Caption = "No hay Update de nueva versión disponible."
        ProgressBar1.value = 3
        Command1.Enabled = True
        Command2.Enabled = True
        Timer2.Interval = 0
        Exit Sub
    End If

    status$ = "Extrayendo actualizacion..."
   
        TransferSuccess = GetInternetFile(Inet1, "http://www.interbandas.com.ar/liveupdate/RM100_update.cbm", App.Path & "\Update")
        'TransferSuccess = GetInternetFile(Inet1, "http://www.liveupdate.interbandas.com.ar/RM100_update.cbm", App.Path & "\Update")
        
    If TransferSuccess = False Then
        ProgressBar1.value = 3
        Command1.Enabled = True
        Command2.Enabled = True
        Timer2.Interval = 0
        Exit Sub
    Else
        ProgressBar1.value = 3
        Timer2.Interval = 0
        x = MsgBox("Live Update Completado!", vbInformation)
        Command1.Enabled = True
        Command2.Enabled = True
        FrmWarning.Show
    End If

End Sub

Private Sub Command2_Click()

End

End Sub

Private Sub Form_Load()

'///abrimos un archivo con el detalle de la version actual del programa
On Error GoTo er
Open App.Path & "\VerChk.dat" For Input As #31
Input #31, myVer, myPath
Close #31

myVer = Trim(myVer)

status$ = "Idle"
UpdateTime = 0
Exit Sub

er:
MsgBox "No se encontró la información de ninguna versión disponible.", vbCritical
End

End Sub

Private Sub Timer1_Timer()
If Inet1.StillExecuting = False Then
    StatusBar1.Panels(1).Text = "Status: Idle"
Else
    StatusBar1.Panels(1).Text = "Status: " & status$
End If

End Sub

Private Sub Timer2_Timer()
    UpdateTime = UpdateTime + 1
    StatusBar1.Panels(2).Text = "Download Time:" & Str$(UpdateTime) & " Seconds"
End Sub
