VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "RmPlayer"
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicAnmR 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   210
      ScaleHeight     =   1080
      ScaleWidth      =   1155
      TabIndex        =   39
      Top             =   1245
      Width           =   1155
   End
   Begin VB.PictureBox PicAnmL 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   210
      ScaleHeight     =   1095
      ScaleWidth      =   1155
      TabIndex        =   38
      Top             =   135
      Width           =   1155
   End
   Begin PicClip.PictureClip PicRR 
      Left            =   2550
      Top             =   5265
      _ExtentX        =   1032
      _ExtentY        =   820
      _Version        =   393216
      Cols            =   25
   End
   Begin PicClip.PictureClip PicLL 
      Left            =   1875
      Top             =   5265
      _ExtentX        =   1005
      _ExtentY        =   820
      _Version        =   393216
      Cols            =   25
   End
   Begin PicClip.PictureClip SmallClip 
      Left            =   1650
      Top             =   4110
      _ExtentX        =   3731
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   14
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   4485
      ScaleHeight     =   240
      ScaleWidth      =   1230
      TabIndex        =   28
      Top             =   1380
      Width           =   1290
      Begin VB.PictureBox tr6 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   990
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   34
         ToolTipText     =   "Tiempo transcurrido"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox tr1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   45
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   33
         ToolTipText     =   "Tiempo transcurrido"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox tr2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   240
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   32
         ToolTipText     =   "Tiempo transcurrido"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox tr3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   420
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   31
         ToolTipText     =   "Tiempo transcurrido"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox tr4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   615
         ScaleHeight     =   180
         ScaleWidth      =   195
         TabIndex        =   30
         ToolTipText     =   "Tiempo transcurrido"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox tr5 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   810
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   29
         ToolTipText     =   "Tiempo transcurrido"
         Top             =   15
         Width           =   190
      End
   End
   Begin VB.Timer TmrClock 
      Left            =   1185
      Top             =   5265
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   1695
      ScaleHeight     =   240
      ScaleWidth      =   1575
      TabIndex        =   19
      Top             =   1965
      Width           =   1635
      Begin VB.PictureBox Cp5 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   780
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   27
         ToolTipText     =   "Hora actual"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Cp4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   585
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   26
         ToolTipText     =   "Hora actual"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Cp3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   390
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   25
         ToolTipText     =   "Hora actual"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Cp2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   210
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   24
         ToolTipText     =   "Hora actual"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Cp1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   15
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   23
         ToolTipText     =   "Hora actual"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Cp6 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   960
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   22
         ToolTipText     =   "Hora actual"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Cp7 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1155
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   21
         ToolTipText     =   "Hora actual"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Cp8 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1350
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   20
         ToolTipText     =   "Hora actual"
         Top             =   15
         Width           =   190
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   3705
      ScaleHeight     =   240
      ScaleWidth      =   1575
      TabIndex        =   10
      Top             =   1965
      Width           =   1635
      Begin VB.PictureBox Tp8 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1350
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   18
         ToolTipText     =   "Hora de lanzamiento del próximo tema"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Tp7 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1155
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   17
         ToolTipText     =   "Hora de lanzamiento del próximo tema"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Tp6 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   960
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   16
         ToolTipText     =   "Hora de lanzamiento del próximo tema"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Tp1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   15
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   15
         ToolTipText     =   "Hora de lanzamiento del próximo tema"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Tp2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   210
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   14
         ToolTipText     =   "Hora de lanzamiento del próximo tema"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Tp3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   390
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   13
         ToolTipText     =   "Hora de lanzamiento del próximo tema"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Tp4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   585
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   12
         ToolTipText     =   "Hora de lanzamiento del próximo tema"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Tp5 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   780
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   11
         ToolTipText     =   "Hora de lanzamiento del próximo tema"
         Top             =   15
         Width           =   190
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00000000&
      Height          =   315
      Left            =   3735
      ScaleHeight     =   255
      ScaleWidth      =   1995
      TabIndex        =   8
      Top             =   795
      Width           =   2055
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   15
         TabIndex        =   9
         ToolTipText     =   "Nombre de dispositivo en funcionamiento"
         Top             =   15
         Width           =   1965
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   1710
      ScaleHeight     =   240
      ScaleWidth      =   2535
      TabIndex        =   6
      Top             =   1380
      Width           =   2595
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   15
         TabIndex        =   7
         ToolTipText     =   "Nombre del tema actualmente en reproducción"
         Top             =   15
         Width           =   2475
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   165
      Left            =   1635
      ScaleHeight     =   105
      ScaleWidth      =   3060
      TabIndex        =   4
      Top             =   3345
      Width           =   3120
      Begin VB.PictureBox Ll 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   0
         ScaleHeight     =   150
         ScaleMode       =   0  'User
         ScaleWidth      =   105
         TabIndex        =   5
         Top             =   0
         Width           =   100
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   165
      Left            =   1635
      ScaleHeight     =   105
      ScaleMode       =   0  'User
      ScaleWidth      =   3060
      TabIndex        =   2
      Top             =   3495
      Width           =   3120
      Begin VB.PictureBox Lr 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   0
         ScaleHeight     =   150
         ScaleWidth      =   105
         TabIndex        =   3
         Top             =   0
         Width           =   100
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      Height          =   300
      Left            =   5460
      TabIndex        =   1
      ToolTipText     =   "Cerrar RadioMaker Mini-Player"
      Top             =   1980
      Width           =   330
   End
   Begin VB.PictureBox picMainSkin 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2490
      Left            =   0
      ScaleHeight     =   2490
      ScaleWidth      =   6165
      TabIndex        =   0
      Top             =   0
      Width           =   6165
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00000000&
         Height          =   840
         Left            =   1680
         ScaleHeight     =   52
         ScaleMode       =   0  'User
         ScaleWidth      =   101
         TabIndex        =   36
         Top             =   240
         Width           =   1575
         Begin VB.PictureBox Picfft1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00404000&
            BorderStyle     =   0  'None
            Height          =   690
            Left            =   30
            ScaleHeight     =   46
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   96
            TabIndex        =   37
            Top             =   45
            Width           =   1440
         End
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         Height          =   180
         Left            =   3435
         TabIndex        =   35
         Top             =   2025
         Width           =   225
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'viariabled de manejo de tiempo en relojes TopMenu
Dim LHora As String
Dim LMinutos As String
Dim LSegundos As String
Dim NHora(1 To 2) As String
Dim NMinutos(1 To 2) As String
Dim NSegundos(1 To 2) As String

'variabled de manejo de fecha en relojes TopMenu
Dim LMes As String
Dim LDia As String
Dim LAno As String
Dim NMes(1 To 2) As String
Dim NDia(1 To 2) As String
Dim NAno(1 To 4) As String

Dim Contador As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub RestoreDisplay()

'setea los displays predeterminados a 0
'por defecto 00:00:00.
        
'TIME DISPLAY FOR MINI-RM
Main.Tp1.Picture = Main.SmallClip.GraphicCell(10)
Main.Tp2.Picture = Main.SmallClip.GraphicCell(10)
Main.Tp3.Picture = Main.SmallClip.GraphicCell(12)
Main.Tp4.Picture = Main.SmallClip.GraphicCell(10)
Main.Tp5.Picture = Main.SmallClip.GraphicCell(10)
Main.Tp6.Picture = Main.SmallClip.GraphicCell(12)
Main.Tp7.Picture = Main.SmallClip.GraphicCell(10)
Main.Tp8.Picture = Main.SmallClip.GraphicCell(10)

Main.Cp1.Picture = Main.SmallClip.GraphicCell(10)
Main.Cp2.Picture = Main.SmallClip.GraphicCell(10)
Main.Cp3.Picture = Main.SmallClip.GraphicCell(12)
Main.Cp4.Picture = Main.SmallClip.GraphicCell(10)
Main.Cp5.Picture = Main.SmallClip.GraphicCell(10)
Main.Cp6.Picture = Main.SmallClip.GraphicCell(12)
Main.Cp7.Picture = Main.SmallClip.GraphicCell(10)
Main.Cp8.Picture = Main.SmallClip.GraphicCell(10)

Main.tr1.Picture = Main.SmallClip.GraphicCell(10)
Main.tr2.Picture = Main.SmallClip.GraphicCell(10)
Main.tr3.Picture = Main.SmallClip.GraphicCell(10)
Main.tr4.Picture = Main.SmallClip.GraphicCell(12)
Main.tr5.Picture = Main.SmallClip.GraphicCell(10)
Main.tr6.Picture = Main.SmallClip.GraphicCell(10)

End Sub

Private Sub MinClock(WNumber As String)

On Error GoTo NoClock

LHora = Left$(WNumber, 2)
LMinutos = Mid$(WNumber, 4, 2)
LSegundos = Right$(WNumber, 2)

'hora
NHora(1) = Left$(LHora, 1)
NHora(2) = Right$(LHora, 1)
'minutos
NMinutos(1) = Left$(LMinutos, 1)
NMinutos(2) = Right$(LMinutos, 1)
'segundos
NSegundos(1) = Left$(LSegundos, 1)
NSegundos(2) = Right$(LSegundos, 1)

'display the time
If NHora(1) = "0" Then Main.Cp1.Picture = Main.SmallClip.GraphicCell(0)
If NHora(1) = "1" Then Main.Cp1.Picture = Main.SmallClip.GraphicCell(1)
If NHora(1) = "2" Then Main.Cp1.Picture = Main.SmallClip.GraphicCell(2)

If NHora(2) = "0" Then Main.Cp2.Picture = Main.SmallClip.GraphicCell(0)
If NHora(2) = "1" Then Main.Cp2.Picture = Main.SmallClip.GraphicCell(1)
If NHora(2) = "2" Then Main.Cp2.Picture = Main.SmallClip.GraphicCell(2)
If NHora(2) = "3" Then Main.Cp2.Picture = Main.SmallClip.GraphicCell(3)
If NHora(2) = "4" Then Main.Cp2.Picture = Main.SmallClip.GraphicCell(4)
If NHora(2) = "5" Then Main.Cp2.Picture = Main.SmallClip.GraphicCell(5)
If NHora(2) = "6" Then Main.Cp2.Picture = Main.SmallClip.GraphicCell(6)
If NHora(2) = "7" Then Main.Cp2.Picture = Main.SmallClip.GraphicCell(7)
If NHora(2) = "8" Then Main.Cp2.Picture = Main.SmallClip.GraphicCell(8)
If NHora(2) = "9" Then Main.Cp2.Picture = Main.SmallClip.GraphicCell(9)

Main.Cp3.Picture = Main.SmallClip.GraphicCell(11)

If NMinutos(1) = "0" Then Main.Cp4.Picture = Main.SmallClip.GraphicCell(0)
If NMinutos(1) = "1" Then Main.Cp4.Picture = Main.SmallClip.GraphicCell(1)
If NMinutos(1) = "2" Then Main.Cp4.Picture = Main.SmallClip.GraphicCell(2)
If NMinutos(1) = "3" Then Main.Cp4.Picture = Main.SmallClip.GraphicCell(3)
If NMinutos(1) = "4" Then Main.Cp4.Picture = Main.SmallClip.GraphicCell(4)
If NMinutos(1) = "5" Then Main.Cp4.Picture = Main.SmallClip.GraphicCell(5)

If NMinutos(2) = "0" Then Main.Cp5.Picture = Main.SmallClip.GraphicCell(0)
If NMinutos(2) = "1" Then Main.Cp5.Picture = Main.SmallClip.GraphicCell(1)
If NMinutos(2) = "2" Then Main.Cp5.Picture = Main.SmallClip.GraphicCell(2)
If NMinutos(2) = "3" Then Main.Cp5.Picture = Main.SmallClip.GraphicCell(3)
If NMinutos(2) = "4" Then Main.Cp5.Picture = Main.SmallClip.GraphicCell(4)
If NMinutos(2) = "5" Then Main.Cp5.Picture = Main.SmallClip.GraphicCell(5)
If NMinutos(2) = "6" Then Main.Cp5.Picture = Main.SmallClip.GraphicCell(6)
If NMinutos(2) = "7" Then Main.Cp5.Picture = Main.SmallClip.GraphicCell(7)
If NMinutos(2) = "8" Then Main.Cp5.Picture = Main.SmallClip.GraphicCell(8)
If NMinutos(2) = "9" Then Main.Cp5.Picture = Main.SmallClip.GraphicCell(9)

Main.Cp6.Picture = Main.SmallClip.GraphicCell(11)

If NSegundos(1) = "0" Then Main.Cp7.Picture = Main.SmallClip.GraphicCell(0)
If NSegundos(1) = "1" Then Main.Cp7.Picture = Main.SmallClip.GraphicCell(1)
If NSegundos(1) = "2" Then Main.Cp7.Picture = Main.SmallClip.GraphicCell(2)
If NSegundos(1) = "3" Then Main.Cp7.Picture = Main.SmallClip.GraphicCell(3)
If NSegundos(1) = "4" Then Main.Cp7.Picture = Main.SmallClip.GraphicCell(4)
If NSegundos(1) = "5" Then Main.Cp7.Picture = Main.SmallClip.GraphicCell(5)

If NSegundos(2) = "0" Then Main.Cp8.Picture = Main.SmallClip.GraphicCell(0)
If NSegundos(2) = "1" Then Main.Cp8.Picture = Main.SmallClip.GraphicCell(1)
If NSegundos(2) = "2" Then Main.Cp8.Picture = Main.SmallClip.GraphicCell(2)
If NSegundos(2) = "3" Then Main.Cp8.Picture = Main.SmallClip.GraphicCell(3)
If NSegundos(2) = "4" Then Main.Cp8.Picture = Main.SmallClip.GraphicCell(4)
If NSegundos(2) = "5" Then Main.Cp8.Picture = Main.SmallClip.GraphicCell(5)
If NSegundos(2) = "6" Then Main.Cp8.Picture = Main.SmallClip.GraphicCell(6)
If NSegundos(2) = "7" Then Main.Cp8.Picture = Main.SmallClip.GraphicCell(7)
If NSegundos(2) = "8" Then Main.Cp8.Picture = Main.SmallClip.GraphicCell(8)
If NSegundos(2) = "9" Then Main.Cp8.Picture = Main.SmallClip.GraphicCell(9)

Exit Sub

NoClock:
End Sub

Private Sub SetAudioLevel(WLeft, WRight)

'level scope meter sub

Dim l, Lft As Integer
Dim r, Rgt As Integer
Dim i As Integer
Static ZMax%, RMax%

On Error Resume Next
'right level meter
If WRight > 180 Then
    RMax = (WRight * 24) + 100 'clip
Else
    RMax = (WRight * 24)
End If

'left level meter
If WLeft > 180 Then
    ZMax = (WLeft * 24) + 100  'clip
Else
    ZMax = (WLeft * 24)
End If

Lr.Width = RMax
Ll.Width = ZMax

'NOT IMPLEMENTED YET

End Sub

Private Sub Command1_Click()

Main.TmrClock.Interval = 0
Main.TmrClock.Enabled = False

Main.WindowState = 1
'Unload Main
 
End Sub

Private Sub Form_Load()
    
    Dim WindowRegion As Long
    
    'load led1
    'Picture1.Picture = LoadResPicture("BACK_LED", 0)
    'Ll.Picture = LoadResPicture("FRONT_LED", 0)
    'Ll.Width = 1
    'load led2
    'Picture2.Picture = LoadResPicture("BACK_LED", 0)
    'Lr.Picture = LoadResPicture("FRONT_LED", 0)
    'Lr.Width = 1
    
    'load nums
    SmallClip.Picture = LoadResPicture("NUM_SMALL", 0)
    SmallClip.Cols = 14
    'load buff (for level animation)
    PicLL.Picture = LoadResPicture("BUFF_L", 0)
    PicLL.Cols = 25
    PicRR.Picture = LoadResPicture("BUFF_R", 0)
    PicRR.Cols = 25
    'some buff pics
    PicAnmR.Picture = PicRR.GraphicCell(0)
    PicAnmL.Picture = PicLL.GraphicCell(0)
    
    'paint clocks
    Call RestoreDisplay

    ' I set all these settings here so you won't forget
    ' them and have a non-working demo... Set them in
    ' design time
    picMainSkin.ScaleMode = vbPixels
    picMainSkin.AutoRedraw = True
    picMainSkin.AutoSize = True
    picMainSkin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
        
    Set picMainSkin.Picture = LoadResPicture("RM_MIN_SPC", 0)
    
    Me.Width = picMainSkin.Width
    Me.Height = picMainSkin.Height
    
    WindowRegion = MakeRegion(picMainSkin)
    SetWindowRgn Me.hWnd, WindowRegion, True
    
    'real clock timer
    TmrClock.Enabled = True
    TmrClock.Interval = 1000
    
    Me.Left = Screen.Width - Me.Width
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    
End Sub

Private Sub Label2_Change()

Call RestoreDisplay

End Sub

Private Sub PicAnmL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

      ' Pass the handling of the mouse down message to
      ' the (non-existing really) form caption, so that
      ' the form itself will be dragged when the picture is dragged.
      '
      ' If you have Win 98, Make sure that the "Show window
      ' contents while dragging" display setting is on for nice results.
      
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub

Private Sub PicAnmR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

      ' Pass the handling of the mouse down message to
      ' the (non-existing really) form caption, so that
      ' the form itself will be dragged when the picture is dragged.
      '
      ' If you have Win 98, Make sure that the "Show window
      ' contents while dragging" display setting is on for nice results.
      
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub

Private Sub picMainSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
      ' Pass the handling of the mouse down message to
      ' the (non-existing really) form caption, so that
      ' the form itself will be dragged when the picture is dragged.
      '
      ' If you have Win 98, Make sure that the "Show window
      ' contents while dragging" display setting is on for nice results.
      
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub

Private Sub TmrClock_Timer()

Dim RealClock As String

RealClock = Time$

'update the clock
Call MinClock(RealClock)

End Sub
