VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RM100 - CD Ripper"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4695
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Examinar"
      Height          =   315
      Left            =   3645
      TabIndex        =   12
      Top             =   5220
      Width           =   960
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Cerrar"
      Height          =   315
      Left            =   3600
      TabIndex        =   14
      Top             =   5670
      Width           =   1005
   End
   Begin VB.CommandButton cmdRip 
      Caption         =   "< E X T R A E R >"
      Height          =   315
      Left            =   90
      TabIndex        =   13
      Top             =   5670
      Width           =   2340
   End
   Begin VB.ListBox TrackList 
      Height          =   2085
      Left            =   90
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1080
      Width           =   4515
   End
   Begin VB.TextBox txtOutPath 
      Height          =   315
      Left            =   810
      TabIndex        =   11
      Top             =   5220
      Width           =   2805
   End
   Begin VB.ComboBox DriveList 
      Height          =   315
      Left            =   105
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3300
      Width           =   4485
   End
   Begin VB.Frame Frame1 
      Caption         =   "Extraer a"
      Height          =   1440
      Left            =   105
      TabIndex        =   0
      Top             =   3720
      Width           =   4500
      Begin VB.ComboBox cmbQuality 
         Height          =   315
         ItemData        =   "frmMain.frx":08CA
         Left            =   840
         List            =   "frmMain.frx":08EC
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   3540
      End
      Begin VB.OptionButton optWAV 
         Caption         =   "WAV"
         Height          =   255
         Left            =   1845
         TabIndex        =   4
         Top             =   0
         Width           =   675
      End
      Begin VB.OptionButton optMP3 
         Caption         =   "MP3"
         Height          =   255
         Left            =   1005
         TabIndex        =   3
         Top             =   0
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.ComboBox cmbBitrate 
         Height          =   315
         ItemData        =   "frmMain.frx":0936
         Left            =   2460
         List            =   "frmMain.frx":0964
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   960
         Width           =   1890
      End
      Begin VB.CheckBox chkPrivate 
         Caption         =   "Privado"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   1020
         Width           =   855
      End
      Begin VB.CheckBox chkOriginal 
         Caption         =   "Original"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   780
         Width           =   855
      End
      Begin VB.CheckBox chkCRC 
         Caption         =   "CRC"
         Height          =   255
         Left            =   135
         TabIndex        =   7
         Top             =   1020
         Width           =   1335
      End
      Begin VB.CheckBox chkCopyright 
         Caption         =   "Copyrighted"
         Height          =   255
         Left            =   135
         TabIndex        =   6
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Bitrate"
         Height          =   195
         Left            =   2460
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Calidad"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   45
      Picture         =   "frmMain.frx":09A7
      ToolTipText     =   "Ayuda"
      Top             =   45
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   4590
      X2              =   45
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Label4 
      Caption         =   "(c) 2001 4D Software Inc."
      Height          =   240
      Left            =   45
      TabIndex        =   18
      Top             =   675
      Width           =   3300
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   855
      Picture         =   "frmMain.frx":0DE9
      Top             =   0
      Width           =   3750
   End
   Begin VB.Label Label1 
      Caption         =   "Destino:"
      Height          =   255
      Left            =   105
      TabIndex        =   17
      Top             =   5265
      Width           =   675
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   45
      Picture         =   "frmMain.frx":1EE1
      ToolTipText     =   "Ayuda"
      Top             =   45
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuChangeName 
         Caption         =   "Rename"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' This project and modifications to LAME_ENC.DLL and
' AKRRIP32.DLL was made by Arto Rusanen
' http://www.4dsoftware.8m.com

' Credits...

' LAME was originally developed by Mike Cheng (www.uq.net.au/~zzmcheng).
' Now maintained by Mark Taylor (www.sulaco.org/mp3).

' You can find LAME and its source from
' http://www.mp3dev.org

' AKRip was orginally made by Andy Key and you can find AKRip and its source
' from http://akrip.sourceforge.net/

Option Explicit

Private Sub Form_Load()
    
  ' Initialize Exception Filter
  SetUnhandledExceptionFilter AddressOf MyExceptionFilter

  cmbQuality.ListIndex = 1
  cmbBitrate.ListIndex = 9
  
  'Find CD Drive adapters
  Dim DriveCount As Long
  Dim MyInfo As CDREC
  ChDrive App.Path
  ChDir App.Path

  DriveCount = GetNumAdapters + 1
  
  Dim i As Long
  For i = 1 To DriveCount
    MyInfo = GetDriveInformation(i - 1)
    DriveList.AddItem StripNullsArray(MyInfo.id)
  Next i
  
  DriveList.ListIndex = 0
  txtOutPath.Text = App.Path
End Sub

Private Sub DriveList_Click()
  ' Init selected drive and read its TOC
  On Error Resume Next
  TrackList.Clear
  
  Call DeInitCDDrive
  If Not InitCDDrive(DriveList.ListIndex) Then Exit Sub
  
  Dim i As Long
  i = 1
  Do While MSB2LONG(DiscToc.tracks(i + 1).addr) <> 0
    TrackList.AddItem "Pista " & i
    i = i + 1
  Loop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image3.Visible = False
Image2.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Remove exception filter...
  SetUnhandledExceptionFilter 0
End Sub


Private Sub cmdRip_Click()
  If optMP3.value Then
    ' Fill beConfig structure....
    Dim beConfig As PBE_CONFIG
    beConfig.dwConfig = BE_CONFIG_LAME
    
    With beConfig.format.LHV1
      .dwStructVersion = 1
      .dwStructSize = Len(beConfig)
      .dwSampleRate = 44100         '// INPUT FREQUENCY
      .dwReSampleRate = 0           '// DON"T RESAMPLE
      .nMode = BE_MP3_MODE_JSTEREO  '// OUTPUT IN STREO
      .dwBitrate = val(cmbBitrate.Text)
      .dwMpegVersion = MPEG1        '// MPEG VERSION (I or II)
      .dwPsyModel = 0               '// USE DEFAULT PSYCHOACOUSTIC MODEL
      .dwEmphasis = 0               '// NO EMPHASIS TURNED ON
      .bNoRes = True                '// No Bit resorvoir
      
      .bCopyright = chkCopyright.value = 1
      .bCRC = chkCRC.value = 1
      .bOriginal = chkOriginal.value = 1
      .bPrivate = chkPrivate.value = 1
      
      Select Case cmbQuality.ListIndex  '// QUALITY PRESET SETTING
        Case 0: .nPreset = LQP_LOW_QUALITY
        Case 1: .nPreset = LQP_NORMAL_QUALITY
        Case 2: .nPreset = LQP_HIGH_QUALITY
        Case 3: .nPreset = LQP_VOICE_QUALITY
        Case 4: .nPreset = LQP_PHONE
        Case 5: .nPreset = LQP_RADIO
        Case 6: .nPreset = LQP_TAPE
        Case 7: .nPreset = LQP_HIFI
        Case 8: .nPreset = LQP_CD
        Case 9: .nPreset = LQP_STUDIO
      End Select
    End With
  End If
  
  ' Rip all tracks that are selected
  Dim TrackNo As Long
  For TrackNo = 1 To TrackList.ListCount
    If Cancelled Then Exit For
    If TrackList.Selected(TrackNo - 1) Then
      If optMP3.value Then
        Call RipMP3(AddSlash(txtOutPath.Text) & TrackList.List(TrackNo - 1) & ".mp3", MSB2LONG(DiscToc.tracks(TrackNo).addr), MSB2LONG(DiscToc.tracks(TrackNo + 1).addr), beConfig)
      Else
        Call RipWAV(AddSlash(txtOutPath.Text) & TrackList.List(TrackNo - 1) & ".wav", MSB2LONG(DiscToc.tracks(TrackNo).addr), MSB2LONG(DiscToc.tracks(TrackNo + 1).addr))
      End If
    End If
  Next TrackNo
End Sub

Private Sub cmdQuit_Click()
 
 frmMain.WindowState = 1
 'Unload Me
 
End Sub

Private Sub Command1_Click()
  txtOutPath.Text = BrowseForFolder
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image3.Visible = True
Image2.Visible = False

End Sub

' Change Name
Private Sub mnuChangeName_Click()
  If TrackList.ListCount > 0 Then _
    TrackList.List(TrackList.ListIndex) = InputBox("Nuevo nombre...")
End Sub

Private Sub TrackList_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then Call mnuChangeName_Click
End Sub

Private Sub TrackList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuPopup
End Sub

