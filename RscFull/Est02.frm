VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form Est02 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "ESTACION 02 - Detenido"
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   15330
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Est02.frx":0000
   ScaleHeight     =   5160
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton E2Play 
      Height          =   375
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   122
      ToolTipText     =   "Reproducir"
      Top             =   4605
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton E2Pause 
      Height          =   375
      Left            =   4650
      Style           =   1  'Graphical
      TabIndex        =   121
      ToolTipText     =   "Pausar"
      Top             =   4605
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.PictureBox E2p7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2730
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   120
      Top             =   900
      Width           =   190
   End
   Begin VB.PictureBox E2p8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2925
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   119
      Top             =   900
      Width           =   190
   End
   Begin VB.PictureBox E2p9 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3105
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   118
      Top             =   900
      Width           =   190
   End
   Begin VB.PictureBox E2p10 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3300
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   117
      Top             =   900
      Width           =   190
   End
   Begin VB.PictureBox E2p11 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3495
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   116
      Top             =   900
      Width           =   190
   End
   Begin VB.PictureBox E2p1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2730
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   115
      Top             =   615
      Width           =   190
   End
   Begin VB.PictureBox E2p2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2925
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   114
      Top             =   615
      Width           =   190
   End
   Begin VB.PictureBox E2p3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3105
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   113
      Top             =   615
      Width           =   190
   End
   Begin VB.PictureBox E2p4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3300
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   112
      Top             =   615
      Width           =   190
   End
   Begin VB.PictureBox E2p5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3495
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   111
      Top             =   615
      Width           =   190
   End
   Begin VB.CommandButton E2Stop 
      Height          =   375
      Left            =   5250
      Style           =   1  'Graphical
      TabIndex        =   110
      ToolTipText     =   "Detener"
      Top             =   4605
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton E2New 
      Height          =   375
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   109
      ToolTipText     =   "Nuevo archivo de estacion"
      Top             =   4605
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton E2Open 
      Height          =   375
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   108
      ToolTipText     =   "Abrir archivo de Estacion"
      Top             =   4605
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton E2Save 
      Height          =   375
      Left            =   7020
      Style           =   1  'Graphical
      TabIndex        =   107
      ToolTipText     =   "Guardar Archivo de Estacion"
      Top             =   4605
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   210
      Left            =   3945
      ScaleHeight     =   150
      ScaleWidth      =   3240
      TabIndex        =   105
      Top             =   465
      Width           =   3300
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
         ScaleWidth      =   3240
         TabIndex        =   106
         Top             =   0
         Width           =   3240
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   210
      Left            =   3945
      ScaleHeight     =   150
      ScaleMode       =   0  'User
      ScaleWidth      =   3240
      TabIndex        =   103
      Top             =   705
      Width           =   3300
      Begin VB.PictureBox Lr 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   0
         ScaleHeight     =   150
         ScaleWidth      =   3240
         TabIndex        =   104
         Top             =   0
         Width           =   3240
      End
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   8145
      ScaleHeight     =   1305
      ScaleWidth      =   3780
      TabIndex        =   84
      Top             =   480
      Width           =   3780
      Begin VB.PictureBox Picfft2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   330
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   135
         TabIndex        =   87
         Top             =   495
         Width           =   2025
      End
      Begin VB.PictureBox pcontup 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3150
         Picture         =   "Est02.frx":25536
         ScaleHeight     =   255
         ScaleWidth      =   405
         TabIndex        =   86
         ToolTipText     =   "Modo CONTINUO activado"
         Top             =   225
         Width           =   405
      End
      Begin VB.PictureBox pcontdw 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3150
         Picture         =   "Est02.frx":25B0E
         ScaleHeight     =   255
         ScaleWidth      =   405
         TabIndex        =   85
         ToolTipText     =   "Modo CONTINUO desactivado"
         Top             =   225
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Lfft 
         BackStyle       =   0  'Transparent
         Caption         =   "FFT"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   360
         TabIndex        =   102
         ToolTipText     =   "Modo FFT"
         Top             =   225
         Width           =   375
      End
      Begin VB.Label Lspc 
         BackStyle       =   0  'Transparent
         Caption         =   "SPC"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   765
         TabIndex        =   101
         ToolTipText     =   "Modo Espectro"
         Top             =   225
         Width           =   375
      End
      Begin VB.Label Lspcz 
         BackStyle       =   0  'Transparent
         Caption         =   "Izq"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1305
         TabIndex        =   100
         ToolTipText     =   "Espectro izquierdo"
         Top             =   225
         Width           =   285
      End
      Begin VB.Label Lspcd 
         BackStyle       =   0  'Transparent
         Caption         =   "Der"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1620
         TabIndex        =   99
         ToolTipText     =   "Espectro derecho"
         Top             =   225
         Width           =   285
      End
      Begin VB.Label Lspcb 
         BackStyle       =   0  'Transparent
         Caption         =   "Amb"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1980
         TabIndex        =   98
         ToolTipText     =   "Espectro combinado"
         Top             =   225
         Width           =   375
      End
      Begin VB.Label LAplay 
         BackStyle       =   0  'Transparent
         Caption         =   "Autoplay"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   2430
         TabIndex        =   97
         ToolTipText     =   "Autoreproducción al hacer click"
         Top             =   225
         Width           =   645
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F-In/Out:"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   2385
         TabIndex        =   96
         Top             =   495
         Width           =   690
      End
      Begin VB.Label LFin 
         BackStyle       =   0  'Transparent
         Caption         =   "Auto"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   3150
         TabIndex        =   95
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CUE:"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   2430
         TabIndex        =   94
         Top             =   720
         Width           =   645
      End
      Begin VB.Label LCue 
         BackStyle       =   0  'Transparent
         Caption         =   "Man"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   3150
         TabIndex        =   93
         Top             =   720
         Width           =   375
      End
      Begin VB.Label fft2 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   2430
         TabIndex        =   92
         Top             =   990
         Width           =   105
      End
      Begin VB.Label fft4 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   2610
         TabIndex        =   91
         Top             =   990
         Width           =   150
      End
      Begin VB.Label fft6 
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   2835
         TabIndex        =   90
         Top             =   990
         Width           =   150
      End
      Begin VB.Label fft8 
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   3060
         TabIndex        =   89
         Top             =   990
         Width           =   150
      End
      Begin VB.Label fft10 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   3285
         TabIndex        =   88
         Top             =   990
         Width           =   195
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "CUE"
      ForeColor       =   &H000000FF&
      Height          =   870
      Left            =   12015
      TabIndex        =   77
      Top             =   2505
      Width           =   3195
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   81
         Text            =   "00:00:00"
         Top             =   495
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   80
         Text            =   "00:00:00"
         Top             =   495
         Width           =   975
      End
      Begin VB.CommandButton Command9 
         Caption         =   "<M"
         Height          =   300
         Left            =   1170
         TabIndex        =   79
         ToolTipText     =   "Marcar la posicion de inicio de CUE"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command10 
         Caption         =   "<M"
         Height          =   300
         Left            =   2655
         TabIndex        =   78
         ToolTipText     =   "Marcar la posicion de fin de CUE"
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CUE Inicio"
         Height          =   255
         Left            =   165
         TabIndex        =   83
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CUE Final"
         Height          =   255
         Left            =   1665
         TabIndex        =   82
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Posicionamiento "
      ForeColor       =   &H000000FF&
      Height          =   1020
      Left            =   8055
      TabIndex        =   71
      Top             =   3405
      Width           =   7155
      Begin ComctlLib.Slider E2Pos 
         Height          =   255
         Left            =   90
         TabIndex        =   72
         Top             =   270
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   450
         _Version        =   327682
         Max             =   100
         TickFrequency   =   5
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "En proceso:"
         Height          =   255
         Left            =   210
         TabIndex        =   76
         Top             =   720
         Width           =   915
      End
      Begin VB.Label LblCurrent 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1125
         TabIndex        =   75
         Top             =   675
         Width           =   1170
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Finalización:"
         Height          =   255
         Left            =   4950
         TabIndex        =   74
         Top             =   720
         Width           =   915
      End
      Begin VB.Label LblEnd 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5850
         TabIndex        =   73
         Top             =   675
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Paneo"
      ForeColor       =   &H000000FF&
      Height          =   1020
      Left            =   12015
      TabIndex        =   66
      Top             =   1440
      Width           =   3195
      Begin VB.CommandButton CmdAutoPan 
         Caption         =   "AUTO PANEO"
         Height          =   255
         Left            =   960
         TabIndex        =   67
         ToolTipText     =   "Paneo Izq>Der - Der>Izq - automatico"
         Top             =   660
         Width           =   1215
      End
      Begin ComctlLib.Slider E2Slide 
         Height          =   255
         Left            =   270
         TabIndex        =   68
         Top             =   270
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   450
         _Version        =   327682
         Max             =   100
         SelStart        =   50
         TickFrequency   =   5
         Value           =   50
      End
      Begin VB.Label Label3 
         Caption         =   "Dr"
         Height          =   255
         Left            =   2880
         TabIndex        =   70
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Iz"
         Height          =   255
         Left            =   90
         TabIndex        =   69
         Top             =   300
         Width           =   165
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Volumen"
      ForeColor       =   &H000000FF&
      Height          =   915
      Left            =   12015
      TabIndex        =   62
      Top             =   480
      Width           =   3195
      Begin VB.CommandButton CmdFOut 
         Caption         =   "FADE OUT"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2115
         TabIndex        =   64
         ToolTipText     =   "Fade-Out automatico"
         Top             =   570
         Width           =   975
      End
      Begin VB.CommandButton CmdFIN 
         Caption         =   "FADE IN"
         Enabled         =   0   'False
         Height          =   255
         Left            =   90
         TabIndex        =   63
         ToolTipText     =   "Fade-In automatico"
         Top             =   570
         Width           =   855
      End
      Begin ComctlLib.Slider E2Vol 
         Height          =   255
         Left            =   30
         TabIndex        =   65
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         _Version        =   327682
         Max             =   100
         TickFrequency   =   5
      End
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   0
      Left            =   8445
      Max             =   20
      TabIndex        =   61
      Top             =   1830
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   1
      Left            =   8760
      Max             =   20
      TabIndex        =   60
      Top             =   1830
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   2
      Left            =   9075
      Max             =   20
      TabIndex        =   59
      Top             =   1830
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   3
      Left            =   9390
      Max             =   20
      TabIndex        =   58
      Top             =   1830
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   4
      Left            =   9705
      Max             =   20
      TabIndex        =   57
      Top             =   1830
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   5
      Left            =   10020
      Max             =   20
      TabIndex        =   56
      Top             =   1830
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   6
      Left            =   10335
      Max             =   20
      TabIndex        =   55
      Top             =   1830
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   7
      Left            =   10650
      Max             =   20
      TabIndex        =   54
      Top             =   1830
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   8
      Left            =   10965
      Max             =   20
      TabIndex        =   53
      Top             =   1830
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   9
      Left            =   11280
      Max             =   20
      TabIndex        =   52
      Top             =   1830
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1050
      Index           =   10
      Left            =   11655
      Max             =   20
      TabIndex        =   51
      Top             =   2070
      Value           =   18
      Width           =   240
   End
   Begin VB.PictureBox E2p0 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2520
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   50
      Top             =   615
      Width           =   190
   End
   Begin VB.PictureBox E2p6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2520
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   49
      Top             =   900
      Width           =   190
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   0
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   1485
      Width           =   1830
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   1
      Left            =   1950
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   1485
      Width           =   1830
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   2
      Left            =   3810
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   1485
      Width           =   1830
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   3
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   1485
      Width           =   1800
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   4
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2085
      Width           =   1830
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   5
      Left            =   1950
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2085
      Width           =   1830
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   6
      Left            =   3810
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2085
      Width           =   1830
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   7
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2085
      Width           =   1800
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   8
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2685
      Width           =   1830
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   9
      Left            =   1950
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2685
      Width           =   1830
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   10
      Left            =   3810
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2685
      Width           =   1830
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   11
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2685
      Width           =   1800
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   12
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3285
      Width           =   1830
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   13
      Left            =   1950
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3285
      Width           =   1830
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   14
      Left            =   3810
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3285
      Width           =   1830
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   15
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3285
      Width           =   1800
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   16
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3885
      Width           =   1830
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   17
      Left            =   1950
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3885
      Width           =   1830
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   18
      Left            =   3810
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3885
      Width           =   1830
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   19
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3885
      Width           =   1800
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   20
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4485
      Width           =   1830
   End
   Begin VB.CommandButton E21 
      Height          =   585
      Index           =   21
      Left            =   1950
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4485
      Width           =   1830
   End
   Begin VB.CommandButton P21 
      Caption         =   "9"
      Height          =   255
      Index           =   8
      Left            =   7140
      TabIndex        =   26
      ToolTipText     =   "Página numero 9"
      Top             =   1065
      Width           =   345
   End
   Begin VB.CommandButton P21 
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   6810
      TabIndex        =   25
      ToolTipText     =   "Página numero 8"
      Top             =   1065
      Width           =   345
   End
   Begin VB.CommandButton P21 
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   6480
      TabIndex        =   24
      ToolTipText     =   "Página numero 7"
      Top             =   1065
      Width           =   345
   End
   Begin VB.CommandButton P21 
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   6150
      TabIndex        =   23
      ToolTipText     =   "Página numero 6"
      Top             =   1065
      Width           =   345
   End
   Begin VB.CommandButton P21 
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   5820
      TabIndex        =   22
      ToolTipText     =   "Página numero 5"
      Top             =   1065
      Width           =   345
   End
   Begin VB.CommandButton P21 
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   5490
      TabIndex        =   21
      ToolTipText     =   "Página numero 4"
      Top             =   1065
      Width           =   345
   End
   Begin VB.CommandButton P21 
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   20
      ToolTipText     =   "Página numero 3"
      Top             =   1065
      Width           =   345
   End
   Begin VB.CommandButton P21 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   4830
      TabIndex        =   19
      ToolTipText     =   "Página numero 2"
      Top             =   1065
      Width           =   345
   End
   Begin VB.CommandButton P21 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   4500
      TabIndex        =   18
      ToolTipText     =   "Página numero 1"
      Top             =   1065
      Width           =   345
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000018&
      Caption         =   "<"
      Height          =   435
      Left            =   7260
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Establecer propiedades de audio"
      Top             =   465
      Width           =   225
   End
   Begin VB.CommandButton SaveCUE 
      Height          =   375
      Left            =   11205
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Guardar Preferencias"
      Top             =   4620
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton OpenCUE 
      Height          =   375
      Left            =   10830
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Abrir Preferencias"
      Top             =   4620
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton E2Import 
      Caption         =   "I"
      Height          =   375
      Left            =   11760
      TabIndex        =   14
      ToolTipText     =   "Importar archivo CUE"
      Top             =   4620
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.CommandButton NewCUE 
      Height          =   375
      Left            =   10455
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Restaurar Preferencias"
      Top             =   4620
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton CmdRestore 
      Caption         =   "R"
      Height          =   375
      Left            =   9045
      TabIndex        =   12
      ToolTipText     =   "Restablecer valores a predeterminado"
      Top             =   4620
      UseMaskColor    =   -1  'True
      Width           =   1005
   End
   Begin VB.CommandButton CmdActualiz 
      Caption         =   "A"
      Height          =   375
      Left            =   8130
      TabIndex        =   11
      ToolTipText     =   "Actualizar datos"
      Top             =   4620
      UseMaskColor    =   -1  'True
      Width           =   915
   End
   Begin VB.CommandButton E2Cue 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Activar / Desactivar CUE"
      Top             =   4605
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.TextBox TxtName 
      Height          =   645
      Left            =   6615
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Text3"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Timer TmOutAuto 
      Left            =   12390
      Top             =   5400
   End
   Begin VB.Timer TmrCUE 
      Left            =   14775
      Top             =   5355
   End
   Begin VB.Timer TMin 
      Left            =   11400
      Top             =   5400
   End
   Begin VB.Timer Tmout 
      Left            =   11895
      Top             =   5400
   End
   Begin VB.Timer TmrScopeLite2 
      Left            =   180
      Top             =   5415
   End
   Begin RM100.TitelBar TitelBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   130
      Top             =   0
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   661
      BackColor       =   8421504
      BackColorCover  =   3
      BackColorV2Begin=   4210752
      BackColorV2End  =   0
      BackColorV1Begin=   4210752
      BackColorV1End  =   0
      ForeColor       =   16777215
      ShowMaximizedEnabled=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " ESTACION 02 - Detenido"
      CaptionPosX     =   1
      BorderNormal    =   2
      BorderColorDarkLight=   12632256
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      ForeColor       =   &H00808000&
      Height          =   510
      Left            =   195
      TabIndex        =   128
      Top             =   615
      Width           =   2220
   End
   Begin VB.Label Pn 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3960
      TabIndex        =   129
      ToolTipText     =   "Numero de Página"
      Top             =   1050
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   90
      Stretch         =   -1  'True
      Top             =   450
      Width           =   3735
   End
   Begin VB.Shape E1Shape 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   3840
      Top             =   4515
      Width           =   3645
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "1k"
      Height          =   195
      Left            =   9840
      TabIndex        =   127
      Top             =   3150
      Width           =   240
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "125hz"
      Height          =   195
      Left            =   8460
      TabIndex        =   126
      Top             =   3150
      Width           =   465
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "16k"
      Height          =   195
      Left            =   11235
      TabIndex        =   125
      Top             =   3150
      Width           =   285
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Rvb"
      Height          =   195
      Left            =   11610
      TabIndex        =   124
      Top             =   3150
      Width           =   330
   End
   Begin VB.Image E2Pic 
      Height          =   4560
      Left            =   7575
      Top             =   480
      Width           =   390
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   8040
      Top             =   4515
      Width           =   7185
   End
   Begin VB.Image Image2 
      Height          =   1470
      Left            =   8070
      Top             =   1845
      Width           =   300
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "FX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   11640
      TabIndex        =   123
      ToolTipText     =   "Numero de Página"
      Top             =   1815
      Width           =   255
   End
   Begin VB.Label LblOutvol 
      Caption         =   "0"
      Height          =   255
      Left            =   10140
      TabIndex        =   9
      Top             =   5850
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LblInvol 
      Caption         =   "100"
      Height          =   255
      Left            =   9690
      TabIndex        =   8
      Top             =   5850
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LblCurrVol 
      Caption         =   "100"
      Height          =   255
      Left            =   9690
      TabIndex        =   7
      Top             =   5640
      Width           =   885
   End
   Begin VB.Label LIndex 
      BackColor       =   &H00C0FFC0&
      Height          =   240
      Left            =   4635
      TabIndex        =   6
      Top             =   5775
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label LblCurrByte 
      Caption         =   "0"
      Height          =   255
      Left            =   8505
      TabIndex        =   4
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label LblEndCue 
      Caption         =   "0"
      Height          =   255
      Left            =   9705
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LblStartCUE 
      Caption         =   "0"
      Height          =   255
      Left            =   8505
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Fi 
      BackColor       =   &H0080FF80&
      Height          =   255
      Left            =   4635
      TabIndex        =   1
      Top             =   5415
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Fn 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1440
      TabIndex        =   0
      Top             =   5415
      Visible         =   0   'False
      Width           =   3135
   End
End
Attribute VB_Name = "Est02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'dimensiones de conversion
Dim ConvertNm           'numeros
Dim ConvertNNm As Long
Dim ConvertTx As String   'textos
Dim ConvertTxT As String
Dim EstNum As Long

Dim FileExt As String

'Dimensiones de Archivos
Dim FileN As String
Dim FileNPath As String
Dim Completo As String
Dim SSTitle As String
Dim FileTP As String

'dimensiones de resultado
Dim Result As String
Dim RResult As String
Dim ResultInfo As Boolean

'Dimensiones de tiempo
Dim TimeNcv As String
Dim PosNcv As String

Private Sub UpdatePos()

Dim ByteLen As String
Dim TimeLen As String
Dim FTime As String
Dim Convt1 As Long

On Error Resume Next
If Est12Control.StopLabel2.Caption = "Stream" Then
    TimeLen = Stream02GetLen(1) 'get len of file in time=1
    FTime = FormatSegs(TimeLen) 'formateamos el tiempo
    E2Pos.Min = 0
    If FTime = 0 Or FTime = "" Then
        E2Pos.Max = FTime + 1
    Else
        E2Pos.Max = FTime
    End If
    If FTime <= 100 Then
        E2Pos.TickFrequency = 1
    Else
        If FTime > 100 And FTime < 200 Then
            E2Pos.TickFrequency = 2
        Else
            If FTime > 200 And FTime < 300 Then
                E2Pos.TickFrequency = 3
            Else
                If FTime > 300 And FTime < 400 Then
                    E2Pos.TickFrequency = 4
                Else
                    E2Pos.TickFrequency = 5
                End If
            End If
        End If
    End If
    E2Pos.Value = 0
    E2Vol.Value = 100
    E2Slide.Value = 0
    LblEnd.Caption = ConvSecToMin(CInt(FTime))
    E2Pos.SmallChange = 10
    E2Pos.LargeChange = 10
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
        ByteLen = Music02GetLen(1)
        Convt1 = CLng(ByteLen)
        Convt1 = Convt1
        E2Pos.Min = 0
        If Convt1 = 0 Then
            E2Pos.Max = Convt1 + 1
        Else
            E2Pos.Max = Convt1
        End If
        E2Pos.TickFrequency = 1
        E2Pos.Value = 0
        E2Pos.ToolTipText = Str$(E2Pos.Value)
        E2Vol.Value = 100
        E2Slide.Value = 0
        LblEnd.Caption = Convt1
        E2Pos.SmallChange = 1
        E2Pos.LargeChange = 1
    Else
        Exit Sub
    End If
End If

End Sub

Public Sub SetAudioLevel(WLeft, WRight)

Dim l, Lft As Integer
Dim R, Rgt As Integer
Dim i As Integer
Static ZMax%, RMax%

On Error Resume Next
WRight = WRight / 12
WLeft = WLeft / 12

'right level meter
'If WRight > 180 Then
'    RMax = (WRight * 24) + 100 'clip
'Else
'    RMax = (WRight * 24)
'End If

'left level meter
'If WLeft > 180 Then
'    ZMax = (WLeft * 24) + 100  'clip
'Else
'    ZMax = (WLeft * 24)
'End If

Lr.Width = WRight 'RMax
Ll.Width = WLeft 'ZMax

End Sub

Private Sub DeployAudioFile(WConNum As Integer)

Dim FileNTest As String
Dim FileExt As String

If XPlorer.File1.filename = "" Or XPlorer.File1.filename = " " Then
    MsgBox "El Archivo o Directorio seleccionado es incorrecto.", vbCritical
    Exit Sub
End If

'.wav, .mp3, .it, .xm
FileExt = StripExtFromFile(XPlorer.File1.filename)
FileN = StripFileFromExt(XPlorer.File1.filename)
FileNPath = Right$(XPlorer.lblPath, Len(XPlorer.lblPath) - 2)
Completo = Right$(XPlorer.lblPath, Len(XPlorer.lblPath) - 2) & "\" & XPlorer.File1.filename

'seleccion de formato de archivo y extraccion de informacion header
Select Case Trim(UCase(FileExt))
    
    'STREAM TYPE WAV-MP1-MP2-MP3-OGG
    Case "WAV", "MP1", "MP2", "MP3", "OGG"
        Est12Data.N2(WConNum).Caption = Completo                  'nombre y path
        Est12Data.c2(WConNum).Caption = FileN                     'nombre solo
        'gets the file len and convert into time
        ConvertTx = FileLoadLen(Completo, "Stream")
        TimeNcv = FormatSegs(ConvertTx)
        Result = ConvSecToMin(CInt(TimeNcv))
        'put the file time into est01
        Est12Data.D2(WConNum).Caption = Result
        E21(WConNum).Caption = FileN    'nombre del archivo
        E21(WConNum).BackColor = &H8000000F
        E21(WConNum).ToolTipText = "Duración: " & Result
        Est12Data.V2(WConNum).Caption = "Stream"
                    
    'MUSIC TYPE XM-MOD-S3M-IT-MTM-MO3-UMX
    Case "XM", "MOD", "S3M", "IT", "MTM", "MO3", "UMX"
        Est12Data.N2(WConNum).Caption = Completo                  'nombre y path
        Est12Data.c2(WConNum).Caption = FileN                     'nombre solo
        Est12Data.D2(WConNum).Caption = ""
        E21(WConNum).Caption = FileN    'nombre del archivo
        E21(WConNum).BackColor = &H8000000F
        E21(WConNum).ToolTipText = ""
        Est12Data.V2(WConNum).Caption = "Music"
        
    'TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TNDTND
    Case "TND"
        MsgBox LoadResString(191), vbInformation, "Radio Maker"
        E21(WConNum).BackColor = &H8000000F

    Case Else
        MsgBox LoadResString(191), vbInformation, "Radio Maker"
        E21(WConNum).BackColor = &H8000000F

End Select

End Sub

Private Sub OpenE2PageFile()

'extraemos los datos necesarios para realizar la operacion
ConvertNm = Pn.Caption
Select Case ConvertNm
    Case "1"
        ConvertNNm = 1
    Case "2"
        ConvertNNm = 2
    Case "3"
        ConvertNNm = 3
    Case "4"
        ConvertNNm = 4
    Case "5"
        ConvertNNm = 5
    Case "6"
        ConvertNNm = 6
    Case "7"
        ConvertNNm = 7
    Case "8"
        ConvertNNm = 8
    Case "9"
        ConvertNNm = 9
End Select

ConvertTx = Trim(Fn.Caption)
If ConvertTx = "" Or ConvertTx = " " Then
    TopMenu.EstCmd.InitDir = App.path & AppEstDir
    TopMenu.EstCmd.Filter = "Archivo de Estacion (*.est)|*.est|Archivos de Estacion"
    TopMenu.EstCmd.DialogTitle = "ESTACION 02 - Abrir archivo"
    TopMenu.EstCmd.FilterIndex = 1
    TopMenu.EstCmd.ShowOpen
    ConvertTx = TopMenu.EstCmd.filename
    EstNum = 2
    Result = OpenEstFile(EstNum, ConvertNNm, ConvertTx)
    If Result = "NotOK" Then
        Exit Sub
    End If
    Fn.Caption = ConvertTx
Else
    EstNum = 2
    Result = OpenEstFile(EstNum, ConvertNNm, ConvertTx)
    If Result = "NotOK" Then
        Exit Sub
    End If
End If

End Sub

Private Sub CmdAutoPan_Click()

Dim PanOrigen As Long
Dim PanRight As Long
Dim PanLeft As Long
Dim ActualPan As Long

PanOrigen = 0
PanLeft = -100
PanRight = 100
ActualPan = E2Slide.Value

While ActualPan < PanRight
    ActualPan = E2Slide.Value + 5   'de o a 100
    E2Slide.Value = ActualPan
Wend
While ActualPan > PanOrigen
    ActualPan = E2Slide.Value - 5   'de 100 a 0
    E2Slide.Value = ActualPan
Wend
While ActualPan > PanLeft
    ActualPan = E2Slide.Value - 5   'de 0 a -100
    E2Slide.Value = ActualPan
Wend
While ActualPan < PanOrigen
    ActualPan = E2Slide.Value + 5   'de -100 a 0
    E2Slide.Value = ActualPan
Wend

End Sub

Private Sub CmdFIN_Click()

TMin.Enabled = True
TMin.Interval = 30

End Sub

Private Sub CmdFOut_Click()

Tmout.Enabled = True
Tmout.Interval = 30

End Sub

Private Sub CmdRestore_Click()

'E2Vol.value = 100
'E2Slide.value = 0
'E2Pos.value = 0
E2Pos.SelStart = 0
E2Pos.SelLength = 0
Text1.text = "00:00:00"
Text2.text = "00:00:00"
LblStartCUE.Caption = 0
LblEndCue.Caption = 0

Dim i As Integer
'restore all EQ presets and reverb
For i = 0 To 10
    fxsc(i).Value = 10
Next i
fxsc(10).Value = 18

End Sub

Private Sub Command1_Click()

If Command1.Caption = "<" Then
    Est02.Width = 15360
    Est02.Left = 0
    Command1.Caption = ">"
    Command1.ToolTipText = "Ocultar AVANZADO"
Else
    ShowWindow "Est02"
    Command1.Caption = "<"
    Command1.ToolTipText = "Mostrar AVANZADO"
End If

End Sub

Private Sub Command10_Click()

Text2.text = LblCurrent.Caption
LblEndCue.Caption = LblCurrByte.Caption
E2Pos.SelLength = E2Pos.Value - E2Pos.SelStart

If LCue.Caption = "Auto" Then
    Call E2Cue_Click
End If

End Sub

Private Sub Command9_Click()

Text1.text = LblCurrent.Caption
LblStartCUE.Caption = LblCurrByte.Caption
E2Pos.SelStart = E2Pos.Value

Text2.SetFocus

End Sub

Private Sub E21_Click(index As Integer)

'desactivamos los CUE que esten activados
If Est01.E1Cue.Caption = "Desactivar CUE" Then
    Est01.E1Cue.Caption = "Activar CUE"
    Est01.E1Cue.BackColor = &H8000000F    'gris
    Est01.TmrCUE.Interval = 0
    Est01.TmrCUE.Enabled = False
End If
If E2Cue.Caption = "Desactivar CUE" Then
    E2Cue.Caption = "Activar CUE"
    E2Cue.BackColor = &H8000000F    'gris
    TmrCUE.Interval = 0
    TmrCUE.Enabled = False
End If

Dim X As Integer
X = index
If E21(X).Caption = "" Or E21(X).Caption = " " Then Exit Sub

'load and play the selected file
Est12Control.Origen2.Caption = "E2"
Result = Estacion02Play(index)
If Result = "NotOk" Then Exit Sub

RestoreDisplay 2     'sets the default display
RestoreAllActiveColor 2 'desactivate all controls
ChangeActiveColor index, 2  'activate the current control

Fi.Caption = index

'gets the config device data
ConfigData = OpenConfigFile

If ConfigData.Aud_Show_FTT = 1 Or ConfigData.Aud_Show_SCOPE = 1 Then
    'activate the level meter
    TmrScopeLite2.Enabled = True
    TmrScopeLite2.Interval = 25
Else
    'deactivate the level meter
    TmrScopeLite2.Interval = 0
    TmrScopeLite2.Enabled = False
End If

'activate the clock timer
TopMenu.ProcTimer.Enabled = True
TopMenu.ProcTimer.Interval = 1
'actualizamos los controles
UpdatePos

'************************************************************
'Automatic open the presets file for the stream selected file
Dim ContNum As Integer
Dim filename As String
Dim LenFN As Long
Dim FileNTest As String
Dim NameFile As String

If Est02.Fi.Caption = "" Then MsgBox "No se selecciono el tema.": Exit Sub

ContNum = CInt(Est02.Fi.Caption)    'extraemos el index del control
filename = Trim(Est12Data.N2(ContNum).Caption)    'extraemos el path y el archivo de audio
NameFile = StripFileFromExt(filename)
filename = Trim(NameFile) & AppCUEFileExt

'abrimos la informacion CUe
OpenCUEFile 2, filename

'starts the fade in/out
If LFin.Caption = "Auto" Then
    If Stream01IsPlaying = True Or Music01IsPlaying = True Then
        If Est12Control.Origen1.Caption = "E1" Then
            Est01.TmoutAuto.Enabled = True
            Est01.TmoutAuto.Interval = 30
        End If
    End If
    If Stream02IsPlaying = True Or Music02IsPlaying = True Then
        If Est12Control.Origen2.Caption = "E2" Then
            E2Vol.Value = 0
            TMin.Enabled = True
            TMin.Interval = 30
        End If
    End If
End If

'chequeamos por el cue auto
If LCue.Caption = "Auto" Then
    Call E2Cue_Click
End If

End Sub

Private Sub E21_DragDrop(index As Integer, Source As Control, X As Single, Y As Single)

DeployAudioFile index   'drag & drop the selected file in xplorer

End Sub

Private Sub E21_DragOver(index As Integer, Source As Control, X As Single, Y As Single, State As Integer)

Select Case State
    Case 0  'drag not finished
        XPlorer.File1.DragIcon = XPlorer.ExCombo.DragIcon
        E21(index).BackColor = &H80FF80    'verde (modificacion)
    Case 1  'finished drag
        XPlorer.File1.DragIcon = XPlorer.tvwDirTree.DragIcon
        E21(index).BackColor = &H8000000F  'gris (normal)
End Select

End Sub

Private Sub E21_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'button 1=left button
'button 2=right button
'button 4=mid button

If Button = 2 Then
    'deploy options menu
    If E21(index).Caption = "" Or E21(index).Caption = " " Then
        Exit Sub
    End If
    'deploy options menu
    TxtName.Visible = True
    TxtName.Top = E21(index).Top
    TxtName.Left = E21(index).Left
    TxtName.Height = E21(index).Height
    TxtName.Width = E21(index).Width
    TxtName.text = E21(index).Caption
    TxtName.SetFocus
    'seteamos el label para saber de que control se trata
    Lindex.Caption = index
Else
    If Button = 4 Then
        'mark control to delete content
        E21(index).BackColor = &HFF&        'rojo
    Else
        'nothing to do
    End If
End If

End Sub

Private Sub E2Cue_Click()

Dim Len1 As Long
Dim Len2 As Long

Len1 = CLng(LblStartCUE.Caption)
Len2 = CLng(LblEndCue.Caption)
If Len1 <= 0 Then
    Exit Sub
Else
    If Len2 <= 0 Then
        Exit Sub
    End If
End If

If E2Cue.Caption = "Activar CUE" Then
    E2Cue.Caption = "Desactivar CUE"
    E2Cue.BackColor = &HFFFF&   'amarillo
    TmrCUE.Enabled = True
    TmrCUE.Interval = 100
Else
    E2Cue.Caption = "Activar CUE"
    E2Cue.BackColor = &H8000000F    'gris
    TmrCUE.Interval = 0
    TmrCUE.Enabled = False
End If

End Sub

Private Sub E2Import_Click()

On Error Resume Next
TopMenu.NTSCmd.InitDir = App.path & AppEstDir
TopMenu.NTSCmd.Filter = "NetShow region (*.txt)|*.txt|NetShow region"
TopMenu.NTSCmd.DialogTitle = "ESTACION 02 - importar archivo"
TopMenu.NTSCmd.CancelError = True
TopMenu.NTSCmd.ShowOpen

If err.Number = 32755 Then Exit Sub

ConvertTx = TopMenu.NTSCmd.filename

Result = GetNetShowAudioRegion(ConvertTx, 2)
If Result = "NotOk" Then
    MsgBox "Ha ocurrido un Error al intentar procesar el archivo especificado.", vbCritical
    Exit Sub
End If

End Sub

Private Sub E2New_Click()

SetDefControl 2   'set the default control´s caption
Pn.Caption = "1"
Fn.Caption = ""

End Sub

Private Sub E2Open_Click()

'extraemos los datos necesarios para realizar la operacion
ConvertNm = Pn.Caption
Select Case ConvertNm
    Case "1"
        ConvertNNm = 1
    Case "2"
        ConvertNNm = 2
    Case "3"
        ConvertNNm = 3
    Case "4"
        ConvertNNm = 4
    Case "5"
        ConvertNNm = 5
    Case "6"
        ConvertNNm = 6
    Case "7"
        ConvertNNm = 7
    Case "8"
        ConvertNNm = 8
    Case "9"
        ConvertNNm = 9
End Select

On Error Resume Next
TopMenu.EstCmd.InitDir = App.path & AppEstDir
TopMenu.EstCmd.Filter = "Archivo de Estación (*.est)|*.est|Archivos de Estación"
TopMenu.EstCmd.DialogTitle = "ESTACION 02 - Abrir archivo de estación"
TopMenu.EstCmd.CancelError = True
TopMenu.EstCmd.ShowOpen

If err.Number = 32755 Then Exit Sub

ConvertTx = TopMenu.EstCmd.filename

EstNum = 2
Result = OpenEstFile(EstNum, ConvertNNm, ConvertTx)
If Result = "NotOK" Then
    Exit Sub
End If
Fn.Caption = ConvertTx

End Sub

Private Sub E2Pause_Click()

If Est12Control.StopLabel2.Caption = "Stream" And Est12Control.Origen2.Caption = "E2" Then
    Stream02Stop   'stream stop
Else
    If Est12Control.StopLabel2.Caption = "Music" And Est12Control.Origen2.Caption = "E2" Then
        Music02Stop    'music stop
    Else
        Exit Sub
    End If
End If

Est02.Caption = "ESTACION 02 - Pausado"

End Sub

Private Sub E2Play_Click()

'desactivamos los CUE que esten activados
If Est01.E1Cue.Caption = "Desactivar CUE" Then
    Est01.E1Cue.Caption = "Activar CUE"
    Est01.E1Cue.BackColor = &H8000000F    'gris
    Est01.TmrCUE.Interval = 0
    Est01.TmrCUE.Enabled = False
End If
If E2Cue.Caption = "Desactivar CUE" Then
    E2Cue.Caption = "Activar CUE"
    E2Cue.BackColor = &H8000000F    'gris
    TmrCUE.Interval = 0
    TmrCUE.Enabled = False
End If

If Est12Control.StopLabel2.Caption = "Stream" And Est12Control.Origen2.Caption = "E2" Then
    If Est02.pcontup.Visible = True Then    'loop enabled?
        Stream02Play (BASS_SAMPLE_LOOP)
    Else
        Stream02Play (0)
    End If
Else
    If Est12Control.StopLabel2.Caption = "Music" And Est12Control.Origen2.Caption = "E2" Then
        Music02Play    'Music play
    Else
        Exit Sub
    End If
End If

Est02.Caption = "ESTACION 02 - Reproduciendo"
RestoreDisplay 2
Est12Control.Origen2.Caption = "E2"
Label1.ForeColor = &HFFFF00

'gets the config device data
ConfigData = OpenConfigFile

If ConfigData.Aud_Show_FTT = 1 Or ConfigData.Aud_Show_SCOPE = 1 Then
    'activate the level meter
    TmrScopeLite2.Enabled = True
    TmrScopeLite2.Interval = 25
Else
    'deactivate the level meter
    TmrScopeLite2.Interval = 0
    TmrScopeLite2.Enabled = False
End If

'activamos el timer de posicion
TopMenu.ProcTimer.Enabled = True
TopMenu.ProcTimer.Interval = 1
'actualizamos los controles
UpdatePos

'starts the fade in/out
If LFin.Caption = "Auto" Then
    If Stream01IsPlaying = True Or Music01IsPlaying = True Then
        If Est12Control.Origen1.Caption = "E1" Then
            Est01.TmoutAuto.Enabled = True
            Est01.TmoutAuto.Interval = 30
        End If
    End If
    If Stream02IsPlaying = True Or Music02IsPlaying = True Then
        If Est12Control.Origen2.Caption = "E2" Then
            E2Vol.Value = 0
            TMin.Enabled = True
            TMin.Interval = 30
        End If
    End If
End If

'chequeamos por el cue auto
If LCue.Caption = "Auto" Then
    Call E2Cue_Click
End If

End Sub

Private Sub E2Pos_Scroll()

Dim Cnv1 As Long

If Est12Control.StopLabel2.Caption = "Stream" And Est12Control.Origen2.Caption = "E2" Then
    Cnv1 = E2Pos.Value
    'change the stream position
    Stream02SetPosition Cnv1, 1
    E2Pos.ToolTipText = ConvSecToMin(CInt(E2Pos.Value))
Else
    If Est12Control.StopLabel2.Caption = "Music" And Est12Control.Origen2.Caption = "E2" Then
        Cnv1 = E2Pos.Value
        'change the music position
        Music02SetPosition Cnv1, 0
        E2Pos.ToolTipText = Str$(E2Pos.Value)
    Else
        Exit Sub
    End If
End If

End Sub

Public Sub E2Save_Click()

'extraemos los datos necesarios para realizar la operacion
ConvertNm = Pn.Caption
Select Case ConvertNm
    Case "1"
        ConvertNNm = 1
    Case "2"
        ConvertNNm = 2
    Case "3"
        ConvertNNm = 3
    Case "4"
        ConvertNNm = 4
    Case "5"
        ConvertNNm = 5
    Case "6"
        ConvertNNm = 6
    Case "7"
        ConvertNNm = 7
    Case "8"
        ConvertNNm = 8
    Case "9"
        ConvertNNm = 9
End Select

ConvertTxT = Trim(Fn.Caption)

On Error Resume Next
If ConvertTxT = "" Or ConvertTxT = " " Then
    TopMenu.EstCmd.InitDir = App.path & AppEstDir
    TopMenu.EstCmd.Filter = "Archivo de Estación (*.est)|*.est|Archivos de Estación"
    TopMenu.EstCmd.DialogTitle = "ESTACION 01 - Guardar archivo de estación"
    TopMenu.EstCmd.FilterIndex = 1
    TopMenu.EstCmd.CancelError = True
    TopMenu.EstCmd.ShowSave

    If err.Number = 32755 Then Exit Sub
    
    ConvertTx = TopMenu.EstCmd.filename

    Fn.Caption = ConvertTx
    EstNum = 2
    Result = SaveEstFile(EstNum, ConvertNNm, ConvertTx)
    If Result = "NotOK" Then
        Exit Sub
    End If
Else
    ConvertTx = Trim(Fn.Caption)
    EstNum = 2
    Result = SaveEstFile(EstNum, ConvertNNm, ConvertTx)
    If Result = "NotOK" Then
        Exit Sub
    End If
End If

End Sub

Private Sub E2Slide_Change()

If Est12Control.StopLabel2.Caption = "Stream" And Est12Control.Origen2.Caption = "E2" Then
    'change the stream pan position
    Stream02SetPan (E2Slide.Value)
    E2Slide.ToolTipText = E2Slide.Value
Else
    If Est12Control.StopLabel2.Caption = "Music" And Est12Control.Origen2.Caption = "E2" Then
        'change the music pan position
        Music02SetPan (E2Slide.Value)
        E2Slide.ToolTipText = E2Slide.Value
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub E2Slide_Scroll()

If Est12Control.StopLabel2.Caption = "Stream" And Est12Control.Origen2.Caption = "E2" Then
    'change the stream pan position
    Stream02SetPan (E2Slide.Value)
    E2Slide.ToolTipText = E2Slide.Value
Else
    If Est12Control.StopLabel2.Caption = "Music" And Est12Control.Origen2.Caption = "E2" Then
        'change the music pan position
        Music02SetPan (E2Slide.Value)
        E2Slide.ToolTipText = E2Slide.Value
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub E2Stop_Click()

'chequeamos por el fade-out automatico
If LFin.Caption = "Auto" Then
    TmoutAuto.Enabled = True
    TmoutAuto.Interval = 30
    Exit Sub
End If

If Est12Control.StopLabel2.Caption = "Stream" And Est12Control.Origen2.Caption = "E2" Then
    Stream02Restart    'stream restart
    Stream02Stop       'stream stop
Else
    If Est12Control.StopLabel2.Caption = "Music" And Est12Control.Origen2.Caption = "E2" Then
        Music02Restart     'music restart
        Music02Stop         'music stop
    Else
        GoSub Cont
    End If
End If

Cont:
'desactivamos el scope
TmrScopeLite2.Interval = 0
TmrScopeLite2.Enabled = False
'reset the displays
Lr.Width = 0
Ll.Width = 0
Picfft2.Cls

'chequeamos el cue auto in est02
If E2Cue.Caption = "Desactivar CUE" Then
    E2Cue.Caption = "Activar CUE"
    E2Cue.BackColor = &H8000000F    'gris
    TmrCUE.Interval = 0
    TmrCUE.Enabled = False
End If

End Sub

Private Sub E2Vol_Change()

If Est12Control.StopLabel2.Caption = "Stream" And Est12Control.Origen2.Caption = "E2" Then
    'change the stream volume
    Stream02SetVolume (E2Vol.Value)
    E2Vol.ToolTipText = E2Vol.Value
    LblCurrVol.Caption = E2Vol.Value
Else
    If Est12Control.StopLabel2.Caption = "Music" And Est12Control.Origen2.Caption = "E2" Then
        'change the music volume
        Music02SetVolume (E2Vol.Value)
        E2Vol.ToolTipText = E2Vol.Value
        LblCurrVol.Caption = E2Vol.Value
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub E2Vol_Scroll()

If Est12Control.StopLabel2.Caption = "Stream" And Est12Control.Origen2.Caption = "E2" Then
    'change the stream volume
    Stream02SetVolume (E2Vol.Value)
    E2Vol.ToolTipText = E2Vol.Value
    LblCurrVol.Caption = E2Vol.Value
Else
    If Est12Control.StopLabel2.Caption = "Music" And Est12Control.Origen2.Caption = "E2" Then
        'change the music volume
        Music02SetVolume (E2Vol.Value)
        E2Vol.ToolTipText = E2Vol.Value
        LblCurrVol.Caption = E2Vol.Value
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub fft10_Click()

If fft10.ForeColor = &H808000 Then   'verde oscuro
    fft10.ForeColor = &HFFFF00   'verde claro
    fft4.ForeColor = &H808000
    fft6.ForeColor = &H808000
    fft8.ForeColor = &H808000
    fft2.ForeColor = &H808000
End If

End Sub

Private Sub fft2_Click()

If fft2.ForeColor = &H808000 Then   'verde oscuro
    fft2.ForeColor = &HFFFF00   'verde claro
    fft4.ForeColor = &H808000
    fft6.ForeColor = &H808000
    fft8.ForeColor = &H808000
    fft10.ForeColor = &H808000
End If

End Sub

Private Sub fft4_Click()

If fft4.ForeColor = &H808000 Then   'verde oscuro
    fft4.ForeColor = &HFFFF00   'verde claro
    fft2.ForeColor = &H808000
    fft6.ForeColor = &H808000
    fft8.ForeColor = &H808000
    fft10.ForeColor = &H808000
End If

End Sub

Private Sub fft6_Click()

If fft6.ForeColor = &H808000 Then   'verde oscuro
    fft6.ForeColor = &HFFFF00   'verde claro
    fft4.ForeColor = &H808000
    fft2.ForeColor = &H808000
    fft8.ForeColor = &H808000
    fft10.ForeColor = &H808000
End If

End Sub

Private Sub fft8_Click()

If fft8.ForeColor = &H808000 Then   'verde oscuro
    fft8.ForeColor = &HFFFF00   'verde claro
    fft4.ForeColor = &H808000
    fft6.ForeColor = &H808000
    fft2.ForeColor = &H808000
    fft10.ForeColor = &H808000
End If

End Sub

Private Sub Form_Load()

'*** load the commands strings
E2Cue.Caption = LoadResString(2007)
E2Import.Caption = LoadResString(2006)
CmdRestore.Caption = LoadResString(2005)
CmdActualiz.Caption = LoadResString(2004)

'*** load commands pictures
    E2Pic.Picture = LoadResPicture("EST_02", 0)
    'load led1
    Picture1.Picture = LoadResPicture("BACK_LED", 0)
    Ll.Picture = LoadResPicture("FRONT_LED", 0)
    'load led2
    Picture2.Picture = LoadResPicture("BACK_LED", 0)
    Lr.Picture = LoadResPicture("FRONT_LED", 0)
    'Load control pictures
    E2Play.Picture = LoadResPicture("R_PLAY", 0)
    E2Pause.Picture = LoadResPicture("R_PAUSE", 0)
    E2Stop.Picture = LoadResPicture("R_STOP", 0)
    '--- more...
    E2New.Picture = LoadResPicture("ICO_NEW", 0)
    E2Open.Picture = LoadResPicture("ICO_OPEN", 0)
    E2Save.Picture = LoadResPicture("ICO_SAVE", 0)
    '--- and more...
    NewCUE.Picture = LoadResPicture("ICO_NEW", 0)
    OpenCUE.Picture = LoadResPicture("ICO_OPEN", 0)
    SaveCUE.Picture = LoadResPicture("ICO_SAVE", 0)
    '--- and much mooooore....
    Image1.Picture = LoadResPicture("EST_PANEL", 0)
    Picture3.Picture = LoadResPicture("EST_PANEL_FFT", 0)

    'reset the size
    Lr.Width = 0
    Ll.Width = 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
HideWindow "Est02"

End Sub

Private Sub Form_Terminate()

HideWindow "Est02"

End Sub

Private Sub Form_Unload(Cancel As Integer)

HideWindow "Est02"

End Sub

Private Sub fxsc_Change(index As Integer)

UpdateFX02 (index)

End Sub

Private Sub fxsc_Scroll(index As Integer)

UpdateFX02 (index)

End Sub

Private Sub LAplay_Click()

If LAplay.ForeColor = &H808000 Then
    LAplay.ForeColor = &HFFFF00 'claro
Else
    LAplay.ForeColor = &H808000 'oscuro
End If

End Sub

Private Sub LCue_Click()

If LCue.Caption = "Man" Then
    LCue.Caption = "Auto"
Else
    LCue.Caption = "Man"
End If

End Sub

Private Sub Lfft_Click()

If Lfft.ForeColor = &H808000 Then 'verde oscuro
    Lfft.ForeColor = &HFFFF00   'verde claro
    Lspc.ForeColor = &H808000   'verde oscuro
    Lspcz.ForeColor = &H808000   'verde oscuro
    Lspcd.ForeColor = &H808000   'verde oscuro
    Lspcb.ForeColor = &H808000   'verde oscuro
End If

End Sub

Private Sub LFin_Click()

If LFin.Caption = "Man" Then
    LFin.Caption = "Auto"
    CmdFIN.Enabled = False
    CmdFOut.Enabled = False
Else
    LFin.Caption = "Man"
    CmdFIN.Enabled = True
    CmdFOut.Enabled = True
End If

End Sub

Private Sub Lspc_Click()

If Lspc.ForeColor = &H808000 Then 'verde oscuro
    Lspc.ForeColor = &HFFFF00   'verde claro
    Lspcb.ForeColor = &HFFFF00   'verde claro
    Lspcz.ForeColor = &H808000   'verde oscuro
    Lspcd.ForeColor = &H808000   'verde oscuro
    Lfft.ForeColor = &H808000   'verde oscuro
End If

End Sub

Private Sub Lspcb_Click()

If Lspc.ForeColor = &HFFFF00 Then 'verde claro
    Lspc.ForeColor = &HFFFF00   'verde claro
    Lspcb.ForeColor = &HFFFF00   'verde claro
    Lspcz.ForeColor = &H808000   'verde oscuro
    Lspcd.ForeColor = &H808000   'verde oscuro
    Lfft.ForeColor = &H808000   'verde oscuro
End If

End Sub

Private Sub Lspcd_Click()

If Lspc.ForeColor = &HFFFF00 Then 'verde claro
    Lspc.ForeColor = &HFFFF00   'verde claro
    Lspcd.ForeColor = &HFFFF00   'verde claro
    Lspcz.ForeColor = &H808000   'verde oscuro
    Lspcb.ForeColor = &H808000   'verde oscuro
    Lfft.ForeColor = &H808000   'verde oscuro
End If

End Sub

Private Sub Lspcz_Click()

If Lspc.ForeColor = &HFFFF00 Then 'verde claro
    Lspc.ForeColor = &HFFFF00   'verde claro
    Lspcz.ForeColor = &HFFFF00   'verde claro
    Lspcb.ForeColor = &H808000   'verde oscuro
    Lspcd.ForeColor = &H808000   'verde oscuro
    Lfft.ForeColor = &H808000   'verde oscuro
End If

End Sub

Private Sub NewCUE_Click()

CmdRestore_Click

End Sub

Private Sub OpenCUE_Click()

Dim ContNum As Integer
Dim filename As String
Dim LenFN As Long
Dim FileNTest As String
Dim NameFile As String

If Est02.Fi.Caption = "" Then MsgBox "Primero deberá cargar o seleccionar un tema.": Exit Sub

ContNum = CInt(Est02.Fi.Caption)    'extraemos el index del control
filename = Trim(Est12Data.N2(ContNum).Caption)    'extraemos el path y el archivo de audio
NameFile = StripFileFromExt(filename)
filename = Trim(NameFile) & AppCUEFileExt

'abrimos la informacion CUe
OpenCUEFile 2, filename

End Sub

Private Sub P21_Click(index As Integer)

If TxtName.Visible = True Then
    TxtName.Visible = False
End If

ConvertTxT = Trim(Fn.Caption)
If ConvertTxT = "" Or ConvertTxT = " " Then Exit Sub

Select Case index
    Case 0
        Call E2Save_Click   'save the old page file
        SetDefControl 2   'set the default control´s caption
        Pn.Caption = "1"
        Call OpenE2PageFile 'Open the new page file
    Case 1
        Call E2Save_Click
        SetDefControl 2
        Pn.Caption = "2"
        Call OpenE2PageFile
    Case 2
        Call E2Save_Click
        SetDefControl 2
        Pn.Caption = "3"
        Call OpenE2PageFile
    Case 3
        Call E2Save_Click
        SetDefControl 2
        Pn.Caption = "4"
        Call OpenE2PageFile
    Case 4
        Call E2Save_Click
        SetDefControl 2
        Pn.Caption = "5"
        Call OpenE2PageFile
    Case 5
        Call E2Save_Click
        SetDefControl 2
        Pn.Caption = "6"
        Call OpenE2PageFile
    Case 6
        Call E2Save_Click
        SetDefControl 2
        Pn.Caption = "7"
        Call OpenE2PageFile
    Case 7
        Call E2Save_Click
        SetDefControl 2
        Pn.Caption = "8"
        Call OpenE2PageFile
    Case 8
        Call E2Save_Click
        SetDefControl 2
        Pn.Caption = "9"
        Call OpenE2PageFile
End Select

End Sub

Private Sub pcontdw_Click()

pcontdw.Visible = False
pcontup.Visible = True

End Sub

Private Sub pcontup_Click()

pcontdw.Visible = True
pcontup.Visible = False

End Sub

Private Sub SaveCUE_Click()

Dim ContNum As Integer
Dim filename As String
Dim LenFN As Long
Dim FileNTest As String
Dim NameFile As String

If Est02.Fi.Caption = "" Then MsgBox "Primero deberá cargar o seleccionar un tema.": Exit Sub

ContNum = CInt(Est02.Fi.Caption)    'extraemos el index del control
filename = Trim(Est12Data.N2(ContNum).Caption)    'extraemos el path y el archivo de audio
NameFile = StripFileFromExt(filename)
filename = Trim(NameFile) & AppCUEFileExt

'guardamos la informacion CUe
SaveCUEFile 2, filename

End Sub

Private Sub TMin_Timer()

If E2Vol.Value = 100 Or E2Vol.Value = CLng(LblInvol.Caption) Then
    TMin.Interval = 0
    TMin.Enabled = False
Else
    E2Vol.Value = E2Vol.Value + 1
End If

End Sub

Private Sub Tmout_Timer()

If E2Vol.Value = 0 Or E2Vol.Value = CLng(LblOutvol.Caption) Then
    Tmout.Interval = 0
    Tmout.Enabled = False
Else
    E2Vol.Value = E2Vol.Value - 1
End If

End Sub

Private Sub TmOutAuto_Timer()

If E2Vol.Value = 0 Then
    If Est12Control.StopLabel2.Caption = "Stream" And Est12Control.Origen2.Caption = "E2" Then
        Stream02Restart    'stream restart
        Stream02Stop       'stream stop
    Else
        If Est12Control.StopLabel2.Caption = "Music" And Est12Control.Origen2.Caption = "E2" Then
            Music02Restart     'music restart
            Music02Stop         'music stop
        Else
            'desactivamos el scope
            TmrScopeLite2.Interval = 0
            TmrScopeLite2.Enabled = False
            'reset the displays
            Lr.Width = 0
            Ll.Width = 0
            Picfft2.Cls
            TmoutAuto.Interval = 0
            TmoutAuto.Enabled = False
        End If
    End If
    'desactivamos el scope
    TmrScopeLite2.Interval = 0
    TmrScopeLite2.Enabled = False
    'reset the displays
    Lr.Width = 0
    Ll.Width = 0
    Picfft2.Cls
    TmoutAuto.Interval = 0
    TmoutAuto.Enabled = False
Else
    E2Vol.Value = E2Vol.Value - 1
End If

End Sub

Private Sub TmrCUE_Timer()

    Dim StartByte As Long
    Dim EndByte As Long
    Dim ActualByte As Long
    
    StartByte = CLng(LblStartCUE.Caption)
    EndByte = CLng(LblEndCue.Caption)
    ActualByte = CLng(LblCurrByte.Caption)
    
    'calculations
    EndByte = (EndByte / 6000) / 3
    ActualByte = (ActualByte / 6000) / 3
    
    'change the stream position for a cue start
If Est12Control.StopLabel2.Caption = "Stream" Then
    Do While ActualByte >= EndByte
        Stream02SetPosition StartByte, 2
        Exit Do
    Loop
    E2Pos.ToolTipText = ConvSecToMin(CInt(E2Pos.Value))
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
'        'Cnv1 = E2Pos.Value
'        'change the music position
'        'Music02SetPosition Cnv1, 0
'        'E2Pos.ToolTipText = Str$(E2Pos.Value)
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub TmrScopeLite2_Timer()

'funciona tanto como para la Estacion02 como para la Tanda02
Dim LLft
Dim RRgt
Dim SType As String

If Est12Control.StopLabel2 = "Stream" Then
    If Est12Control.Origen2.Caption = "E2" Then
        LLft = Stream02GetLEFTLevel
        RRgt = Stream02GetRIGHTLevel
        Est02.SetAudioLevel LLft, RRgt
        SType = "Stream"
    End If
End If

If Est12Control.StopLabel2 = "Music" Then
    If Est12Control.Origen2.Caption = "E2" Then
        LLft = Music02GetLEFTLevel
        RRgt = Music02GetRIGHTLevel
        Est02.SetAudioLevel LLft, RRgt
        SType = "Music"
    End If
End If

'chequeamos por el tipo de display en est01
If Lfft.ForeColor = &HFFFF00 Then 'verde claro
    If fft2.ForeColor = &HFFFF00 Then   'verde claro
        Call DrawFFT(2, SType, 2) 'fft spectrum display
    Else
        If fft4.ForeColor = &HFFFF00 Then   'verde claro
            Call DrawFFT(2, SType, 4) 'fft spectrum display
        Else
            If fft6.ForeColor = &HFFFF00 Then   'verde claro
                Call DrawFFT(2, SType, 6) 'fft spectrum display
            Else
                If fft8.ForeColor = &HFFFF00 Then   'verde claro
                    Call DrawFFT(2, SType, 8) 'fft spectrum display
                Else
                    If fft10.ForeColor = &HFFFF00 Then   'verde claro
                        Call DrawFFT(2, SType, 10) 'fft spectrum display
                    Else
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
Else
    If Lspc.ForeColor = &HFFFF00 Then
        If Lspcz.ForeColor = &HFFFF00 Then  'scope izquiero
            Call DrawScope(&HFFFF00, &H808000, 5, 0, 130, 50, 2, SType, ScopeSideBySide)
        End If
        If Lspcd.ForeColor = &HFFFF00 Then  'scope derecho
            Call DrawScope(&H808000, &HFFFF00, 5, 0, 130, 50, 2, SType, ScopeSideBySide)
        End If
        If Lspcb.ForeColor = &HFFFF00 Then  'scope dual
            Call DrawScope(&HFFFF00, &HFFFF00, 5, 0, 130, 50, 2, SType, ScopeDouble)
        End If
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)

Dim IDX As Integer

If KeyAscii = 13 Then   'ENTER
    IDX = CInt(Lindex.Caption)
    E21(IDX).Caption = TxtName.text
    Est12Data.c2(IDX).Caption = TxtName.text
    TxtName.Visible = False
End If
If KeyAscii = 27 Then   'ESCAPE
    TxtName.Visible = False
End If

End Sub
