VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form Tanda01 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TANDA - Detenido"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7530
   ControlBox      =   0   'False
   FillColor       =   &H00808080&
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox T1View 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4485
      Left            =   120
      ScaleHeight     =   4425
      ScaleWidth      =   7245
      TabIndex        =   76
      Top             =   960
      Width           =   7305
   End
   Begin ComctlLib.Slider T1Vol 
      Height          =   255
      Left            =   1590
      TabIndex        =   74
      Top             =   8280
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   450
      _Version        =   327682
   End
   Begin ComctlLib.ProgressBar Prbar1 
      Height          =   285
      Left            =   120
      TabIndex        =   73
      Top             =   5490
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   6870
      Max             =   0
      Min             =   10
      TabIndex        =   68
      Top             =   150
      Value           =   3
      Width           =   135
   End
   Begin VB.CommandButton CmdBlock 
      Height          =   375
      Left            =   2220
      Picture         =   "Tanda01.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "Agregar / Eliminar / modificar bloques"
      Top             =   6525
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.PictureBox T1F8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   7230
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   64
      TabStop         =   0   'False
      ToolTipText     =   "Hora de FINALIZACION de la Tanda en reproducci�n"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1F7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   7050
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   63
      TabStop         =   0   'False
      ToolTipText     =   "Hora de FINALIZACION de la Tanda en reproducci�n"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1F6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6870
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   62
      TabStop         =   0   'False
      ToolTipText     =   "Hora de FINALIZACION de la Tanda en reproducci�n"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1F5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6690
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   61
      TabStop         =   0   'False
      ToolTipText     =   "Hora de FINALIZACION de la Tanda en reproducci�n"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1F4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6495
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   60
      TabStop         =   0   'False
      ToolTipText     =   "Hora de FINALIZACION de la Tanda en reproducci�n"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1F3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6300
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   59
      TabStop         =   0   'False
      ToolTipText     =   "Hora de FINALIZACION de la Tanda en reproducci�n"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1F2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6120
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   58
      TabStop         =   0   'False
      ToolTipText     =   "Hora de FINALIZACION de la Tanda en reproducci�n"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1F1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   5925
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   57
      TabStop         =   0   'False
      ToolTipText     =   "Hora de FINALIZACION de la Tanda en reproducci�n"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1I8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4620
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   55
      TabStop         =   0   'False
      ToolTipText     =   "Hora de INICIO de reproducci�n"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1I7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4440
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   54
      TabStop         =   0   'False
      ToolTipText     =   "Hora de INICIO de reproducci�n"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1I6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4260
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   53
      TabStop         =   0   'False
      ToolTipText     =   "Hora de INICIO de reproducci�n"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1I5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4080
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Hora de INICIO de reproducci�n"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1I4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3885
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   51
      TabStop         =   0   'False
      ToolTipText     =   "Hora de INICIO de reproducci�n"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1I3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3690
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "Hora de INICIO de reproducci�n"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1I2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3510
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   49
      TabStop         =   0   'False
      ToolTipText     =   "Hora de INICIO de reproducci�n"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1I1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3315
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "Hora de INICIO de reproducci�n"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1t1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   795
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   46
      TabStop         =   0   'False
      ToolTipText     =   "Duraci�n TOTAL de la Tanda"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1t2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   990
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "Duraci�n TOTAL de la Tanda"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1t3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1170
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   "Duraci�n TOTAL de la Tanda"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1t4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1365
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   43
      TabStop         =   0   'False
      ToolTipText     =   "Duraci�n TOTAL de la Tanda"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1t5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1560
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   42
      TabStop         =   0   'False
      ToolTipText     =   "Duraci�n TOTAL de la Tanda"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1t6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1740
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "Duraci�n TOTAL de la Tanda"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1t7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1920
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "Duraci�n TOTAL de la Tanda"
      Top             =   5940
      Width           =   190
   End
   Begin VB.PictureBox T1t8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2100
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "Duraci�n TOTAL de la Tanda"
      Top             =   5940
      Width           =   190
   End
   Begin VB.CommandButton T1OrderA 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Reordenar / actualizar tiempo desde el tema seleccionado hacia el final de la lista"
      Top             =   6525
      UseMaskColor    =   -1  'True
      Width           =   800
   End
   Begin VB.CommandButton T1Order 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Reordenar / actualizar tiempo desde el comienzo de la Lista"
      Top             =   6525
      UseMaskColor    =   -1  'True
      Width           =   800
   End
   Begin VB.Timer ClockTimer 
      Left            =   2190
      Top             =   9585
   End
   Begin VB.CommandButton T1Stop 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Detener"
      Top             =   6525
      UseMaskColor    =   -1  'True
      Width           =   500
   End
   Begin VB.CommandButton T1Play 
      Enabled         =   0   'False
      Height          =   375
      Left            =   930
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Reproducir seleccionado"
      Top             =   6525
      UseMaskColor    =   -1  'True
      Width           =   700
   End
   Begin VB.CommandButton T1Next 
      Enabled         =   0   'False
      Height          =   375
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Reproducir continuo"
      Top             =   6525
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.PictureBox T1p6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3330
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   450
      Width           =   190
   End
   Begin VB.PictureBox T1p0 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3330
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   165
      Width           =   190
   End
   Begin VB.Timer SyncTimer 
      Left            =   2190
      Top             =   9135
   End
   Begin VB.TextBox Intr 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   6600
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   26
      Text            =   "2"
      Top             =   180
      Width           =   405
   End
   Begin VB.CommandButton T1Del 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3765
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Eliminar"
      Top             =   6525
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton T1Down 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3405
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Bajar"
      Top             =   6525
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton T1Up 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3045
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Subir"
      Top             =   6525
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Timer TmOut2 
      Left            =   5925
      Top             =   8505
   End
   Begin VB.Timer TmIn2 
      Left            =   5430
      Top             =   8505
   End
   Begin VB.Timer TmOut1 
      Left            =   5925
      Top             =   8010
   End
   Begin VB.Timer TmIn1 
      Left            =   5430
      Top             =   8010
   End
   Begin VB.CommandButton T1Save 
      Height          =   375
      Left            =   7065
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Guardar Tanda"
      Top             =   6525
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton T1Open 
      Height          =   375
      Left            =   6705
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Abrir Tanda"
      Top             =   6525
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton T1New 
      Height          =   375
      Left            =   6345
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Nueva Tanda"
      Top             =   6525
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton T1Prop 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Propiedades de audio"
      Top             =   6525
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.PictureBox T1p7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3540
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   450
      Width           =   190
   End
   Begin VB.PictureBox T1p8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3735
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   450
      Width           =   190
   End
   Begin VB.PictureBox T1p9 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3915
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   450
      Width           =   190
   End
   Begin VB.PictureBox T1p10 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4110
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   450
      Width           =   190
   End
   Begin VB.PictureBox T1p11 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4305
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   450
      Width           =   190
   End
   Begin VB.PictureBox T1p1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3540
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   165
      Width           =   190
   End
   Begin VB.PictureBox T1p2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3735
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   165
      Width           =   190
   End
   Begin VB.PictureBox T1p3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3915
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   165
      Width           =   190
   End
   Begin VB.PictureBox T1p4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4110
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   165
      Width           =   190
   End
   Begin VB.PictureBox T1p5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4305
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   165
      Width           =   190
   End
   Begin ComctlLib.Slider T2Vol 
      Height          =   255
      Left            =   1590
      TabIndex        =   75
      Top             =   8670
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   450
      _Version        =   327682
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   780
      Top             =   8370
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Label BlkFn 
      BackColor       =   &H00FFFF00&
      Height          =   195
      Left            =   2190
      TabIndex        =   72
      Top             =   7890
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "no definido.blk"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   5520
      TabIndex        =   71
      Top             =   480
      Width           =   1275
   End
   Begin VB.Label LBlk 
      BackStyle       =   0  'Transparent
      Caption         =   "/ Man"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   6870
      TabIndex        =   70
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bloque:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   4860
      TabIndex        =   69
      Top             =   480
      Width           =   585
   End
   Begin VB.Label FTime 
      BackColor       =   &H00FF8080&
      Caption         =   "---"
      Height          =   240
      Left            =   3090
      TabIndex        =   66
      Top             =   7380
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "FIN:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   5580
      TabIndex        =   65
      Top             =   5940
      Width           =   315
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "INICIO:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2745
      TabIndex        =   56
      Top             =   5940
      Width           =   555
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   150
      TabIndex        =   47
      Top             =   5940
      Width           =   585
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   90
      TabIndex        =   38
      Top             =   5880
      Width           =   7395
   End
   Begin VB.Label SyncStream 
      BackColor       =   &H000080FF&
      Height          =   240
      Left            =   2685
      TabIndex        =   33
      Top             =   9225
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label SyncLabel 
      BackColor       =   &H000040C0&
      Caption         =   "0"
      Height          =   240
      Left            =   4620
      TabIndex        =   32
      Top             =   9225
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Ltime 
      BackColor       =   &H00FF8080&
      Caption         =   "00:00:00"
      Height          =   240
      Left            =   2190
      TabIndex        =   31
      Top             =   7380
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Dev-2:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   180
      TabIndex        =   30
      Top             =   450
      Width           =   510
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dev-1:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   180
      TabIndex        =   29
      Top             =   180
      Width           =   510
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "segs."
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   7035
      TabIndex        =   28
      Top             =   180
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Inter:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   6135
      TabIndex        =   27
      Top             =   180
      Width           =   420
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "F-In/Out:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   4860
      TabIndex        =   25
      Top             =   180
      Width           =   690
   End
   Begin VB.Label LFin 
      BackStyle       =   0  'Transparent
      Caption         =   "Auto"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   5625
      TabIndex        =   24
      Top             =   180
      Width           =   375
   End
   Begin VB.Label LKey 
      BackColor       =   &H0080FF80&
      Caption         =   "1"
      Height          =   240
      Left            =   3870
      TabIndex        =   23
      Top             =   7650
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Fn 
      BackColor       =   &H00008000&
      Height          =   210
      Left            =   2190
      TabIndex        =   22
      Top             =   7650
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Shape T1Shape 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   90
      Top             =   6465
      Width           =   7410
   End
   Begin VB.Label T2Name 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   675
      TabIndex        =   21
      Top             =   450
      Width           =   2580
   End
   Begin VB.Label T1Name 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   675
      TabIndex        =   20
      Top             =   180
      Width           =   2580
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   4770
      Picture         =   "Tanda01.frx":0102
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2700
   End
   Begin VB.Image Image2 
      Height          =   885
      Left            =   45
      Picture         =   "Tanda01.frx":1972
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4635
   End
   Begin VB.Menu BlockMnu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu Blockmenu_Comandos 
         Caption         =   "&Insertar comandos"
         Begin VB.Menu Blockmenu_Comandos_HM 
            Caption         =   "Reproducir Hora y minutos"
            Enabled         =   0   'False
         End
         Begin VB.Menu Blockmenu_Comandos_TH 
            Caption         =   "Reproducir Temperatura y humedad"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu BlockMnu_sep0 
         Caption         =   "-"
      End
      Begin VB.Menu BlockMnu_define 
         Caption         =   "&Definir bloque de utilizaci�n..."
         Enabled         =   0   'False
      End
      Begin VB.Menu BlockMnu_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu BlockMnu_insert 
         Caption         =   "&Insertar bloque..."
         Enabled         =   0   'False
      End
      Begin VB.Menu BlockMnu_delete 
         Caption         =   "&Eliminar bloque"
         Enabled         =   0   'False
      End
      Begin VB.Menu BlockMnu_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu BlockMnu_config 
         Caption         =   "&Configurar bloques..."
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "Tanda01"
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

'dimensiones de listitem
Dim ItmZ As ListItem
Dim TxtKey As String
Dim NewKey As String
Dim ANum As Integer
Dim ONum As Integer
Dim NNum As Long
Dim nIndex As Integer

'dimensiones Tanda time check
Dim LzTime As String
Dim AcTime As String
Dim LzTh1, LzTm1, LzTs1 As Integer
Dim AcTh1, AcTm1, AcTs1 As Integer

Sub DeployTandaFile()

Dim OTime As String
Dim NTime As String
Dim RTime As String
Dim Time1 As Double
Dim Time2 As Double
Dim TMint As Double
Dim Resultado As Double

If XPlorer.File1.filename = "" Or XPlorer.File1.filename = " " Then
    MsgBox LoadResString(137), vbCritical
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
        ONum = T1View.ListItems.count
        NNum = ONum + 1
        TxtKey = "r"
        NewKey = TxtKey & NNum
        
        Set ItmZ = T1View.ListItems.Add(NNum, NewKey, Completo) 'path & file
        ItmZ.SubItems(1) = "Stream"     'file type
        ItmZ.SubItems(2) = FileN        'file name
        'gets the file len and convert into time
        ConvertTx = FileLoadLen(Completo, "Stream")
        TimeNcv = FormatSegs(ConvertTx)
        Result = ConvSecToMin(CInt(TimeNcv))
        'refresh the time display
            OTime = Trim(Tanda01.Ltime.Caption)
            NTime = Trim(Result)
            Time1 = ConvMinToSec(OTime)
            Time2 = ConvMinToSec(NTime)
            'tiempo de mixado intermedio
            TMint = CDbl(Trim(Tanda01.Intr.text))
            Resultado = Time1 + Time2
            Resultado = (Resultado - TMint) + 1
            RTime = ConvSecToMin(Resultado)
            SetSumTime RTime, 1
            Tanda01.Ltime.Caption = RTime
        'put the rest of info
        ItmZ.SubItems(3) = Result      'duracion del tema
        ItmZ.SubItems(4) = "00:00:00"  'poner aqui la hora de lanzamiento
        ItmZ.SubItems(5) = "-----"     'poner aqui el path & file del mixado
        ItmZ.SubItems(6) = "-----"     'poner aqui el type del mixado
        ItmZ.SubItems(7) = "-----"     'poner aqui el nombre del mixado interm.
        ItmZ.SubItems(8) = "00:00:00"     'poner aqui la duracion del mixado
        ItmZ.SubItems(9) = "00:00"  'poner aqui la hora de lanzam. del mixado
        'Completo                  'nombre y path
        'FileN                     'nombre solo
            
    'MUSIC TYPE XM-MOD-S3M-IT-MTM-MO3-UMX
    Case "XM", "MOD", "S3M", "IT", "MTM", "MO3", "UMX"
        MsgBox LoadResString(184), vbInformation, "Radio Maker"
        
    'TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TNDTND
    Case "tnd", "Tnd", "tNd", "tnD", "TNd", "TND", "tND"
        MsgBox LoadResString(184), vbInformation, "Radio Maker"
        
    Case Else
        MsgBox LoadResString(184), vbInformation, "Radio Maker"

End Select

End Sub

Private Sub Blockmenu_Comandos_HM_Click()

    ONum = T1View.ListItems.count
    NNum = ONum + 1
    TxtKey = "r"
    NewKey = TxtKey & NNum
    
    Completo = "-----"
    Set ItmZ = T1View.ListItems.Add(NNum, NewKey, Completo) 'path & file

    'Set ItmX = T1View.ListItems.Item.ForeColor = &H8000000D
    ItmZ.SubItems(1) = "Command"     'file type
    ItmZ.SubItems(2) = ">>>>>> Reproducir Hora"        'file name
    'gets the file len and convert into time
    'put the rest of info
    ItmZ.SubItems(3) = "00:00:00"  'duracion del tema
    ItmZ.SubItems(4) = "00:00:00"  'poner aqui la hora de lanzamiento
    ItmZ.SubItems(5) = "-----"     'poner aqui el path & file del mixado
    ItmZ.SubItems(6) = "-----"     'poner aqui el type del mixado
    ItmZ.SubItems(7) = "-----"     'poner aqui el nombre del mixado interm.
    ItmZ.SubItems(8) = "00:00:00"     'poner aqui la duracion del mixado
    ItmZ.SubItems(9) = "00:00"  'poner aqui la hora de lanzam. del mixado
    'T1View.SelectedItem.ForeColor = &H8000000D

End Sub


Private Sub BlockMnu_config_Click()

'/// Configurar los bloques publicitarios.

FrmBlock.Show

End Sub

Private Sub BlockMnu_define_Click()

'/// display the Open dialog box
On Error Resume Next
TopMenu.BlockCmd.InitDir = App.path & AppBlockDir
TopMenu.BlockCmd.Filter = "Archivo de Bloque (*.blk)|*.blk|Archivo de Bloque"
TopMenu.BlockCmd.DialogTitle = "Bloques de publicidad - Abrir archivo de bloque."
TopMenu.BlockCmd.CancelError = True
TopMenu.BlockCmd.ShowOpen

If err.Number = 32755 Then Exit Sub

BlkFn.Caption = TopMenu.BlockCmd.filename
Label10.Caption = StripFileFromDir(BlkFn.Caption)

End Sub

Private Sub BlockMnu_delete_Click()

'/// eliminar un bloque publicitario.

End Sub

Private Sub BlockMnu_insert_Click()

'/// Insertar un bloque publicitario.

End Sub

Private Sub CmdBlock_Click()

'display the block menu
PopupMenu BlockMnu

End Sub

Private Sub Form_Load()

'*** load commands pictures
    T1Next.Picture = LoadResPicture("R_NEXT", 0)
    T1Play.Picture = LoadResPicture("R_PLAY", 0)
    T1Stop.Picture = LoadResPicture("R_STOP", 0)
    T1Up.Picture = LoadResPicture("ICO_UP", 0)
    T1Down.Picture = LoadResPicture("ICO_DOWN", 0)
    T1Del.Picture = LoadResPicture("ICO_DELETE", 0)
    T1Prop.Picture = LoadResPicture("ICO_PROP", 0)
    T1Order.Picture = LoadResPicture("R_SYNC_ALL", 0)
    T1OrderA.Picture = LoadResPicture("R_SYNC_SELECTED", 0)
    T1New.Picture = LoadResPicture("ICO_NEW", 0)
    T1Open.Picture = LoadResPicture("ICO_OPEN", 0)
    T1Save.Picture = LoadResPicture("ICO_SAVE", 0)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

HideWindow "Tnd01"

End Sub

Private Sub Form_Terminate()

HideWindow "Tnd01"

End Sub

Private Sub Form_Unload(Cancel As Integer)

HideWindow "Tnd01"

End Sub

Private Sub Label10_Click()

MsgBox "Opci�n no implementada.", vbInformation
Exit Sub

Call BlockMnu_define_Click

End Sub

Private Sub LBlk_Click()

MsgBox "Opci�n no implementada.", vbInformation
Exit Sub

If LBlk.Caption = "/ Man" Then
    LBlk.Caption = "/ Auto"
    BlockMnu_insert.Enabled = False
    BlockMnu_delete.Enabled = False
    If Trim(Label10.Caption) = "no definido.blk" Then
        Call Label10_Click
    End If
Else
    LBlk.Caption = "/ Man"
    BlockMnu_insert.Enabled = True
    BlockMnu_delete.Enabled = True
End If

End Sub

Private Sub LFin_Click()

If LFin.Caption = "Man" Then
    LFin.Caption = "Auto"
Else
    LFin.Caption = "Man"
End If

End Sub

Private Sub SyncTimer_Timer()

Dim SStream As String
Dim SyncTime As String

Dim a1 As Long
Dim a2 As Long

SStream = Trim(SyncStream.Caption)
SyncTime = Trim(SyncLabel.Caption)  'syn time (in seconds)
a1 = CLng(SyncTime)

Select Case SStream
    Case "Stream01"
        a2 = Stream01GetPosition(1) 'position in seconds
        'MsgBox "Actual Pos: " & a2 & " - Synctime: " & a1
        If a2 >= a1 Then
            SyncStream.Caption = ""
            SyncLabel.Caption = ""
            SyncTimer.Interval = 0
            SyncTimer.Enabled = False
            Call Tanda01.T1Next_Click
        End If
        
    Case "Stream02"
        a2 = Stream02GetPosition(1) 'position in seconds
        'MsgBox "Actual Pos: " & a2 & " - Synctime: " & a1
        If a2 >= a1 Then
            SyncStream.Caption = ""
            SyncLabel.Caption = ""
            SyncTimer.Interval = 0
            SyncTimer.Enabled = False
            Call Tanda01.T1Next_Click
        End If
        
    Case Else
        SyncStream.Caption = ""
        SyncLabel.Caption = ""
        SyncTimer.Interval = 0
        SyncTimer.Enabled = False
End Select

End Sub

Private Sub T1Del_Click()

Dim TTime As String
Dim RTime As String
Dim Time1 As Long
Dim Time2 As Long
Dim Resultado As Long
Dim TMint As Double

On Error Resume Next
'primero extraemos el tiempo del item seleccionado
TTime = Trim(Ltime.Caption)
'RTime = Trim(T1View.SelectedItem.SubItems(3).Text)
'restamos el tiempo del item al del total
Time1 = ConvMinToSec(TTime)
Time2 = ConvMinToSec(RTime)
'tiempo de mixado intermedio
TMint = CDbl(Trim(Tanda01.Intr.text))
Resultado = Time1 - Time2
Resultado = (Resultado + TMint) - 1
Result = ConvSecToMin(Resultado)
'restauramos el display
SetSumTime Result, 1
Ltime.Caption = Result

'eliminamos el item seleccionado
nIndex = T1View.SelectedItem.index
T1View.ListItems.Remove (nIndex)

'seleccionamos el item anterior al mismo
T1View.ListItems.Item(nIndex - 1).Selected = True
If T1View.ListItems.count < 1 Then
    T1View.ListItems.Clear
    Fn.Caption = ""
    Ltime.Caption = "00:00:00"
    Call RestoreDisplay(5)
    'deactivate al controls
    T1Next.Enabled = False
    T1Play.Enabled = False
    T1Stop.Enabled = False
    T1Up.Enabled = False
    T1Down.Enabled = False
    T1Del.Enabled = False
    T1Prop.Enabled = False
    T1Order.Enabled = False
    T1OrderA.Enabled = False
End If
Exit Sub

er:
End Sub

Private Sub T1Down_Click()

Dim DataA(0 To 9) As String
Dim DataKa As String

Dim DataB(0 To 9) As String
Dim DataKb As String

Dim ONum As Integer
Dim nCount As Integer
Dim NNum As Integer

On Error GoTo Continue
'chequeos necesarios
nCount = T1View.ListItems.count
ONum = T1View.SelectedItem.index
NNum = T1View.SelectedItem.index + 1

If NNum > nCount Or nCount = ONum Then Exit Sub

'extraemos los datos del item
DataA(0) = T1View.SelectedItem.text    'file & path
'DataA(1) = T1View.SelectedItem.SubItems(1).Text   'filetype
'DataA(2) = T1View.SelectedItem.SubItems(2).Text  'filename
'DataA(3) = T1View.SelectedItem.SubItems(3).Text
'DataA(4) = T1View.SelectedItem.SubItems(4).Text
'DataA(5) = T1View.SelectedItem.SubItems(5).Text
'DataA(6) = T1View.SelectedItem.SubItems(6).Text
'DataA(7) = T1View.SelectedItem.SubItems(7).Text
'DataA(8) = T1View.SelectedItem.SubItems(8).Text
'DataA(9) = T1View.SelectedItem.SubItems(9).Text
'DataKa = T1View.SelectedItem.Key

'seleccionamos el siguiente item hacia abajo
nIndex = NNum
T1View.ListItems.Item(nIndex).Selected = True

'extraemos los datos del item
DataB(0) = T1View.SelectedItem.text    'file & path
'DataB(1) = T1View.SelectedItem.SubItems(1).Text   'filetype
'DataB(2) = T1View.SelectedItem.SubItems(2).Text  'filename
'DataB(3) = T1View.SelectedItem.SubItems(3).Text
'DataB(4) = T1View.SelectedItem.SubItems(4).Text
'DataB(5) = T1View.SelectedItem.SubItems(5).Text
'DataB(6) = T1View.SelectedItem.SubItems(6).Text
'DataB(7) = T1View.SelectedItem.SubItems(7).Text
'DataB(8) = T1View.SelectedItem.SubItems(8).Text
'DataB(9) = T1View.SelectedItem.SubItems(9).Text
'DataKb = T1View.SelectedItem.Key

'ponemos los nuevos datos
T1View.ListItems.Remove (nIndex)
'Set ItmX = T1View.ListItems.Add(nIndex, DataKb, DataA(0)) 'path & file
'ItmX.SubItems(1) = DataA(1)
'ItmX.SubItems(2) = DataA(2)
'ItmX.SubItems(3) = DataA(3)
'ItmX.SubItems(4) = DataA(4)
'ItmX.SubItems(5) = DataA(5)
'ItmX.SubItems(6) = DataA(6)
'ItmX.SubItems(7) = DataA(7)
'ItmX.SubItems(8) = DataA(8)
'ItmX.SubItems(9) = DataA(9)

'seleccionamos el index anterior
nIndex = nIndex - 1
T1View.ListItems.Item(nIndex).Selected = True

'ponemos los nuevos datos
T1View.ListItems.Remove (nIndex)
'Set ItmX = T1View.ListItems.Add(nIndex, DataKa, DataB(0)) 'path & file
'ItmX.SubItems(1) = DataB(1)
'ItmX.SubItems(2) = DataB(2)
'ItmX.SubItems(3) = DataB(3)
'ItmX.SubItems(4) = DataB(4)
'ItmX.SubItems(5) = DataB(5)
'ItmX.SubItems(6) = DataB(6)
'ItmX.SubItems(7) = DataB(7)
'ItmX.SubItems(8) = DataB(8)
'ItmX.SubItems(9) = DataB(9)

'una vez finalizado. seleccionamos el item
nIndex = nIndex + 1
T1View.ListItems.Item(nIndex).Selected = True
Exit Sub

Continue:
    'nothing to do....
End Sub

Private Sub T1New_Click()

T1View.ListItems.Clear
Fn.Caption = ""
Ltime.Caption = "00:00:00"
Call RestoreDisplay(5)

'deactivate al controls
T1Next.Enabled = False
T1Play.Enabled = False
T1Stop.Enabled = False
T1Up.Enabled = False
T1Down.Enabled = False
T1Del.Enabled = False
T1Prop.Enabled = False
T1Order.Enabled = False
T1OrderA.Enabled = False

End Sub

Sub T1Next_Click()

'gets the count of item in list
If T1View.ListItems.count < 1 Then Exit Sub

'desabilitaciones
T1View.Enabled = False
T1Next.Enabled = False
T1Play.Enabled = False
T1Stop.Enabled = True
T1Up.Enabled = False
T1Down.Enabled = False
T1Del.Enabled = False
T1Prop.Enabled = False
T1Order.Enabled = False
T1OrderA.Enabled = False
T1New.Enabled = False
T1Open.Enabled = False
T1Save.Enabled = False

'deactivate all controls
RestoreAllActiveColor 1

'starts the fadeout
If Stream01IsPlaying = True Or Music01IsPlaying = True Then 'stream01 fade out
    If LFin.Caption = "Auto" Then
        If Est12Control.Origen1.Caption = "E1" Then
            Est01.TmOutAuto.Enabled = True
            Est01.TmOutAuto.Interval = 50
        Else
            TmOut1.Enabled = True
            TmOut1.Interval = 50
        End If
    End If
End If
If Stream02IsPlaying = True Or Music02IsPlaying = True Then 'stream02 fade out
    If LFin.Caption = "Auto" Then
        If Est12Control.Origen2.Caption = "E2" Then
            Est02.TmOutAuto.Enabled = True
            Est02.TmOutAuto.Interval = 50
        Else
            TmOut2.Enabled = True
            TmOut2.Interval = 50
        End If
    End If
End If

'//// gets the file to play
Dim nIndex As Integer

On Error GoTo Continue
nIndex = CInt(LKey.Caption) 'get the index of the file to play
T1View.ListItems.Item(nIndex).Selected = True   'select the item

'gets the file info
FileN = T1View.SelectedItem.text    'file
'FileTP = T1View.SelectedItem.SubItems(1).Text   'filetype
'SSTitle = T1View.SelectedItem.SubItems(2).Text  'file title

'//// checks for file exists
If FileExist(FileN) = False Then
    nIndex = nIndex + 1
    If nIndex > T1View.ListItems.count Then
        nIndex = 1
    End If
    Tanda01.T1View.ListItems.Item(nIndex).Selected = True
    'gets the file info of new file
    FileN = T1View.SelectedItem.text    'file
'    FileTP = T1View.SelectedItem.SubItems(1).Text   'filetype
'    SSTitle = T1View.SelectedItem.SubItems(2).Text  'file title
End If

'****************** FILE CUE & FX PRESETS load...
Dim filename As String
Dim NameFile As String

filename = Trim(FileN)    'extraemos el path y el archivo de audio
NameFile = StripFileFromExt(filename)
filename = Trim(NameFile) & AppCUEFileExt

'**************** COMENZAMOS LA REPRODUCCION DEL AUDIO

'Chequeamos los dispositivos en uso y decidimos cual usar (dev1 or dev2)
If Stream01IsPlaying = True Or Music01IsPlaying = True Then
    If Stream02IsPlaying = True Or Music02IsPlaying = True Then
        'activate the fade in
        TmIn1.Enabled = True
        TmIn1.Interval = 50
        'close stream1 and play the file
        Stream01Stop
        'load and play the selected file
        Call Tanda01Play(FileN, SSTitle, FileTP, "Yes")  '//// USE DEV 1 ////
        'load CUE info & FX info
        OpenCUEFile 1, filename
        'activate the clock timer
        TopMenu.ProcTimer.Enabled = True
        TopMenu.ProcTimer.Interval = 1
    Else
        'activate the fade in
        TmIn2.Enabled = True
        TmIn2.Interval = 50
        'close stream2 and play the file
        Stream02Stop
        'load and play the selected file
        Call Tanda02Play(FileN, SSTitle, FileTP, "Yes")  '//// USE DEV 2 ////
        'load CUE info & FX info
        OpenCUEFile 2, filename
        'activate the clock timer
        TopMenu.ProcTimer.Enabled = True
        TopMenu.ProcTimer.Interval = 1
    End If
Else
    If Stream02IsPlaying = True Or Music02IsPlaying = True Then
        'activate the fade in
        TmIn1.Enabled = True
        TmIn1.Interval = 50
        'close stream1 and play the file
        Stream01Stop
        'load and play the selected file
        Call Tanda01Play(FileN, SSTitle, FileTP, "Yes")  '//// USE DEV 1 ////
        'load CUE info & FX info
        OpenCUEFile 1, filename
        'activate the clock timer
        TopMenu.ProcTimer.Enabled = True
        TopMenu.ProcTimer.Interval = 1
    Else
        'activate the fade in
        TmIn1.Enabled = True
        TmIn1.Interval = 50
        'close stream1 and play the file
        Stream01Stop
        'load and play the selected file
        Call Tanda01Play(FileN, SSTitle, FileTP, "Yes")  '//// USE DEV 1 ////
        'load CUE info & FX info
        OpenCUEFile 1, filename
        'activate the clock timer
        TopMenu.ProcTimer.Enabled = True
        TopMenu.ProcTimer.Interval = 1
    End If
End If

'sets the next item to play
LKey.Caption = nIndex + 1

Tanda01.Caption = "TANDA - Reproduciendo"
'reseteamos la hora de lanzamiento a la hora actual
'del sistema y actualizamos los temas subsiguientes
'a su correspondiente hora de lanzamiento.
Call OrderTndTime("ResetSelected")

If FTime.Caption = "---" Then
    Call SetStartTime
End If
Exit Sub

Continue:
'habilitaciones
'desabilitaciones
T1View.Enabled = True
T1Next.Enabled = True
T1Play.Enabled = True
T1Stop.Enabled = True
T1Up.Enabled = True
T1Down.Enabled = True
T1Del.Enabled = True
T1Prop.Enabled = True
T1Order.Enabled = True
T1OrderA.Enabled = True
T1New.Enabled = True
T1Open.Enabled = True
T1Save.Enabled = True

'disable the sync timer
SyncTimer.Interval = 0
SyncTimer.Enabled = False

End Sub

Private Sub T1Open_Click()

On Error Resume Next
TopMenu.TandaCmd.InitDir = App.path & AppTandaDir & "\"
TopMenu.TandaCmd.Filter = "Archivo de Tanda (*.tnd)|*.tnd|Archivos de Tanda"
TopMenu.TandaCmd.DialogTitle = "TANDAS - Abrir archivo"
TopMenu.TandaCmd.CancelError = True
TopMenu.TandaCmd.ShowOpen

If err.Number = 32755 Then Exit Sub

    'restauramos los valores a 0
    T1View.ListItems.Clear
    Fn.Caption = ""
    Ltime.Caption = "00:00:00"
    Call RestoreDisplay(5)

ConvertTx = TopMenu.TandaCmd.filename

Result = OpenTandaFile(ConvertTx)
If Result = "NotOK" Then
    'MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
    Exit Sub
End If

Fn.Caption = ConvertTx

End Sub

Private Sub T1Order_Click()

Call OrderTndTime("ResetAll")

End Sub

Sub T1OrderA_Click()

Call OrderTndTime("ResetSelected")

End Sub

Sub T1Play_Click()

'get the list item count
If T1View.ListItems.count < 1 Then Exit Sub

'deactivate all controls
RestoreAllActiveColor 1

'starts the fadeout an other device in use (if there is another in use)
If Stream01IsPlaying = True Or Music01IsPlaying = True Then
    If LFin.Caption = "Auto" Then
        If Est12Control.Origen1.Caption = "E1" Then
            Est01.TmOutAuto.Enabled = True
            Est01.TmOutAuto.Interval = 50
        Else
            TmOut1.Enabled = True
            TmOut1.Interval = 50
        End If
    End If
End If

If Stream02IsPlaying = True Or Music02IsPlaying = True Then
    If LFin.Caption = "Auto" Then
        If Est12Control.Origen2.Caption = "E2" Then
            Est02.TmOutAuto.Enabled = True
            Est02.TmOutAuto.Interval = 50
        Else
            TmOut2.Enabled = True
            TmOut2.Interval = 50
        End If
    End If
End If

Dim nIndex As Integer

On Error GoTo err
nIndex = CInt(LKey.Caption) 'get the index of the file to play
T1View.ListItems.Item(nIndex).Selected = True   'select the item

'//// gets the file info
FileN = T1View.SelectedItem.text    'file
'FileTP = T1View.SelectedItem.SubItems(1).Text   'filetype
'SSTitle = T1View.SelectedItem.SubItems(2).Text  'file title

'//// checks for file exists
If FileExist(FileN) = False Then
    nIndex = nIndex + 1
    If nIndex > T1View.ListItems.count Then
        nIndex = 1
    End If
    Tanda01.T1View.ListItems.Item(nIndex).Selected = True
    'gets the file info of new file
    FileN = T1View.SelectedItem.text    'file
'    FileTP = T1View.SelectedItem.SubItems(1).Text   'filetype
'    SSTitle = T1View.SelectedItem.SubItems(2).Text  'file title
End If

'****************** FILE CUE & FX PRESETS load...
Dim filename As String
Dim NameFile As String

filename = Trim(FileN)    'extraemos el path y el archivo de audio
NameFile = StripFileFromExt(filename)
filename = Trim(NameFile) & AppCUEFileExt

'**************** COMENZAMOS LA REPRODUCCION DEL ARCHIVO DE AUDIO
Tanda01.Caption = "TANDA - Reproduciendo"

'Chequeamos los dispositivos en uso y decidimos cual usar (dev1 or dev2)
If Stream01IsPlaying = True Or Music01IsPlaying = True Then
    If Stream02IsPlaying = True Or Music02IsPlaying = True Then
        'activate the fade in
        TmIn1.Enabled = True
        TmIn1.Interval = 50
        'close stream1 and play the file
        Stream01Stop
        'load and play the selected file
        Call Tanda01Play(FileN, SSTitle, FileTP, "No")  '//// USE DEV 1 ////
        'load CUE info & FX info
        OpenCUEFile 1, filename
        'activate the clock timer
        TopMenu.ProcTimer.Enabled = True
        TopMenu.ProcTimer.Interval = 1
    Else
        'activate the fade in
        TmIn2.Enabled = True
        TmIn2.Interval = 50
        'close stream2 and play the file
        Stream02Stop
        'load and play the selected file
        Call Tanda02Play(FileN, SSTitle, FileTP, "No")  '//// USE DEV 2 ////
        'load CUE info & FX info
        OpenCUEFile 2, filename
        'activate the clock timer
        TopMenu.ProcTimer.Enabled = True
        TopMenu.ProcTimer.Interval = 1
    End If
Else
    If Stream02IsPlaying = True Or Music02IsPlaying = True Then
        'activate the fade in
        TmIn1.Enabled = True
        TmIn1.Interval = 50
        'close stream1 and play the file
        Stream01Stop
        'load and play the selected file
        Call Tanda01Play(FileN, SSTitle, FileTP, "No")  '//// USE DEV 1 ////
        'load CUE info & FX info
        OpenCUEFile 1, filename
        'activate the clock timer
        TopMenu.ProcTimer.Enabled = True
        TopMenu.ProcTimer.Interval = 1
    Else
        'activate the fade in
        TmIn1.Enabled = True
        TmIn1.Interval = 50
        'close stream1 and play the file
        Stream01Stop
        'load and play the selected file
        Call Tanda01Play(FileN, SSTitle, FileTP, "No")  '//// USE DEV 1 ////
        'load CUE info & FX info
        OpenCUEFile 1, filename
        'activate the clock timer
        TopMenu.ProcTimer.Enabled = True
        TopMenu.ProcTimer.Interval = 1
    End If
End If

err:
'//// disable the sync timer
SyncTimer.Interval = 0
SyncTimer.Enabled = False

End Sub

Private Sub T1Prop_Click()

Dim filename As String

filename = Trim(Tanda01.T1View.SelectedItem.text)    'file & path

'chequeamos por la validez de los datos
If FileExist(filename) = False Then
    MsgBox "El archivo de audio seleccionado no existe o fu� eliminado.", vbCritical
    MsgBox "Remueva el item de la lista para evitar futuros inconvenientes.", vbInformation
    Exit Sub
Else
    'display de audio prop window
    AudioProp.Show
End If

End Sub

Sub T1Save_Click()

ConvertTxT = Trim(Fn.Caption)

On Error Resume Next
If ConvertTxT = "" Or ConvertTxT = " " Then
    TopMenu.TandaCmd.InitDir = App.path & AppTandaDir & "\"
    TopMenu.TandaCmd.Filter = "Archivo de Tandas (*.tnd)|*.tnd|Archivos de Tanda"
    TopMenu.TandaCmd.DialogTitle = "TANDAS - Guardar archivo"
    TopMenu.TandaCmd.FilterIndex = 1
    TopMenu.TandaCmd.CancelError = True
    TopMenu.TandaCmd.ShowSave

    If err.Number = 32755 Then Exit Sub
    
    ConvertTx = TopMenu.TandaCmd.filename

    Fn.Caption = ConvertTx
    Result = SaveTandaFile(ConvertTx)
    If Result = "NotOK" Then
        MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
        Exit Sub
    End If
Else
    ConvertTx = Trim(Fn.Caption)
    Kill ConvertTx
    Result = SaveTandaFile(ConvertTx)
    If Result = "NotOK" Then
        'MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
        Exit Sub
    End If
End If

End Sub

Private Sub T1Stop_Click()

If T1View.ListItems.count < 1 Then Exit Sub

'habilitaciones
T1View.Enabled = True
T1Next.Enabled = True
T1Play.Enabled = True
T1Stop.Enabled = True
T1Up.Enabled = True
T1Down.Enabled = True
T1Del.Enabled = True
T1Prop.Enabled = True
T1Order.Enabled = True
T1OrderA.Enabled = True
T1New.Enabled = True
T1Open.Enabled = True
T1Save.Enabled = True

If Est12Control.Origen1.Caption = "T1" Then
    If Stream01IsPlaying = True Then
        TmOut1.Enabled = True
        TmOut1.Interval = 50
        Tanda01.Caption = "TANDA - Detenido"
    Else
        If Music01IsPlaying = True Then
            'Music01Stop
            'Music01Restart
            'Tanda01.Caption = "TANDA - Detenido"
        Else
            'nothing to do
        End If
    End If
Else
    'nothing to do
End If

If Est12Control.Origen2.Caption = "T2" Then
    If Stream02IsPlaying = True Then
        TmOut2.Enabled = True
        TmOut2.Interval = 50
        Tanda01.Caption = "TANDA - Detenido"
    Else
        If Music02IsPlaying = True Then
            'Music02Stop
            'Music02Restart
            'Tanda01.Caption = "TANDA - Detenido"
        Else
            'nothing to do
        End If
    End If
Else
    'nothing to do
End If

'disable the sync timer
SyncTimer.Interval = 0
SyncTimer.Enabled = False

'set the time to nothing
FTime.Caption = "---"

End Sub

Private Sub T1Up_Click()

Dim DataA(0 To 9) As String
Dim DataKa As String

Dim DataB(0 To 9) As String
Dim DataKb As String

Dim ONum As Integer
Dim nCount As Integer
Dim NNum As Integer

On Error GoTo Continue
'chequeos necesarios
nCount = T1View.ListItems.count
ONum = T1View.SelectedItem.index
NNum = T1View.SelectedItem.index - 1

If NNum < 0 Or ONum = 1 Then Exit Sub

'extraemos los datos del item seleccionado
DataA(0) = T1View.SelectedItem.text    'file & path
'DataA(1) = T1View.SelectedItem.SubItems(1).Text   'filetype
'DataA(2) = T1View.SelectedItem.SubItems(2).Text  'filename
'DataA(3) = T1View.SelectedItem.SubItems(3).Text
'DataA(4) = T1View.SelectedItem.SubItems(4).Text
'DataA(5) = T1View.SelectedItem.SubItems(5).Text
'DataA(6) = T1View.SelectedItem.SubItems(6).Text
'DataA(7) = T1View.SelectedItem.SubItems(7).Text
'DataA(8) = T1View.SelectedItem.SubItems(8).Text
'DataA(9) = T1View.SelectedItem.SubItems(9).Text
DataKa = T1View.SelectedItem.Key

'seleccionamos el siguiente item hacia abajo
nIndex = NNum
T1View.ListItems.Item(nIndex).Selected = True

'extraemos los datos del item
DataB(0) = T1View.SelectedItem.text    'file & path
'DataB(1) = T1View.SelectedItem.SubItems(1).Text   'filetype
'DataB(2) = T1View.SelectedItem.SubItems(2).Text  'filename
'DataB(3) = T1View.SelectedItem.SubItems(3).Text
'DataB(4) = T1View.SelectedItem.SubItems(4).Text
'DataB(5) = T1View.SelectedItem.SubItems(5).Text
'DataB(6) = T1View.SelectedItem.SubItems(6).Text
'DataB(7) = T1View.SelectedItem.SubItems(7).Text
'DataB(8) = T1View.SelectedItem.SubItems(8).Text
'DataB(9) = T1View.SelectedItem.SubItems(9).Text
DataKb = T1View.SelectedItem.Key

'ponemos los nuevos datos
T1View.ListItems.Remove (nIndex)
'Set ItmX = T1View.ListItems.Add(nIndex, DataKb, DataA(0)) 'path & file
'ItmX.SubItems(1) = DataA(1)
'ItmX.SubItems(2) = DataA(2)
'ItmX.SubItems(3) = DataA(3)
'ItmX.SubItems(4) = DataA(4)
'ItmX.SubItems(5) = DataA(5)
'ItmX.SubItems(6) = DataA(6)
'ItmX.SubItems(7) = DataA(7)
'ItmX.SubItems(8) = DataA(8)
'ItmX.SubItems(9) = DataA(9)

'seleccionamos el index anterior
nIndex = nIndex + 1
T1View.ListItems.Item(nIndex).Selected = True

'ponemos los nuevos datos
T1View.ListItems.Remove (nIndex)
'Set ItmX = T1View.ListItems.Add(nIndex, DataKa, DataB(0)) 'path & file
'ItmX.SubItems(1) = DataB(1)
'ItmX.SubItems(2) = DataB(2)
'ItmX.SubItems(3) = DataB(3)
'ItmX.SubItems(4) = DataB(4)
'ItmX.SubItems(5) = DataB(5)
'ItmX.SubItems(6) = DataB(6)
'ItmX.SubItems(7) = DataB(7)
'ItmX.SubItems(8) = DataB(8)
'ItmX.SubItems(9) = DataB(9)

'una vez finalizado. seleccionamos el item
nIndex = nIndex - 1
T1View.ListItems.Item(nIndex).Selected = True
Exit Sub

Continue:
    'nothing to do....
End Sub

Private Sub T1View_Click()

On Error GoTo er
LKey.Caption = T1View.SelectedItem.index
T1Next.Enabled = True
T1Play.Enabled = True
T1Stop.Enabled = True
T1Up.Enabled = True
T1Down.Enabled = True
T1Del.Enabled = True
T1Prop.Enabled = True
T1Order.Enabled = True
T1OrderA.Enabled = True
Exit Sub

er:
T1Next.Enabled = False
T1Play.Enabled = False
T1Stop.Enabled = False
T1Up.Enabled = False
T1Down.Enabled = False
T1Del.Enabled = False
T1Prop.Enabled = False
T1Order.Enabled = False
T1OrderA.Enabled = False
End Sub

Private Sub T1View_DblClick()

Call T1Play_Click

End Sub

Private Sub T1View_DragDrop(Source As Control, X As Single, Y As Single)

DeployTandaFile 'drag & drop the selected file in xplorer

End Sub

Private Sub T1View_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

Select Case State
    Case 0  'drag not finished
        XPlorer.File1.DragIcon = XPlorer.ExCombo.DragIcon
        'E11(Index).BackColor = &H80FF80    'verde (modificacion)
    Case 1  'finished drag
        XPlorer.File1.DragIcon = XPlorer.tvwDirTree.DragIcon
        'E11(Index).BackColor = &H8000000F  'gris (normal)
End Select

End Sub

Private Sub T1View_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'button 1=left button
'button 2=right button
'button 4=mid button

If Button = 2 Then
    'xxxxxx
Else
    'xxxxxx
End If

End Sub

Private Sub T1Vol_Change()

If Est12Control.StopLabel1.Caption = "Stream" Then
    'change the stream volume
    Stream01SetVolume (T1Vol.Value)
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
        'change the music volume
        Music01SetVolume (T1Vol.Value)
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub T1Vol_Scroll()

If Est12Control.StopLabel1.Caption = "Stream" Then
    'change the stream volume
    Stream01SetVolume (T1Vol.Value)
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
        'change the music volume
        Music01SetVolume (T1Vol.Value)
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub T2Vol_Change()

If Est12Control.StopLabel2.Caption = "Stream" Then
    'change the stream volume
    Stream02SetVolume (T2Vol.Value)
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
        'change the music volume
        Music02SetVolume (T2Vol.Value)
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub T2Vol_Scroll()

If Est12Control.StopLabel2.Caption = "Stream" Then
    'change the stream volume
    Stream02SetVolume (T2Vol.Value)
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
        'change the music volume
        Music02SetVolume (T2Vol.Value)
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub TmIn1_Timer()

If T1Vol.Value = 100 Then
    TmIn1.Interval = 0
    TmIn1.Enabled = False
    Exit Sub
Else
    T1Vol.Value = T1Vol.Value + 5
End If

End Sub

Private Sub TmIn2_Timer()

If T2Vol.Value = 100 Then
    TmIn2.Interval = 0
    TmIn2.Enabled = False
    Exit Sub
Else
    T2Vol.Value = T2Vol.Value + 5
End If

End Sub

Private Sub TmOut1_Timer()

If T1Vol.Value = 0 Then
    If Est12Control.StopLabel1.Caption = "Stream" And Est12Control.Origen1.Caption = "T1" Then
        Stream01Restart    'stream restart
        Stream01Stop       'stream stop
    Else
        If Est12Control.StopLabel1.Caption = "Music" And Est12Control.Origen1.Caption = "T1" Then
            Music01Restart     'music restart
            Music01Stop         'music stop
        Else
            Exit Sub
        End If
    End If
    TmOut1.Interval = 0
    TmOut1.Enabled = False
Else
    T1Vol.Value = T1Vol.Value - 5
End If

End Sub

Private Sub TmOut2_Timer()

If T2Vol.Value = 0 Then
    If Est12Control.StopLabel1.Caption = "Stream" And Est12Control.Origen2.Caption = "T2" Then
        Stream02Restart    'stream restart
        Stream02Stop       'stream stop
    Else
        If Est12Control.StopLabel1.Caption = "Music" And Est12Control.Origen2.Caption = "T2" Then
            Music02Restart     'music restart
            Music02Stop         'music stop
        Else
            Exit Sub
        End If
    End If
    TmOut2.Interval = 0
    TmOut2.Enabled = False
Else
    T2Vol.Value = T2Vol.Value - 5
End If

End Sub

Private Sub VScroll1_Change()

Intr.text = VScroll1.Value

End Sub


