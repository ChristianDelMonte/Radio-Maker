VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Config 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Panel de Configuraci�n"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8790
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "Config.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CmdPrg 
      Left            =   135
      Top             =   5895
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cc"
      Height          =   375
      Left            =   7740
      TabIndex        =   3
      ToolTipText     =   "Cancelar"
      Top             =   5940
      Width           =   960
   End
   Begin VB.CommandButton CmdAply 
      Caption         =   "aP"
      Height          =   375
      Left            =   5580
      TabIndex        =   2
      ToolTipText     =   "Aplicar los cambios de configuraci�n."
      Top             =   5940
      Width           =   960
   End
   Begin VB.CommandButton CmdAccept 
      Caption         =   "Ac"
      Height          =   375
      Left            =   6660
      TabIndex        =   1
      ToolTipText     =   "Aceptar y guardar los cambios de configuraci�n"
      Top             =   5940
      Width           =   960
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5595
      Left            =   135
      TabIndex        =   0
      Top             =   270
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   9869
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "Config.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Audio"
      TabPicture(1)   =   "Config.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame10"
      Tab(1).Control(1)=   "Frame9"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Directorios"
      TabPicture(2)   =   "Config.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Seguridad"
      TabPicture(3)   =   "Config.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DefPass"
      Tab(3).Control(1)=   "Frame8"
      Tab(3).Control(2)=   "Frame7"
      Tab(3).Control(3)=   "Frame6"
      Tab(3).Control(4)=   "Frame5"
      Tab(3).Control(5)=   "Frame4"
      Tab(3).ControlCount=   6
      Begin VB.Frame Frame13 
         Caption         =   "Opciones Generales"
         Height          =   4695
         Left            =   225
         TabIndex        =   113
         Top             =   675
         Width           =   8160
         Begin VB.Frame Frame18 
            Caption         =   "Edici�n / Producci�n "
            Height          =   960
            Left            =   135
            TabIndex        =   127
            Top             =   2970
            Width           =   7890
            Begin VB.TextBox GrabName 
               Height          =   285
               Left            =   4095
               Locked          =   -1  'True
               TabIndex        =   132
               Text            =   "ninguno seleccionado"
               Top             =   540
               Width           =   3210
            End
            Begin VB.CommandButton GrabSearch 
               Caption         =   "..."
               Height          =   285
               Left            =   7380
               TabIndex        =   131
               Top             =   540
               Width           =   375
            End
            Begin VB.TextBox EdName 
               Height          =   285
               Left            =   135
               Locked          =   -1  'True
               TabIndex        =   129
               Text            =   "ninguno seleccionado"
               Top             =   540
               Width           =   3210
            End
            Begin VB.CommandButton EdSearch 
               Caption         =   "..."
               Height          =   285
               Left            =   3420
               TabIndex        =   128
               Top             =   540
               Width           =   375
            End
            Begin VB.Label Label13 
               Caption         =   "Seleccione el programa grabador de audio"
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   4095
               TabIndex        =   133
               Top             =   315
               Width           =   3210
            End
            Begin VB.Label Label12 
               Caption         =   "Seleccione el programa editor de audio"
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   135
               TabIndex        =   130
               Top             =   315
               Width           =   3210
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "Reportes"
            Height          =   1500
            Left            =   135
            TabIndex        =   118
            Top             =   1395
            Width           =   7890
            Begin VB.Frame Frame17 
               Caption         =   "Reportar"
               Height          =   1230
               Left            =   3960
               TabIndex        =   123
               Top             =   180
               Width           =   3795
               Begin VB.CheckBox RepAll 
                  Caption         =   "Reportar todas las reproducciones"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   126
                  Top             =   900
                  Value           =   1  'Checked
                  Width           =   2760
               End
               Begin VB.CheckBox Rep2 
                  Caption         =   "Reportar reproducciones en Tanda 01 y 02"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   125
                  Top             =   630
                  Width           =   3435
               End
               Begin VB.CheckBox Rep1 
                  Caption         =   "Reportar reproducciones en Estacion 01 y 02"
                  Height          =   240
                  Left            =   135
                  TabIndex        =   124
                  Top             =   315
                  Width           =   3570
               End
            End
            Begin VB.CommandButton ARepSearch 
               Caption         =   "..."
               Enabled         =   0   'False
               Height          =   285
               Left            =   3420
               TabIndex        =   121
               Top             =   1080
               Width           =   375
            End
            Begin VB.TextBox ARepName 
               Enabled         =   0   'False
               Height          =   285
               Left            =   135
               Locked          =   -1  'True
               TabIndex        =   120
               Text            =   "ninguno seleccionado"
               Top             =   1080
               Width           =   3210
            End
            Begin VB.CheckBox ARep 
               Caption         =   "Activar modulo creador de reporte de reproducciones."
               Height          =   375
               Left            =   135
               TabIndex        =   119
               Top             =   360
               Value           =   1  'Checked
               Width           =   2265
            End
            Begin VB.Label Label11 
               Caption         =   "Seleccione el programa editor de reportes"
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   135
               TabIndex        =   122
               Top             =   855
               Width           =   3210
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "ST, S3m, IT, Mod, XM, OGG"
            Height          =   870
            Left            =   3915
            TabIndex        =   116
            Top             =   405
            Width           =   4110
            Begin VB.CheckBox EName 
               Caption         =   "Extraer informaci�n del archivo autom�ticamente al cargar."
               Height          =   375
               Left            =   90
               TabIndex        =   117
               Top             =   360
               Value           =   1  'Checked
               Width           =   3840
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Mp1, Mp2, Mp3"
            Height          =   870
            Left            =   135
            TabIndex        =   114
            Top             =   405
            Width           =   3615
            Begin VB.CheckBox Etag 
               Caption         =   "Extraer informaci�n TAG autom�ticamente al cargar el archivo."
               Height          =   420
               Left            =   135
               TabIndex        =   115
               Top             =   360
               Value           =   1  'Checked
               Width           =   3345
            End
         End
      End
      Begin VB.CommandButton DefPass 
         Caption         =   "&Definir password"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -74820
         TabIndex        =   96
         Top             =   4770
         Width           =   1545
      End
      Begin VB.Frame Frame10 
         Caption         =   "Visualizaci�n"
         Height          =   4695
         Left            =   -71175
         TabIndex        =   61
         Top             =   675
         Width           =   3435
         Begin VB.CheckBox MVSCOPE 
            Caption         =   "Mostrar Viz SCOPE"
            Height          =   240
            Left            =   150
            TabIndex        =   136
            Top             =   4305
            Value           =   1  'Checked
            Width           =   1995
         End
         Begin VB.CheckBox MVFFT 
            Caption         =   "Mostrar Viz FFT"
            Height          =   195
            Left            =   150
            TabIndex        =   135
            Top             =   4065
            Value           =   1  'Checked
            Width           =   1710
         End
         Begin VB.CheckBox MMRm 
            Caption         =   "Mostrar MiniPlayer al minimizar"
            Height          =   255
            Left            =   150
            TabIndex        =   134
            Top             =   3780
            Value           =   1  'Checked
            Width           =   2550
         End
         Begin VB.Frame Frame12 
            Caption         =   "ST, S3m, IT, Mod, XM, OGG"
            Height          =   1215
            Left            =   135
            TabIndex        =   95
            Top             =   2430
            Width           =   3165
            Begin VB.CheckBox Sr 
               Caption         =   "Samples restante"
               Height          =   195
               Left            =   1530
               TabIndex        =   112
               Top             =   810
               Width           =   1545
            End
            Begin VB.PictureBox Picture24 
               BackColor       =   &H00000000&
               Height          =   330
               Left            =   135
               ScaleHeight     =   270
               ScaleWidth      =   1230
               TabIndex        =   105
               Top             =   765
               Width           =   1290
               Begin VB.PictureBox sr6 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   990
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   111
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox sr1 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   45
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   110
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox sr2 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   240
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   109
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox sr3 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   420
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   108
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox sr4 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   615
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   107
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox sr5 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   810
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   106
                  Top             =   30
                  Width           =   190
               End
            End
            Begin VB.CheckBox Sn 
               Caption         =   "Samples normal"
               Height          =   195
               Left            =   1530
               TabIndex        =   104
               Top             =   405
               Value           =   1  'Checked
               Width           =   1455
            End
            Begin VB.PictureBox Picture17 
               BackColor       =   &H00000000&
               Height          =   330
               Left            =   135
               ScaleHeight     =   270
               ScaleWidth      =   1230
               TabIndex        =   97
               Top             =   360
               Width           =   1290
               Begin VB.PictureBox sn6 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   990
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   103
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox sn1 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   45
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   102
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox sn2 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   240
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   101
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox sn3 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   420
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   100
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox sn4 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   615
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   99
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox sn5 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   810
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   98
                  Top             =   30
                  Width           =   190
               End
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Wav, Mp1, Mp2, Mp3"
            Height          =   2040
            Left            =   135
            TabIndex        =   62
            Top             =   315
            Width           =   3165
            Begin VB.CheckBox Ore 
               Caption         =   "Ondas restante"
               Height          =   195
               Left            =   1530
               TabIndex        =   94
               Top             =   1620
               Width           =   1410
            End
            Begin VB.PictureBox Picture10 
               BackColor       =   &H00000000&
               Height          =   330
               Left            =   135
               ScaleHeight     =   270
               ScaleWidth      =   1230
               TabIndex        =   87
               Top             =   1575
               Width           =   1290
               Begin VB.PictureBox or5 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   810
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   93
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox or4 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   615
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   92
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox or3 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   420
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   91
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox or2 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   240
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   90
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox or1 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   45
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   89
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox or6 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   990
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   88
                  Top             =   30
                  Width           =   190
               End
            End
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00000000&
               Height          =   330
               Left            =   135
               ScaleHeight     =   270
               ScaleWidth      =   1230
               TabIndex        =   80
               Top             =   1170
               Width           =   1290
               Begin VB.PictureBox on5 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   810
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   86
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox on4 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   615
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   85
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox on3 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   420
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   84
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox on2 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   240
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   83
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox on1 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   45
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   82
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox on6 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   990
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   81
                  Top             =   30
                  Width           =   190
               End
            End
            Begin VB.CheckBox Ono 
               Caption         =   "Ondas normal"
               Height          =   195
               Left            =   1530
               TabIndex        =   79
               Top             =   1215
               Value           =   1  'Checked
               Width           =   1365
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H00000000&
               Height          =   330
               Left            =   135
               ScaleHeight     =   270
               ScaleWidth      =   1230
               TabIndex        =   72
               Top             =   765
               Width           =   1290
               Begin VB.PictureBox tr6 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   990
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   78
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox tr1 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   45
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   77
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox tr2 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   240
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   76
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox tr3 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   420
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   75
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox tr4 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   615
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   74
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox tr5 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   810
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   73
                  Top             =   30
                  Width           =   190
               End
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00000000&
               Height          =   330
               Left            =   135
               ScaleHeight     =   270
               ScaleWidth      =   1230
               TabIndex        =   65
               Top             =   360
               Width           =   1290
               Begin VB.PictureBox tn5 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   810
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   71
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox tn4 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   615
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   70
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox tn3 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   420
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   69
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox tn2 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   240
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   68
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox tn1 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   45
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   67
                  Top             =   30
                  Width           =   190
               End
               Begin VB.PictureBox tn6 
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   210
                  Left            =   990
                  ScaleHeight     =   210
                  ScaleWidth      =   195
                  TabIndex        =   66
                  Top             =   30
                  Width           =   190
               End
            End
            Begin VB.CheckBox Tr 
               Caption         =   "Tiempo restante"
               Height          =   195
               Left            =   1530
               TabIndex        =   64
               Top             =   810
               Width           =   1500
            End
            Begin VB.CheckBox Tn 
               Caption         =   "Tiempo normal"
               Height          =   195
               Left            =   1530
               TabIndex        =   63
               Top             =   405
               Value           =   1  'Checked
               Width           =   1410
            End
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Reproducci�n"
         Height          =   4695
         Left            =   -74775
         TabIndex        =   44
         Top             =   675
         Width           =   3390
         Begin VB.Frame Frame2 
            Caption         =   "Archivos ST, S3m, IT, Mod, XM, OGG"
            Height          =   2715
            Left            =   135
            TabIndex        =   54
            Top             =   1845
            Width           =   3120
            Begin VB.CheckBox RN 
               Caption         =   "Ramping Normal"
               Height          =   195
               Left            =   135
               TabIndex        =   59
               Top             =   360
               Value           =   1  'Checked
               Width           =   1545
            End
            Begin VB.CheckBox RS 
               Caption         =   "Ramping Sensitivo"
               Height          =   195
               Left            =   135
               TabIndex        =   58
               Top             =   630
               Width           =   1680
            End
            Begin VB.CheckBox RFt2 
               Caption         =   "Reproducir como FastTracker 2"
               Height          =   195
               Left            =   135
               TabIndex        =   57
               Top             =   1890
               Value           =   1  'Checked
               Width           =   2625
            End
            Begin VB.CheckBox RPt2 
               Caption         =   "Reproducir como ProTracker 1"
               Height          =   195
               Left            =   135
               TabIndex        =   56
               Top             =   2160
               Width           =   2490
            End
            Begin VB.CheckBox SS 
               Caption         =   "Sonido Surround"
               Height          =   195
               Left            =   135
               TabIndex        =   55
               Top             =   2430
               Width           =   2580
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "El ramping mejora la calidad del sonido removiendo los ruidos que este pueda tener. No consume recursos de harware extras."
               ForeColor       =   &H00404040&
               Height          =   825
               Left            =   405
               TabIndex        =   60
               Top             =   900
               Width           =   2490
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Dispositivos"
            Height          =   1410
            Left            =   135
            TabIndex        =   45
            Top             =   315
            Width           =   3120
            Begin VB.CheckBox M8b 
               Caption         =   "8 Bits"
               Height          =   240
               Left            =   135
               TabIndex        =   53
               Top             =   315
               Width           =   780
            End
            Begin VB.CheckBox M16B 
               Caption         =   "16 Bits"
               Height          =   240
               Left            =   1620
               TabIndex        =   52
               Top             =   315
               Value           =   1  'Checked
               Width           =   870
            End
            Begin VB.CheckBox MM 
               Caption         =   "Mono"
               Height          =   195
               Left            =   135
               TabIndex        =   51
               Top             =   585
               Width           =   825
            End
            Begin VB.CheckBox Ms 
               Caption         =   "Estereo"
               Height          =   195
               Left            =   1620
               TabIndex        =   50
               Top             =   585
               Value           =   1  'Checked
               Width           =   915
            End
            Begin VB.CheckBox M3d 
               Caption         =   "3D"
               Height          =   195
               Left            =   135
               TabIndex        =   49
               Top             =   1125
               Width           =   555
            End
            Begin VB.CheckBox MA3d 
               Caption         =   "A3D"
               Height          =   195
               Left            =   1620
               TabIndex        =   48
               Top             =   855
               Width           =   690
            End
            Begin VB.CheckBox Mogg 
               Caption         =   "OGG"
               Height          =   195
               Left            =   1620
               TabIndex        =   47
               Top             =   1125
               Width           =   690
            End
            Begin VB.CheckBox MN 
               Caption         =   "Normal"
               Height          =   195
               Left            =   135
               TabIndex        =   46
               Top             =   855
               Value           =   1  'Checked
               Width           =   825
            End
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Secciones especiales"
         Height          =   915
         Left            =   -74820
         TabIndex        =   37
         Top             =   3780
         Width           =   8250
         Begin VB.CheckBox c5 
            Caption         =   "Al ejecutar un Plug-In"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5715
            TabIndex        =   43
            Top             =   585
            Width           =   1905
         End
         Begin VB.CheckBox c3 
            Caption         =   "Al intentar modificar opciones."
            Enabled         =   0   'False
            Height          =   195
            Left            =   2925
            TabIndex        =   42
            Top             =   585
            Width           =   2490
         End
         Begin VB.CheckBox c1 
            Caption         =   "Al entrar en Configuraci�n."
            Enabled         =   0   'False
            Height          =   195
            Left            =   135
            TabIndex        =   41
            Top             =   585
            Width           =   2220
         End
         Begin VB.CheckBox c4 
            Caption         =   "Al salir de Radio Maker"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5715
            TabIndex        =   40
            Top             =   315
            Width           =   2175
         End
         Begin VB.CheckBox c2 
            Caption         =   "Al entrar en Radio Maker"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2925
            TabIndex        =   38
            Top             =   315
            Width           =   2220
         End
         Begin VB.Label Label10 
            Caption         =   "Aplicar protecci�n al..."
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   90
            TabIndex        =   39
            Top             =   315
            Width           =   1860
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "General"
         Height          =   600
         Left            =   -74820
         TabIndex        =   24
         Top             =   585
         Width           =   8250
         Begin VB.CheckBox None 
            Caption         =   "Ninguna"
            Height          =   195
            Left            =   2655
            TabIndex        =   28
            Top             =   270
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.CheckBox Den 
            Caption         =   "Denegar acceso"
            Height          =   195
            Left            =   5985
            TabIndex        =   27
            Top             =   270
            Width           =   1545
         End
         Begin VB.CheckBox Pass 
            Caption         =   "Solicitar Password"
            Height          =   195
            Left            =   3960
            TabIndex        =   26
            Top             =   270
            Width           =   1635
         End
         Begin VB.Label Label1 
            Caption         =   "�Qu� tipo de protecci�n desea?"
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   135
            TabIndex        =   25
            Top             =   270
            Width           =   2355
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Programaci�n de Tandas"
         Height          =   2355
         Left            =   -69240
         TabIndex        =   23
         Top             =   1305
         Width           =   2670
         Begin VB.CheckBox p3 
            Caption         =   "Deplegar / Eliminar / Modificar una Tanda dentro de la Programaci�n."
            Enabled         =   0   'False
            Height          =   645
            Left            =   90
            TabIndex        =   35
            Top             =   1665
            Width           =   2490
         End
         Begin VB.CheckBox p2 
            Caption         =   "Reproducir / Detener / Pausar una reproducci�n."
            Enabled         =   0   'False
            Height          =   420
            Left            =   90
            TabIndex        =   34
            Top             =   1170
            Width           =   2490
         End
         Begin VB.CheckBox p1 
            Caption         =   "Abrir / Guardar / Crear una nueva Programaci�n."
            Enabled         =   0   'False
            Height          =   420
            Left            =   90
            TabIndex        =   33
            Top             =   675
            Width           =   2220
         End
         Begin VB.Label Label9 
            Caption         =   "Aplicar protecci�n al..."
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   45
            TabIndex        =   36
            Top             =   360
            Width           =   2580
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tanda 01"
         Height          =   2355
         Left            =   -72030
         TabIndex        =   22
         Top             =   1305
         Width           =   2715
         Begin VB.CheckBox t3 
            Caption         =   "Deplegar / Eliminar / Modificar un tema dentro de la Tanda."
            Enabled         =   0   'False
            Height          =   420
            Left            =   135
            TabIndex        =   31
            Top             =   1665
            Width           =   2490
         End
         Begin VB.CheckBox t2 
            Caption         =   "Reproducir / Detener / Pausar una reproducci�n."
            Enabled         =   0   'False
            Height          =   420
            Left            =   135
            TabIndex        =   30
            Top             =   1170
            Width           =   2490
         End
         Begin VB.CheckBox t1 
            Caption         =   "Abrir / Guardar / Crear una nueva Tanda"
            Enabled         =   0   'False
            Height          =   420
            Left            =   135
            TabIndex        =   29
            Top             =   675
            Width           =   2220
         End
         Begin VB.Label Label3 
            Caption         =   "Aplicar protecci�n al..."
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   90
            TabIndex        =   32
            Top             =   360
            Width           =   2580
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Estacion 01 y 02"
         Height          =   2355
         Left            =   -74820
         TabIndex        =   17
         Top             =   1305
         Width           =   2715
         Begin VB.CheckBox e3 
            Caption         =   "Deplegar / Eliminar / Modificar un tema dentro de la estaci�n."
            Enabled         =   0   'False
            Height          =   420
            Left            =   135
            TabIndex        =   20
            Top             =   1665
            Width           =   2490
         End
         Begin VB.CheckBox e2 
            Caption         =   "Reproducir / Detener / Pausar una reproducci�n."
            Enabled         =   0   'False
            Height          =   420
            Left            =   135
            TabIndex        =   19
            Top             =   1170
            Width           =   2490
         End
         Begin VB.CheckBox e1 
            Caption         =   "Abrir / Guardar / Crear un nuevo archivo."
            Enabled         =   0   'False
            Height          =   420
            Left            =   135
            TabIndex        =   18
            Top             =   675
            Width           =   2220
         End
         Begin VB.Label Label2 
            Caption         =   "Aplicar protecci�n al..."
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   90
            TabIndex        =   21
            Top             =   360
            Width           =   2580
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Configuraci�n de Directorios"
         Height          =   3075
         Left            =   -74820
         TabIndex        =   4
         Top             =   720
         Width           =   3075
         Begin VB.CommandButton Exam 
            Caption         =   "..."
            Height          =   285
            Index           =   3
            Left            =   2565
            TabIndex        =   16
            Top             =   2655
            Width           =   375
         End
         Begin VB.CommandButton Exam 
            Caption         =   "..."
            Height          =   285
            Index           =   2
            Left            =   2565
            TabIndex        =   15
            Top             =   1980
            Width           =   375
         End
         Begin VB.CommandButton Exam 
            Caption         =   "..."
            Height          =   285
            Index           =   1
            Left            =   2565
            TabIndex        =   14
            Top             =   1305
            Width           =   375
         End
         Begin VB.TextBox Tx 
            Height          =   285
            Index           =   3
            Left            =   135
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "D:\Audio\Grabaciones"
            Top             =   2655
            Width           =   2400
         End
         Begin VB.TextBox Tx 
            Height          =   285
            Index           =   2
            Left            =   135
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "D:\Audio\Institucionales"
            Top             =   1980
            Width           =   2400
         End
         Begin VB.TextBox Tx 
            Height          =   285
            Index           =   1
            Left            =   135
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "D:\Audio\Comerciales"
            Top             =   1305
            Width           =   2400
         End
         Begin VB.CommandButton Exam 
            Caption         =   "..."
            Height          =   285
            Index           =   0
            Left            =   2565
            TabIndex        =   10
            Top             =   630
            Width           =   375
         End
         Begin VB.TextBox Tx 
            Height          =   285
            Index           =   0
            Left            =   135
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "D:\Audio\Music"
            Top             =   630
            Width           =   2400
         End
         Begin VB.Label Label8 
            Caption         =   "Grabaciones horarias"
            Height          =   195
            Left            =   135
            TabIndex        =   9
            Top             =   2430
            Width           =   1590
         End
         Begin VB.Label Label7 
            Caption         =   "Institucionales"
            Height          =   195
            Left            =   135
            TabIndex        =   8
            Top             =   1755
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Comerciales"
            Height          =   195
            Left            =   135
            TabIndex        =   7
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label5 
            Caption         =   "Temas"
            Height          =   195
            Left            =   135
            TabIndex        =   6
            Top             =   405
            Width           =   600
         End
      End
   End
End
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Paint_Clocks()

'tiempo normal
tn1.Picture = TopMenu.SmallClip.GraphicCell(10)
tn2.Picture = TopMenu.SmallClip.GraphicCell(0)
tn3.Picture = TopMenu.SmallClip.GraphicCell(4)
tn4.Picture = TopMenu.SmallClip.GraphicCell(11)
tn5.Picture = TopMenu.SmallClip.GraphicCell(5)
tn6.Picture = TopMenu.SmallClip.GraphicCell(2)

'tiempo restante
tr1.Picture = TopMenu.SmallClip.GraphicCell(13)
tr2.Picture = TopMenu.SmallClip.GraphicCell(0)
tr3.Picture = TopMenu.SmallClip.GraphicCell(4)
tr4.Picture = TopMenu.SmallClip.GraphicCell(11)
tr5.Picture = TopMenu.SmallClip.GraphicCell(5)
tr6.Picture = TopMenu.SmallClip.GraphicCell(2)

'Ondas normal
on1.Picture = TopMenu.SmallClip.GraphicCell(10)
on2.Picture = TopMenu.SmallClip.GraphicCell(2)
on3.Picture = TopMenu.SmallClip.GraphicCell(5)
on4.Picture = TopMenu.SmallClip.GraphicCell(2)
on5.Picture = TopMenu.SmallClip.GraphicCell(7)
on6.Picture = TopMenu.SmallClip.GraphicCell(0)

'Ondas restante
or1.Picture = TopMenu.SmallClip.GraphicCell(13)
or2.Picture = TopMenu.SmallClip.GraphicCell(2)
or3.Picture = TopMenu.SmallClip.GraphicCell(5)
or4.Picture = TopMenu.SmallClip.GraphicCell(2)
or5.Picture = TopMenu.SmallClip.GraphicCell(7)
or6.Picture = TopMenu.SmallClip.GraphicCell(0)

'Samples Normal
sn1.Picture = TopMenu.SmallClip.GraphicCell(10)
sn2.Picture = TopMenu.SmallClip.GraphicCell(9)
sn3.Picture = TopMenu.SmallClip.GraphicCell(5)
sn4.Picture = TopMenu.SmallClip.GraphicCell(4)
sn5.Picture = TopMenu.SmallClip.GraphicCell(5)
sn6.Picture = TopMenu.SmallClip.GraphicCell(9)

'Samples Restante
sr1.Picture = TopMenu.SmallClip.GraphicCell(13)
sr2.Picture = TopMenu.SmallClip.GraphicCell(9)
sr3.Picture = TopMenu.SmallClip.GraphicCell(2)
sr4.Picture = TopMenu.SmallClip.GraphicCell(4)
sr5.Picture = TopMenu.SmallClip.GraphicCell(5)
sr6.Picture = TopMenu.SmallClip.GraphicCell(9)

End Sub

Private Sub ARep_Click()

If ARep.value = 1 Then
    ARepName.Enabled = True
    ARepSearch.Enabled = True
    RepAll.Enabled = True
    If RepAll.value = 1 Then
        Rep1.Enabled = False
        Rep2.Enabled = False
    Else
        Rep1.Enabled = True
        Rep2.Enabled = True
    End If
Else
    ARepName.Enabled = False
    ARepSearch.Enabled = False
    Rep1.Enabled = False
    Rep2.Enabled = False
    RepAll.Enabled = False
End If

End Sub

Private Sub ARepSearch_Click()

Dim PRGName As String

On Error Resume Next
CmdPrg.InitDir = App.Path
CmdPrg.Filter = "Archivo de Programas (*.exe)|*.exe|Archivos de Programa"
CmdPrg.DialogTitle = "RM100 - Seleccione archivo de programa"
CmdPrg.CancelError = True
CmdPrg.ShowOpen

If err.Number = 32755 Then Exit Sub

PRGName = CmdPrg.filename
ARepName.Text = PRGName

End Sub

Private Sub CmdAccept_Click()

Dim ConfigData As ConfigRecord   'registros de Configuracion
Dim Result As String

'GENERAL OPTIONS
ConfigData.Gen_AutoTAG = Etag.value
ConfigData.Gen_AutoName = EName.value
ConfigData.Gen_ActiveReport = ARep.value
ConfigData.Gen_ReportEst = Rep1.value
ConfigData.Gen_ReportTnd = Rep2.value
ConfigData.Gen_ReportAll = RepAll.value
ConfigData.Gen_ReportProg = SetCipherConfigData(Trim(ARepName.Text))
ConfigData.Gen_EditProg = SetCipherConfigData(Trim(EdName.Text))
ConfigData.Gen_GrabProg = SetCipherConfigData(Trim(GrabName.Text))

'AUDIO OPTIONS
If M8b.value = 1 Then       '1=8bits    2=16bits
    ConfigData.Aud_Type = 1
Else
    If M16B.value = 1 Then
        ConfigData.Aud_Type = 2
    Else
        ConfigData.Aud_Type = 2     'default
    End If
End If

If MM.value = 1 Then        '1=mono     2=stereo (default)
    ConfigData.Aud_Cual = 1
Else
    If Ms.value = 1 Then
        ConfigData.Aud_Cual = 2
    Else
        ConfigData.Aud_Cual = 2     'default
    End If
End If

If MN.value = 1 Then    '1=normal   2=a3d   3=3d    4=ogg
    ConfigData.Aud_Mode = 1
Else
    If MA3d.value = 1 Then
        ConfigData.Aud_Mode = 2
    Else
        If M3d.value = 1 Then
            ConfigData.Aud_Mode = 3
        Else
            If Mogg.value = 1 Then
                ConfigData.Aud_Mode = 4
            Else
                ConfigData.Aud_Mode = 1     'default
            End If
        End If
    End If
End If

If RN.value = 1 Then    '1=normal ramping   2=sensitive ramping
    ConfigData.Aud_Mod_Type = 1
Else
    If RS.value = 1 Then
        ConfigData.Aud_Mod_Type = 2
    Else
        ConfigData.Aud_Mod_Type = 1     'default
    End If
End If

ConfigData.Aud_Mod_Cual = SS.value

If RFt2.value = 1 Then      '1=as ft2       2=as pt2
    ConfigData.Aud_Mod_Mode = 1
Else
    If RPt2.value = 1 Then
        ConfigData.Aud_Mod_Mode = 2
    Else
        ConfigData.Aud_Mod_Mode = 1     'default
    End If
End If

If Tn.value = 1 Then    '1=time normal      2=time rest
    ConfigData.Aud_Disp_Time = 1
    TopMenu.LType.Caption = "Normal"
Else
    If Tr.value = 1 Then
        ConfigData.Aud_Disp_Time = 2
        TopMenu.LType.Caption = "Restante"
    Else
        ConfigData.Aud_Disp_Time = 1    'default
        TopMenu.LType.Caption = "Normal"
    End If
End If

If Ono.value = 1 Then   '1=ondas normal     2=ondas rest
    ConfigData.Aud_Disp_Wave = 1
    TopMenu.OType.Caption = "Normal"
Else
    If Ore.value = 1 Then
        ConfigData.Aud_Disp_Wave = 2
        TopMenu.OType.Caption = "Restante"
    Else
        ConfigData.Aud_Disp_Wave = 1    'default
        TopMenu.OType.Caption = "Normal"
    End If
End If

If Sn.value = 1 Then    '1=samples norm.    2=samples rest
    ConfigData.Aud_Disp_Samp = 1
    TopMenu.SType.Caption = "Normal"
Else
    If Sr.value = 1 Then
        ConfigData.Aud_Disp_Samp = 2
        TopMenu.SType.Caption = "Restante"
    Else
        ConfigData.Aud_Disp_Samp = 1    'default
        TopMenu.SType.Caption = "Normal"
    End If
End If

ConfigData.Aud_Show_MiniRM = MMRm.value
ConfigData.Aud_Show_FTT = MVFFT.value
ConfigData.Aud_Show_SCOPE = MVSCOPE.value

'DIRECTORY OPTIONS
ConfigData.Dir_Tem = SetCipherConfigData(Trim(Tx(0).Text))
ConfigData.Dir_Com = SetCipherConfigData(Trim(Tx(1).Text))
ConfigData.Dir_Inst = SetCipherConfigData(Trim(Tx(2).Text))
ConfigData.Dir_Hor = SetCipherConfigData(Trim(Tx(3).Text))

'SECURITY OPTIONS
If None.value = 1 Then      '1=none     2=password      3=deneid access
    ConfigData.Sec_Type = 1
Else
    If Pass.value = 1 Then
        ConfigData.Sec_Type = 2
    Else
        If Den.value = 1 Then
            ConfigData.Sec_Type = 3
        Else
            ConfigData.Sec_Type = 1     'default
        End If
    End If
End If

ConfigData.Sec_Est12_1 = e1.value
ConfigData.Sec_Est12_2 = e2.value
ConfigData.Sec_Est12_3 = e3.value
ConfigData.Sec_Tnd12_1 = t1.value
ConfigData.Sec_Tnd12_2 = t2.value
ConfigData.Sec_Tnd12_3 = t3.value
ConfigData.Sec_Prg_1 = p1.value
ConfigData.Sec_Prg_2 = p2.value
ConfigData.Sec_Prg_3 = p3.value
ConfigData.Sec_Esp_1 = c1.value
ConfigData.Sec_Esp_2 = c2.value
ConfigData.Sec_Esp_3 = c3.value
ConfigData.Sec_Esp_4 = c4.value
ConfigData.Sec_Esp_5 = c5.value

'guardamos los datos en el archivo de configuracion
Result = SaveConfigFile(ConfigData)
If Result = "NotOk" Then
    'MsgBox "Errrorrrrrrrrrrrrr"
Else
    'xxx
End If

'unload the config window
Unload Me

End Sub

Private Sub CmdCancel_Click()

Unload Me

End Sub

Private Sub Den_Click()

If Den.value = 1 Then
    None.value = 0
    Pass.value = 0
    'habilitaciones
    e1.Enabled = True: e2.Enabled = True: e3.Enabled = True
    t1.Enabled = True: t2.Enabled = True: t3.Enabled = True
    p1.Enabled = True: p2.Enabled = True: p3.Enabled = True
    c1.Enabled = True: c2.Enabled = True: c3.Enabled = True
    c4.Enabled = True: c5.Enabled = True
    'habilitacion definir password
    DefPass.Enabled = False
Else
    If None.value = 0 Then
        If Pass.value = 0 Then
            None.value = 1
        End If
    End If
End If

End Sub

Private Sub EdSearch_Click()

Dim PRGName As String

On Error Resume Next
CmdPrg.InitDir = App.Path
CmdPrg.Filter = "Archivo de Programas (*.exe)|*.exe|Archivos de Programa"
CmdPrg.DialogTitle = "RM100 - Seleccione archivo de programa"
CmdPrg.CancelError = True
CmdPrg.ShowOpen

If err.Number = 32755 Then Exit Sub

PRGName = CmdPrg.filename
EdName.Text = PRGName

End Sub

Private Sub Exam_Click(index As Integer)

Dim Result As String

Result = BrowseForFolder("Seleccione la nueva carpeta.")

Tx(index).Text = Result & "\"

End Sub

Private Sub Form_Load()

'load some resource strings...
Me.Caption = LoadResString(2023)
SSTab1.Tab = 0
SSTab1.Caption = LoadResString(2024)
SSTab1.Tab = 1
SSTab1.Caption = LoadResString(2025)
SSTab1.Tab = 2
SSTab1.Caption = LoadResString(2026)
SSTab1.Tab = 3
SSTab1.Caption = LoadResString(2027)
SSTab1.Tab = 0

CmdAccept.Caption = LoadResString(2000)
CmdCancel.Caption = LoadResString(2001)
CmdAply.Caption = LoadResString(2003)

'paints some clocks displays...
Call Paint_Clocks

'cargamos los datos guardados en el archivo de configuracion
Dim ConfigData As ConfigRecord   'registros de Configuracion
ConfigData = OpenConfigFile

'////////////////////////////////// GENERAL OPTIONS
Etag.value = ConfigData.Gen_AutoTAG
EName.value = ConfigData.Gen_AutoName
ARep.value = ConfigData.Gen_ActiveReport
Rep1.value = ConfigData.Gen_ReportEst
Rep2.value = ConfigData.Gen_ReportTnd
RepAll.value = ConfigData.Gen_ReportAll
ARepName.Text = GetCipherConfigData(Trim(ConfigData.Gen_ReportProg))
EdName.Text = GetCipherConfigData(Trim(ConfigData.Gen_EditProg))
GrabName.Text = GetCipherConfigData(Trim(ConfigData.Gen_GrabProg))

If ARep.value = 0 Then
    Rep1.Enabled = False
    Rep2.Enabled = False
    RepAll.Enabled = False
Else
    If RepAll.value = 1 Then
        Rep1.Enabled = False
        Rep2.Enabled = False
    Else
        Rep1.Enabled = True
        Rep2.Enabled = True
    End If
End If

'////////////////////////////////// AUDIO OPTIONS
Select Case ConfigData.Aud_Type
    Case 1  '8bits
        M8b.value = 1
        M16B.value = 0
    Case Else  '16bits
        M8b.value = 0
        M16B.value = 1
End Select

Select Case ConfigData.Aud_Cual
    Case 1  'mono
        MM.value = 1
        Ms.value = 0
    Case Else  'stereo
        Ms.value = 1
        MM.value = 0
End Select

Select Case ConfigData.Aud_Mode
    Case 1  'normal
        MN.value = 1
        MA3d.value = 0
        M3d.value = 0
        Mogg.value = 0
    Case 2  'a3d
        MA3d.value = 1
        MN.value = 0
        M3d.value = 0
        Mogg.value = 0
    Case 3  '3d
        M3d.value = 1
        MN.value = 0
        MA3d.value = 0
        Mogg.value = 0
    Case 4  'ogg
        Mogg.value = 1
        MN.value = 0
        M3d.value = 0
        MA3d.value = 0
    Case Else   'default
        MN.value = 1
        MA3d.value = 0
        M3d.value = 0
        Mogg.value = 0
End Select

Select Case ConfigData.Aud_Mod_Type
    Case 1  'normal
    RN.value = 1
    RS.value = 0
    Case 2  'sensitive
    RN.value = 0
    RS.value = 1
    Case Else   'default
    RN.value = 1
    RS.value = 0
End Select

Select Case ConfigData.Aud_Mod_Mode
    Case 1  'as ft2
        RFt2.value = 1
        RPt2.value = 0
    Case 2  'as pt2
        RFt2.value = 0
        RPt2.value = 1
    Case Else   'default
        RFt2.value = 1
        RPt2.value = 0
End Select

SS.value = ConfigData.Aud_Mod_Cual

Select Case ConfigData.Aud_Disp_Time
    Case 1  'normal
        Tn.value = 1
        Tr.value = 0
    Case 2  'restante
        Tn.value = 0
        Tr.value = 1
    Case Else   'default
        Tn.value = 1
        Tr.value = 0
End Select

Select Case ConfigData.Aud_Disp_Wave
    Case 1  'normal
        Ono.value = 1
        Ore.value = 0
    Case 2  'restante
        Ono.value = 0
        Ore.value = 1
    Case Else   'default
        Ono.value = 1
        Ore.value = 0
End Select

Select Case ConfigData.Aud_Disp_Samp
    Case 1  'normal
        Sn.value = 1
        Sr.value = 0
    Case 2  'restante
        Sn.value = 0
        Sr.value = 1
    Case Else   'default
        Sn.value = 1
        Sr.value = 0
End Select

MMRm.value = ConfigData.Aud_Show_MiniRM
MVFFT.value = ConfigData.Aud_Show_FTT
MVSCOPE.value = ConfigData.Aud_Show_SCOPE

'////////////////////////////////// DIRECTORY OPTIONS
Tx(0).Text = GetCipherConfigData(Trim(ConfigData.Dir_Tem))
Tx(1).Text = GetCipherConfigData(Trim(ConfigData.Dir_Com))
Tx(2).Text = GetCipherConfigData(Trim(ConfigData.Dir_Inst))
Tx(3).Text = GetCipherConfigData(Trim(ConfigData.Dir_Hor))

'////////////////////////////////// SECURITY OPTIONS
Select Case ConfigData.Sec_Type
    Case 1  'none
        None.value = 1
        Pass.value = 0
        Den.value = 0
    Case 2  'password
        None.value = 0
        Pass.value = 1
        Den.value = 0
    Case 3  'deneid access
        None.value = 0
        Pass.value = 0
        Den.value = 1
    Case Else   'default
        None.value = 1
        Pass.value = 0
        Den.value = 0
End Select

e1.value = ConfigData.Sec_Est12_1
e2.value = ConfigData.Sec_Est12_2
e3.value = ConfigData.Sec_Est12_3
t1.value = ConfigData.Sec_Tnd12_1
t2.value = ConfigData.Sec_Tnd12_2
t3.value = ConfigData.Sec_Tnd12_3
p1.value = ConfigData.Sec_Prg_1
p2.value = ConfigData.Sec_Prg_2
p3.value = ConfigData.Sec_Prg_3
c1.value = ConfigData.Sec_Esp_1
c2.value = ConfigData.Sec_Esp_2
c3.value = ConfigData.Sec_Esp_3
c4.value = ConfigData.Sec_Esp_4
c5.value = ConfigData.Sec_Esp_5

End Sub

Private Sub GrabSearch_Click()

Dim PRGName As String

On Error Resume Next
CmdPrg.InitDir = App.Path
CmdPrg.Filter = "Archivo de Programas (*.exe)|*.exe|Archivos de Programa"
CmdPrg.DialogTitle = "RM100 - Seleccione archivo de programa"
CmdPrg.CancelError = True
CmdPrg.ShowOpen

If err.Number = 32755 Then Exit Sub

PRGName = CmdPrg.filename
GrabName.Text = PRGName

End Sub

Private Sub M16B_Click()

If M16B.value = 1 Then
    M8b.value = 0
Else
    If M8b.value = 0 Then
        M16B.value = 1
    End If
End If

End Sub

Private Sub M3d_Click()

If M3d.value = 1 Then
    MA3d.value = 0
    MN.value = 0
    Mogg.value = 0
Else
    If MA3d.value = 0 Then
        If MN.value = 0 Then
            If Mogg.value = 0 Then
                MN.value = 1
            End If
        End If
    End If
End If

End Sub

Private Sub M8b_Click()

If M8b.value = 1 Then
    M16B.value = 0
Else
    If M16B.value = 0 Then
        M16B.value = 1
    End If
End If

End Sub

Private Sub MA3d_Click()

If MA3d.value = 1 Then
    MN.value = 0
    M3d.value = 0
    Mogg.value = 0
Else
    If MN.value = 0 Then
        If M3d.value = 0 Then
            If Mogg.value = 0 Then
                MN.value = 1
            End If
        End If
    End If
End If

End Sub

Private Sub MM_Click()

If MM.value = 1 Then
    Ms.value = 0
Else
    If Ms.value = 0 Then
        Ms.value = 1
    End If
End If

End Sub

Private Sub MN_Click()

If MN.value = 1 Then
    MA3d.value = 0
    M3d.value = 0
    Mogg.value = 0
Else
    If MA3d.value = 0 Then
        If M3d.value = 0 Then
            If Mogg.value = 0 Then
                MN.value = 1
            End If
        End If
    End If
End If

End Sub

Private Sub Mogg_Click()

If Mogg.value = 1 Then
    MA3d.value = 0
    M3d.value = 0
    MN.value = 0
Else
    If MA3d.value = 0 Then
        If M3d.value = 0 Then
            If MN.value = 0 Then
                MN.value = 1
            End If
        End If
    End If
End If

End Sub

Private Sub Ms_Click()

If Ms.value = 1 Then
    MM.value = 0
Else
    If MM.value = 0 Then
        Ms.value = 1
    End If
End If

End Sub

Private Sub None_Click()

If None.value = 1 Then
    Pass.value = 0
    Den.value = 0
    'deshabilitamos
    e1.Enabled = False: e2.Enabled = False: e3.Enabled = False
    t1.Enabled = False: t2.Enabled = False: t3.Enabled = False
    p1.Enabled = False: p2.Enabled = False: p3.Enabled = False
    c1.Enabled = False: c2.Enabled = False: c3.Enabled = False
    c4.Enabled = False: c5.Enabled = False
    'boton definir password
    DefPass.Enabled = False
Else
    If Pass.value = 0 Then
        If Den.value = 0 Then
            None.value = 1
        End If
    End If
End If

End Sub

Private Sub Ono_Click()

If Ono.value = 1 Then
    Ore.value = 0
Else
    If Ore.value = 0 Then
        Ono.value = 1
    End If
End If

End Sub

Private Sub Ore_Click()

If Ore.value = 1 Then
    Ono.value = 0
Else
    If Ono.value = 0 Then
        Ono.value = 1
    End If
End If

End Sub

Private Sub Pass_Click()

If Pass.value = 1 Then
    None.value = 0
    Den.value = 0
    'habilitaciones
    e1.Enabled = True: e2.Enabled = True: e3.Enabled = True
    t1.Enabled = True: t2.Enabled = True: t3.Enabled = True
    p1.Enabled = True: p2.Enabled = True: p3.Enabled = True
    c1.Enabled = True: c2.Enabled = True: c3.Enabled = True
    c4.Enabled = True: c5.Enabled = True
    'habilitacion definir password
    DefPass.Enabled = True
Else
    If None.value = 0 Then
        If Den.value = 0 Then
            None.value = 1
        End If
    End If
End If

End Sub

Private Sub RepAll_Click()

If RepAll.value = 1 Then
    Rep1.value = 0
    Rep2.value = 0
    Rep1.Enabled = False
    Rep2.Enabled = False
Else
    Rep1.Enabled = True
    Rep2.Enabled = True
End If

End Sub

Private Sub RFt2_Click()

If RFt2.value = 1 Then
    RPt2.value = 0
Else
    If RPt2.value = 0 Then
        RFt2.value = 1
    End If
End If

End Sub

Private Sub RN_Click()

If RN.value = 1 Then
    RS.value = 0
Else
    If RS.value = 0 Then
        RN.value = 1
    End If
End If

End Sub

Private Sub RPt2_Click()

If RPt2.value = 1 Then
    RFt2.value = 0
Else
    If RFt2.value = 0 Then
        RFt2.value = 1
    End If
End If

End Sub

Private Sub RS_Click()

If RS.value = 1 Then
    RN.value = 0
Else
    If RN.value = 0 Then
        RN.value = 1
    End If
End If

End Sub

Private Sub Sn_Click()

If Sn.value = 1 Then
    Sr.value = 0
Else
    If Sr.value = 0 Then
        Sn.value = 1
    End If
End If

End Sub

Private Sub Sr_Click()

If Sr.value = 1 Then
    Sn.value = 0
Else
    If Sn.value = 0 Then
        Sr.value = 1
    End If
End If

End Sub

Private Sub Tn_Click()

If Tn.value = 1 Then
    Tr.value = 0
Else
    If Tr.value = 0 Then
        Tn.value = 1
    End If
End If

End Sub

Private Sub Tr_Click()

If Tr.value = 1 Then
    Tn.value = 0
Else
    If Tn.value = 0 Then
        Tn.value = 1
    End If
End If

End Sub

Private Sub TxTem_Click(index As Integer)

End Sub
