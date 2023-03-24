VERSION 5.00
Begin VB.UserControl ctlWindow 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ControlContainer=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3900
   ScaleWidth      =   4815
   Begin VB.Image ImageTask_TitleBarButton 
      Height          =   210
      Index           =   4
      Left            =   45
      Picture         =   "Window.ctx":0000
      Top             =   2925
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image ImageTask_TitleBarButton 
      Height          =   270
      Index           =   2
      Left            =   45
      Picture         =   "Window.ctx":02AA
      Top             =   2160
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image ImageTask_TitleBarButton 
      Height          =   270
      Index           =   3
      Left            =   600
      Picture         =   "Window.ctx":06DC
      Top             =   2160
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image ImageTask_TitleBarButton 
      Height          =   270
      Index           =   1
      Left            =   600
      Picture         =   "Window.ctx":0B0E
      Top             =   1440
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image ImageTask_TitleBarButton 
      Height          =   270
      Index           =   0
      Left            =   45
      Picture         =   "Window.ctx":0F40
      Top             =   1440
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label LabelTask_Title 
      BackStyle       =   0  'Transparent
      Caption         =   "WindowTitle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   3420
      Width           =   1410
   End
   Begin VB.Label LabelTask_Title 
      BackStyle       =   0  'Transparent
      Caption         =   "WindowTitle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   165
      TabIndex        =   1
      Top             =   3435
      Width           =   1365
   End
   Begin VB.Image ImageTask_Icon_Restore 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   2325
      Picture         =   "Window.ctx":1372
      Top             =   660
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image ImageTask_Icon_OnTop 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   210
      Index           =   3
      Left            =   3495
      Picture         =   "Window.ctx":17A4
      Top             =   2550
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image ImageTask_Icon_OnTop 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   210
      Index           =   2
      Left            =   2895
      Picture         =   "Window.ctx":1A4E
      Top             =   2550
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image ImageTask_Icon_Restore 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   3495
      Picture         =   "Window.ctx":1CF8
      Top             =   660
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image ImageTask_Icon_Restore 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   2895
      Picture         =   "Window.ctx":212A
      Top             =   660
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image ImageTask_Icon_OnTop 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   210
      Index           =   1
      Left            =   2325
      Picture         =   "Window.ctx":255C
      Top             =   2550
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image ImageTask_Icon_OnTop 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   210
      Index           =   0
      Left            =   1785
      Picture         =   "Window.ctx":2806
      Top             =   2550
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image ImageTask_Icon_Help 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   1785
      Picture         =   "Window.ctx":2AB0
      Top             =   1950
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image ImageTask_Icon_Minimize 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   1785
      Picture         =   "Window.ctx":2EE2
      Top             =   1245
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image ImageTask_Icon_Restore 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   1785
      Picture         =   "Window.ctx":3314
      Top             =   660
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image ImageTask_Icon_Close 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   1785
      Picture         =   "Window.ctx":3746
      Top             =   75
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image ImageTask_Icon_Close 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   2235
      Picture         =   "Window.ctx":3B78
      Top             =   75
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image ImageTask_Icon_Minimize 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   2325
      Picture         =   "Window.ctx":3FAA
      Top             =   1245
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image ImageTask_Icon_Help 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   2325
      Picture         =   "Window.ctx":43DC
      Top             =   1950
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image ImageTask_Window 
      Appearance      =   0  'Flat
      Height          =   30
      Index           =   7
      Left            =   1080
      Picture         =   "Window.ctx":480E
      Top             =   930
      Width           =   30
   End
   Begin VB.Image ImageTask_Window 
      Appearance      =   0  'Flat
      Height          =   30
      Index           =   6
      Left            =   135
      Picture         =   "Window.ctx":4860
      Top             =   930
      Width           =   180
   End
   Begin VB.Image ImageTask_Window 
      Appearance      =   0  'Flat
      Height          =   30
      Index           =   5
      Left            =   0
      Picture         =   "Window.ctx":48EA
      Top             =   930
      Width           =   30
   End
   Begin VB.Image ImageTask_Window 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   4
      Left            =   1080
      Picture         =   "Window.ctx":493C
      Top             =   465
      Width           =   30
   End
   Begin VB.Image ImageTask_Window 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   3
      Left            =   0
      Picture         =   "Window.ctx":4A3E
      Top             =   450
      Width           =   30
   End
   Begin VB.Image ImageTask_Window 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   2
      Left            =   1080
      Picture         =   "Window.ctx":4B40
      Top             =   0
      Width           =   75
   End
   Begin VB.Image ImageTask_Window 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   0
      Left            =   0
      Picture         =   "Window.ctx":4D12
      Top             =   0
      Width           =   75
   End
   Begin VB.Image ImageTask_Window 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   1
      Left            =   135
      Picture         =   "Window.ctx":4EE4
      Top             =   0
      Width           =   675
   End
   Begin VB.Image ImageTask_Window 
      Height          =   255
      Index           =   8
      Left            =   135
      Top             =   495
      Width           =   255
   End
End
Attribute VB_Name = "ctlWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ctlWindow.ctl - Modulo para Skins en Ventanas
'
' ctlWindow version 1.0 - junio 13, 2008
'
' Autor y mantenedor:
'     Abraham Araon Herrera Leiva - abrahamaraon@hotmail.com
'
'
' NOTICIA DE DERECHOS DE COPIA, LIMITACION, y LICENCIA:
'
' Si tu modificas este modulo puedes insertar noticias adicionales inmediatamente
' siguiendo esta sentencia.
'
' ctlWindow version 1.0, junio 13, 2008, Abraham Araon Herrera Leiva.
' Creacion del modulo.
'
' Este modulo es destribuido "TAL COMO ESTA". El autor rechaza toda garantia,
' expresada o implicada, incluyendo, sin limitacion, las garantias de
' comerciabilidad y de venta por cualquier proposito. El autor no asume
' obligacion por daño directo, indirecto, accidental, especial, ejemplar,
' o consequencial, los cuales puedan resultar por el uso de este modulo.
'
' Los permisos estan por la presente consecionados a usar, copiar, modificar,
' y distribuir este codigo fuente, o porciones de el, para cualquier proposito,
' sin temor, sujeto a las siguientes restricciones:
'
' 1. El origen de este codigo fuente no debe ser mal representado.
'
' 2. Versiones alteradas deben ser explicitamente marcadas como tal y
'    no deben ser mal representadas como siendo la fuente original.
'
' 3. Esta noticia de derechos de copia no debe ser removida o alterada desde
'    cualquier fuente o ser alterada de la fuente de distribucion.
'
' El autor especificamente permite y anima el uso de este codigo
' fuente como un componente para soporte en productos comerciales. Si tu usas
' este codigo fuente en un producto, un reconocimiento es muy bien apreciado.
'
'
Option Explicit

Public Enum TASK_STATE
   EXPANDED = True
   COLLAPSED = False
End Enum

Public Enum TransparentOpaque
   TRANSPARENT = 0
   OPAQUE = 1
End Enum

Public Enum TYPE_SIZE
   TS_LEFT = &HF001&
   TS_RIGHT = &HF002&
   TS_TOP = &HF003&
   TS_LEFT_TOP = &HF004&
   TS_RIGHT_TOP = &HF005&
   TS_BOTTOM = &HF006&
   TS_LEFT_BOTTOM = &HF007&
   TS_RIGHT_BOTTOM = &HF008&
End Enum

Dim WindowCloseButton As Boolean
Dim WindowMinimizeButton As Boolean
Dim WindowRestoreButton As Boolean
Dim WindowHelpButton As Boolean
Dim WindowOnTopButton As Boolean
Dim WindowSizable As Boolean

Dim WindowState As Boolean
Dim WindowTopState As Boolean
Dim WindowWidth As Long
Dim WindowHeigth As Long
Dim WindowRgn As String
Dim WindowOldHeigth As Long
Dim WindowBackStyle As TransparentOpaque
'
'Default Property Values:
Const m_def_State = True
Const m_def_CloseButton = True
Const m_def_MinimizeButton = True
Const m_def_RestoreButton = True
Const m_def_HelpButton = True
Const m_def_OnTopButton = True
Const m_def_MostTop = False
Const m_def_WindowSizable = False

'Event Declarations:
Event WindowResize()
Event CloseClick()
Event RestoreClick()
Event MinimizeClick()
Event HelpClick()
Event OnTopClick()

Event OnExpand()
Event OnCollapse()

Event BarTitleClick()
Event BarTitleDblClick()
Event BarTitleMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event BarTitleMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event BarTitleMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Event WindowMarkLeftMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event WindowMarkRightMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event WindowMarkBottomMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event WindowMarkLeftBottomMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event WindowMarkRightBottomMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Event WindowMarkLeftMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event WindowMarkRightMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event WindowMarkBottomMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event WindowMarkLeftBottomMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event WindowMarkRightBottomMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Event WindowMarkLeftMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event WindowMarkRightMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event WindowMarkLeftBottomMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event WindowMarkBottomMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event WindowMarkRightBottomMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const HWND_TOPMOST = -1
Private Const HWND_TOP = 0
Private Const HWND_NOTOPMOST = -2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_SYSCOMMAND = &H112
Private Const TITLE_BUTTON_SPACE = 30

Private buildingWindowRgn As Boolean

Public Sub startMove()
   ReleaseCapture
   SendMessage UserControl.Parent.hwnd, WM_NCLBUTTONDOWN, 2&, 0&
End Sub

Public Sub startSize(ByVal SIDE As TYPE_SIZE)
   ReleaseCapture
   SendMessage UserControl.Parent.hwnd, WM_SYSCOMMAND, SIDE, 0&
End Sub

Public Sub popUpSystemMenu()
   ' Send (ALT + SPACE) keys
   SendKeys "% "
End Sub

Public Sub SetStyle()
    On Error GoTo Err
    Dim lCurrentSettings As Long
    Const WS_MINIMIZEBOX = &H20000
    Const WS_MAXIMIZEBOX = &H10000
    Const WS_THICKFRAME = &H40000
    Const WS_DLGFRAME = &H400000
    Const WS_CAPTION = &HC00000
    Const GWL_STYLE = (-16)
    Const WS_SYSMENU = &H80000
    
    lCurrentSettings = GetWindowLong(UserControl.Parent.hwnd, GWL_STYLE)
    lCurrentSettings = lCurrentSettings And Not WS_THICKFRAME
    lCurrentSettings = lCurrentSettings And Not WS_DLGFRAME
    lCurrentSettings = lCurrentSettings And Not WS_CAPTION
    lCurrentSettings = lCurrentSettings And Not WS_MINIMIZEBOX
    lCurrentSettings = lCurrentSettings And Not WS_MAXIMIZEBOX
    lCurrentSettings = lCurrentSettings Or WS_SYSMENU
    
    SetWindowLong UserControl.Parent.hwnd, GWL_STYLE, lCurrentSettings
    Call destructWindowRgn
Err:
End Sub

Public Sub destructWindowRgn()
   SetWindowRgn UserControl.Parent.hwnd, 0&, True
End Sub

Public Sub buildWindowRgn()
   Dim buffer As String
   Dim lBuff As Integer, rBuff As Integer
   Dim X As Long, Y As Long
   Dim hRgn As Long, hRgntmp As Long
   Dim i As Long, j As Long
   Const RGN_AND = 1
   Const RGN_COPY = 5
   Const RGN_DIFF = 4
   Const RGN_OR = 2
   Const RGN_XOR = 3
   
   If Not (buildingWindowRgn) Then
      buildingWindowRgn = True
      
      X = UserControl.Parent.Width / Screen.TwipsPerPixelX   'Registers the size of the
      Y = UserControl.Parent.Height / Screen.TwipsPerPixelY  'form in pixels
      
      
      If (WindowRgn = "") Then
         Call destructWindowRgn
      Else
         hRgn = CreateRectRgn(0, 0, 0, 0)
         i = 1
         Do While (i < Len(WindowRgn))
            buffer = VBA.Mid$(WindowRgn, i, InStr(i, WindowRgn, "|", vbBinaryCompare) - i)
            If (IsNumeric(VBA.Left$(buffer, 1))) Then
               lBuff = CInt(VBA.Mid$(buffer, 1, InStr(1, buffer, ",", vbBinaryCompare)))
               rBuff = CInt(VBA.Mid$(buffer, InStr(1, buffer, ",", vbBinaryCompare) + 1))
               If (hRgn <> 0) Then
                  hRgntmp = CreateRectRgn(lBuff, j, X - rBuff, j + 1)
                  CombineRgn hRgn, hRgn, hRgntmp, RGN_OR
                  DeleteObject hRgntmp
               End If
               j = j + 1
            ElseIf (VBA.UCase$(VBA.Left$(buffer, 3)) = "ALL") Then
               lBuff = CInt(VBA.Mid$((VBA.Mid$(buffer, InStr(1, buffer, ";", vbBinaryCompare) + 1)), 1, InStr(1, (VBA.Mid$(buffer, InStr(1, buffer, ";", vbBinaryCompare) + 1)), ",", vbBinaryCompare)))
               rBuff = CInt(VBA.Mid$((VBA.Mid$(buffer, InStr(1, buffer, ";", vbBinaryCompare) + 1)), InStr(1, (VBA.Mid$(buffer, InStr(1, buffer, ";", vbBinaryCompare) + 1)), ",", vbBinaryCompare) + 1))
               buffer = (VBA.Mid$(buffer, 4, InStr(1, buffer, ";", vbBinaryCompare) - 4))
               buffer = miTrim(buffer)
               If Not (IsNumeric(buffer)) Then buffer = "0"
               If (hRgn <> 0) Then
                  hRgntmp = CreateRectRgn(lBuff, j, X - rBuff, Y + CLng(buffer))
                  CombineRgn hRgn, hRgn, hRgntmp, RGN_OR
                  DeleteObject hRgntmp
               End If
               j = Y + CLng(buffer)
            End If
            i = InStr(i + 1, WindowRgn, "|", vbBinaryCompare) + 1
         Loop
         SetWindowRgn UserControl.Parent.hwnd, hRgn, True
         DeleteObject hRgn
      End If
   End If
   
   buildingWindowRgn = False
End Sub

Public Function LoadSkin(Optional ByVal filename As String) As Boolean
   Dim buffer As String
   Dim filePath As String
   Dim fp As Long
   Dim i As Long
   
   LoadSkin = False
   If (Dir(filename, vbNormal Or vbArchive Or vbReadOnly Or vbHidden Or vbSystem) <> "") Then
      ' Load skin
      On Error GoTo skinError
      filePath = VBA.Mid$(filename, 1, VBA.InStrRev(filename, "\", -1, vbBinaryCompare))
      filename = VBA.Mid$(filename, VBA.InStrRev(filename, "\", -1, vbBinaryCompare) + 1)
      If (filePath = "") Then Exit Function
      If (filename = "") Then Exit Function
      fp = FreeFile
      Open filePath & filename For Input Access Read Lock Write As fp
         ' Clear images
         Call cleanPicture
         ' Clean Window Region
         Call destructWindowRgn: WindowRgn = ""
         Do Until (EOF(fp))
            Line Input #fp, buffer
            buffer = miTrim(buffer)
            Select Case VBA.UCase$(buffer)
               Case "#MAPS"
                  Do Until (EOF(fp)) Or (i >= 22)
                     Line Input #fp, buffer
                     buffer = miTrim(buffer)
                     If (VBA.Mid$(buffer, 1, 2) <> ":\") Then buffer = filePath & buffer
                     If (Dir(buffer, vbNormal Or vbArchive Or vbReadOnly Or vbHidden Or vbSystem) <> "") Then
                        If (i <= 7) Then
                           ImageTask_Window(i).Picture = LoadPicture(buffer)
                        ElseIf (i = 8) Then
                           ImageTask_Icon_Close(0).Picture = LoadPicture(buffer)
                           ImageTask_TitleBarButton(0) = ImageTask_Icon_Close(0)
                        ElseIf (i = 9) Then
                           ImageTask_Icon_Close(1).Picture = LoadPicture(buffer)
                        ElseIf (i = 10) Then
                           ImageTask_Icon_Restore(0).Picture = LoadPicture(buffer)
                           ImageTask_TitleBarButton(1) = ImageTask_Icon_Restore(0)
                        ElseIf (i = 11) Then
                           ImageTask_Icon_Restore(1).Picture = LoadPicture(buffer)
                        ElseIf (i = 12) Then
                           ImageTask_Icon_Restore(2).Picture = LoadPicture(buffer)
                        ElseIf (i = 13) Then
                           ImageTask_Icon_Restore(3).Picture = LoadPicture(buffer)
                        ElseIf (i = 14) Then
                           ImageTask_Icon_Minimize(0).Picture = LoadPicture(buffer)
                           ImageTask_TitleBarButton(2) = ImageTask_Icon_Minimize(0)
                        ElseIf (i = 15) Then
                           ImageTask_Icon_Minimize(1).Picture = LoadPicture(buffer)
                        ElseIf (i = 16) Then
                           ImageTask_Icon_Help(0).Picture = LoadPicture(buffer)
                           ImageTask_TitleBarButton(3) = ImageTask_Icon_Help(0)
                        ElseIf (i = 17) Then
                           ImageTask_Icon_Help(1).Picture = LoadPicture(buffer)
                        ElseIf (i = 18) Then
                           ImageTask_Icon_OnTop(0).Picture = LoadPicture(buffer)
                           ImageTask_TitleBarButton(4) = ImageTask_Icon_OnTop(0)
                        ElseIf (i = 19) Then
                           ImageTask_Icon_OnTop(1).Picture = LoadPicture(buffer)
                        ElseIf (i = 20) Then
                           ImageTask_Icon_OnTop(2).Picture = LoadPicture(buffer)
                        ElseIf (i = 21) Then
                           ImageTask_Icon_OnTop(3).Picture = LoadPicture(buffer)
                        End If
                        i = i + 1
                     End If
                  Loop
               Case "#TITLE-FONT"
                  i = 0
                  Do Until (EOF(fp))
                     Line Input #fp, buffer
                     buffer = miTrim(buffer)
                     If (i = 0) Then
                        Me.FontName = buffer
                     ElseIf (i = 1) Then
                        buffer = VBA.UCase$(buffer)
                        LabelTask_Title(0).FontBold = False: LabelTask_Title(1).FontBold = False
                        LabelTask_Title(0).FontItalic = False: LabelTask_Title(1).FontItalic = False
                        LabelTask_Title(0).FontUnderline = False: LabelTask_Title(1).FontUnderline = False
                        If (InStr(1, buffer, "NEGRITA", vbBinaryCompare) > 0) Or (InStr(1, buffer, "BOLD", vbBinaryCompare) > 0) Then
                           LabelTask_Title(0).FontBold = True: LabelTask_Title(1).FontBold = True
                        ElseIf (InStr(1, buffer, "ITALIC", vbBinaryCompare) > 0) Then
                           LabelTask_Title(0).FontItalic = True: LabelTask_Title(1).FontItalic = True
                        ElseIf (InStr(1, buffer, "UNDERLINE", vbBinaryCompare) > 0) Then
                           LabelTask_Title(0).FontUnderline = True: LabelTask_Title(1).FontUnderline = True
                        End If
                     Else
                        Me.FontSize = buffer
                        Exit Do
                     End If
                     i = i + 1
                  Loop
               Case "#TITLE-COLOR"
                  i = 0
                  Do Until (EOF(fp))
                     Line Input #fp, buffer
                     buffer = miTrim(buffer)
                     If (i = 0) Then
                        LabelTask_Title(0).ForeColor = buffer
                     Else
                        LabelTask_Title(1).ForeColor = buffer
                        Exit Do
                     End If
                     i = i + 1
                  Loop
               Case "#BACKGROUND-COLOR"
                  Do Until (EOF(fp))
                     Line Input #fp, buffer
                     buffer = miTrim(buffer)
                     Me.BackColor = buffer
                     Exit Do
                  Loop
               Case "#BACKGROUND-IMAGE"
                  Do Until (EOF(fp))
                     Line Input #fp, buffer
                     buffer = miTrim(buffer)
                     If (VBA.Mid$(buffer, 1, 2) <> ":\") Then buffer = filePath & buffer
                     Set ImageTask_Window(8).Picture = LoadPicture(buffer)
                     Exit Do
                  Loop
               Case "#WINDOW-RECT"
                  Do Until (EOF(fp))
                     Line Input #fp, buffer
                     buffer = miTrim(buffer)
                     If (buffer <> "#END") Then
                        WindowRgn = WindowRgn & buffer & "|"
                     Else
                        Exit Do
                     End If
                  Loop
               Case "#END"
                  Exit Do
            End Select
         Loop
      Close fp
      Call restorePictures
      Call sizeHeight(UserControl.Height)
      Call sizeWidth(UserControl.Width)
      Call sizeBackgroundPicture
      LoadSkin = True
   End If
   Exit Function
skinError:
   If (Err.Number = 53) Then ' No se ha encontrado el archivo
      Resume Next
   End If
End Function

Public Function setBackgroundPicture(ByVal filename As String) As Boolean
    setBackgroundPicture = False
    If (filename <> "") And (VBA.Dir(filename, vbArchive Or vbNormal Or vbReadOnly Or vbHidden Or vbSystem) <> "") Then
        On Error GoTo endSub
        Set ImageTask_Window(8).Picture = VB.LoadPicture(filename)
        Call sizeBackgroundPicture
        setBackgroundPicture = True
    ElseIf (filename = "") Then
        Set ImageTask_Window(8).Picture = Nothing
        setBackgroundPicture = True
    End If
endSub:
End Function

Private Sub setMarkCursorType()
   If (WindowSizable) Then
      ImageTask_Window(3).MousePointer = vbSizeWE
      ImageTask_Window(4).MousePointer = vbSizeWE
      ImageTask_Window(5).MousePointer = vbSizeNESW
      ImageTask_Window(6).MousePointer = vbSizeNS
      ImageTask_Window(7).MousePointer = vbSizeNWSE
   Else
      ImageTask_Window(3).MousePointer = vbDefault
      ImageTask_Window(4).MousePointer = vbDefault
      ImageTask_Window(5).MousePointer = vbDefault
      ImageTask_Window(6).MousePointer = vbDefault
      ImageTask_Window(7).MousePointer = vbDefault
   End If
End Sub

Private Function miTrim(ByVal szText As String) As String
   Dim i As Long
   i = 1
   Do While (i <= Len(szText))
      If (VBA.Mid$(szText, i, 1) = " ") Or (VBA.Mid$(szText, i, 1) = vbTab) Or (VBA.Mid$(szText, i, 1) = vbCr) Or (VBA.Mid$(szText, i, 1) = vbLf) Then
         szText = VBA.Right$(szText, Len(szText) - 1)
      ElseIf Not ((VBA.Mid$(szText, i, 1) = " ") Or (VBA.Mid$(szText, i, 1) = vbTab) Or (VBA.Mid$(szText, i, 1) = vbCr) Or (VBA.Mid$(szText, i, 1) = vbLf)) Then
         Exit Do
      Else
         i = i + 1
      End If
   Loop
   i = Len(szText)
   Do While (i > 0)
      If (VBA.Mid$(szText, i, 1) = " ") Or (VBA.Mid$(szText, i, 1) = vbTab) Or (VBA.Mid$(szText, i, 1) = vbCr) Or (VBA.Mid$(szText, i, 1) = vbLf) Then
         szText = VBA.Left$(szText, Len(szText) - 1)
         i = i - 1
      Else
         Exit Do
      End If
   Loop
   miTrim = szText
End Function

Private Sub changeState(ByVal State As Boolean)
   If (State = False) Then
      WindowOldHeigth = UserControl.Height
      Call sizeHeight(0)
      WindowState = False
      RaiseEvent OnCollapse
   Else
      Call sizeHeight(WindowOldHeigth)
      WindowState = True
      RaiseEvent OnExpand
   End If
End Sub

Private Sub moveTitleBarButtons()
   Dim middleHeight As Long
   Dim posLeft As Long
   
   posLeft = (ImageTask_Window(0).Width + ImageTask_Window(1).Width + ImageTask_Window(2).Width)
   posLeft = posLeft - TITLE_BUTTON_SPACE
   
   If (WindowCloseButton) Then
      middleHeight = (ImageTask_Window(2).Height / 2) - (ImageTask_TitleBarButton(0).Height / 2)
      posLeft = posLeft - ImageTask_TitleBarButton(0).Width - TITLE_BUTTON_SPACE
      ImageTask_TitleBarButton(0).Top = middleHeight
      ImageTask_TitleBarButton(0).Left = posLeft
      ImageTask_TitleBarButton(0).Visible = True
   Else
      ImageTask_TitleBarButton(0).Visible = False
   End If
   If (WindowRestoreButton) Then
      middleHeight = (ImageTask_Window(2).Height / 2) - (ImageTask_TitleBarButton(1).Height / 2)
      posLeft = posLeft - ImageTask_TitleBarButton(1).Width - TITLE_BUTTON_SPACE
      ImageTask_TitleBarButton(1).Top = middleHeight
      ImageTask_TitleBarButton(1).Left = posLeft
      ImageTask_TitleBarButton(1).Visible = True
   Else
      ImageTask_TitleBarButton(1).Visible = False
   End If
   If (WindowMinimizeButton) Then
      middleHeight = (ImageTask_Window(2).Height / 2) - (ImageTask_TitleBarButton(2).Height / 2)
      posLeft = posLeft - ImageTask_TitleBarButton(2).Width - TITLE_BUTTON_SPACE
      ImageTask_TitleBarButton(2).Top = middleHeight
      ImageTask_TitleBarButton(2).Left = posLeft
      ImageTask_TitleBarButton(2).Visible = True
   Else
      ImageTask_TitleBarButton(2).Visible = False
   End If
   If (WindowHelpButton) Then
      middleHeight = (ImageTask_Window(2).Height / 2) - (ImageTask_TitleBarButton(3).Height / 2)
      posLeft = posLeft - ImageTask_TitleBarButton(3).Width - TITLE_BUTTON_SPACE
      ImageTask_TitleBarButton(3).Top = middleHeight
      ImageTask_TitleBarButton(3).Left = posLeft
      ImageTask_TitleBarButton(3).Visible = True
   Else
      ImageTask_TitleBarButton(3).Visible = False
   End If
   If (WindowOnTopButton) Then
      middleHeight = (ImageTask_Window(2).Height / 2) - (ImageTask_TitleBarButton(4).Height / 2)
      posLeft = posLeft - ImageTask_TitleBarButton(4).Width - TITLE_BUTTON_SPACE
      ImageTask_TitleBarButton(4).Top = middleHeight
      ImageTask_TitleBarButton(4).Left = posLeft
      ImageTask_TitleBarButton(4).Visible = True
   Else
      ImageTask_TitleBarButton(4).Visible = False
   End If
End Sub

Private Function getTitleButtonsWidth() As Long
   getTitleButtonsWidth = TITLE_BUTTON_SPACE
   If (WindowCloseButton) Then
      getTitleButtonsWidth = getTitleButtonsWidth + ImageTask_TitleBarButton(0).Width + TITLE_BUTTON_SPACE
   End If
   If (WindowRestoreButton) Then
      getTitleButtonsWidth = getTitleButtonsWidth + ImageTask_TitleBarButton(1).Width + TITLE_BUTTON_SPACE
   End If
   If (WindowMinimizeButton) Then
      getTitleButtonsWidth = getTitleButtonsWidth + ImageTask_TitleBarButton(2).Width + TITLE_BUTTON_SPACE
   End If
   If (WindowHelpButton) Then
      getTitleButtonsWidth = getTitleButtonsWidth + ImageTask_TitleBarButton(3).Width + TITLE_BUTTON_SPACE
   End If
   If (WindowOnTopButton) Then
      getTitleButtonsWidth = getTitleButtonsWidth + ImageTask_TitleBarButton(4).Width + TITLE_BUTTON_SPACE
   End If
End Function

Private Function getWindowMinimeHeight() As Long
   getWindowMinimeHeight = IIf(ImageTask_Window(0).Height > ImageTask_Window(2).Height, ImageTask_Window(0).Height, ImageTask_Window(2).Height)
   getWindowMinimeHeight = getWindowMinimeHeight + IIf(ImageTask_Window(5).Height > ImageTask_Window(7).Height, ImageTask_Window(5).Height, ImageTask_Window(7).Height)
End Function

Private Function getWindowMinimeWidth() As Long
   getWindowMinimeWidth = IIf(ImageTask_Window(0).Width > ImageTask_Window(5).Width, ImageTask_Window(0).Width, ImageTask_Window(5).Width)
   getWindowMinimeWidth = getWindowMinimeWidth + IIf(ImageTask_Window(2).Width > ImageTask_Window(7).Width, ImageTask_Window(2).Width, ImageTask_Window(7).Width)
   getWindowMinimeWidth = getWindowMinimeWidth + getTitleButtonsWidth()
End Function

Private Sub sizeHeight(ByVal Height As Long)
   If (Height < getWindowMinimeHeight()) Then
      Height = getWindowMinimeHeight()
   End If
   
   WindowHeigth = Height
   
   If UserControl.Height <> Height Then
      UserControl.Height = Height
   End If
      
   'ImageTask_Window(0).Top = 0
   'ImageTask_Window(1).Top = 0
   'ImageTask_Window(2).Top = 0
   ImageTask_Window(3).Top = ImageTask_Window(0).Height
   ImageTask_Window(3).Height = Height - ImageTask_Window(0).Height - ImageTask_Window(5).Height
   ImageTask_Window(4).Top = ImageTask_Window(2).Height
   ImageTask_Window(4).Height = Height - ImageTask_Window(2).Height - ImageTask_Window(7).Height
   ImageTask_Window(5).Top = Height - ImageTask_Window(5).Height
   ImageTask_Window(6).Top = Height - ImageTask_Window(6).Height
   ImageTask_Window(7).Top = Height - ImageTask_Window(7).Height
   ImageTask_Window(8).Top = ImageTask_Window(1).Height
   ImageTask_Window(8).Height = ImageTask_Window(6).Top - ImageTask_Window(1).Height
      
   RaiseEvent WindowResize
End Sub

Private Sub sizeBackgroundPicture()
   ImageTask_Window(8).Top = ImageTask_Window(1).Height
   ImageTask_Window(8).Left = ImageTask_Window(3).Width
   ImageTask_Window(8).Height = ImageTask_Window(6).Top - ImageTask_Window(1).Height
   ImageTask_Window(8).Width = ImageTask_Window(4).Left - ImageTask_Window(3).Width
End Sub

Private Sub sizeWidth(ByVal Width As Long)
   If (Width < getWindowMinimeWidth()) Then
      Width = getWindowMinimeWidth()
   End If
   
   WindowWidth = Width
      
   If UserControl.Width <> Width Then
      UserControl.Width = Width
   End If
   
   'ImageTask_Window(0).Left = 0
   ImageTask_Window(1).Left = ImageTask_Window(0).Width
   ImageTask_Window(1).Width = Width - ImageTask_Window(0).Width - ImageTask_Window(2).Width
   ImageTask_Window(2).Left = Width - ImageTask_Window(2).Width
   'ImageTask_Window(3).Left = 0
   ImageTask_Window(4).Left = Width - ImageTask_Window(4).Width
   'ImageTask_Window(5).Left = 0
   ImageTask_Window(6).Left = ImageTask_Window(5).Left + ImageTask_Window(5).Width
   ImageTask_Window(6).Width = Width - ImageTask_Window(5).Width - ImageTask_Window(7).Width
   ImageTask_Window(7).Left = Width - ImageTask_Window(7).Width
   ImageTask_Window(8).Left = ImageTask_Window(3).Width
   ImageTask_Window(8).Width = ImageTask_Window(4).Left - ImageTask_Window(3).Width
   
   Call moveTitleBarButtons
   Call moveTitleBarCaption
      
   RaiseEvent WindowResize
End Sub

Private Sub moveTitleBarCaption()
   Const LT_LEFT = 90&
   LabelTask_Title(0).Top = (ImageTask_Window(1).Height / 2) - (LabelTask_Title(0).Height / 2)
   LabelTask_Title(1).Top = LabelTask_Title(0).Top + 15
   LabelTask_Title(0).Left = LT_LEFT
   LabelTask_Title(1).Left = LabelTask_Title(0).Left + 15
   LabelTask_Title(0).Height = 225
   LabelTask_Title(1).Height = 225
   
   LabelTask_Title(0).AutoSize = True
   LabelTask_Title(1).AutoSize = True
   LabelTask_Title(0).AutoSize = False
   LabelTask_Title(1).AutoSize = False
   
   If ((LabelTask_Title(0).Width + LT_LEFT) > (UserControl.Width - getTitleButtonsWidth())) Then
      LabelTask_Title(0).Width = UserControl.Width - getTitleButtonsWidth()
      LabelTask_Title(1).Width = UserControl.Width - getTitleButtonsWidth()
   End If
End Sub

Private Sub restorePictures()
   ImageTask_Window(0).Stretch = False
   ImageTask_Window(1).Stretch = False
   ImageTask_Window(2).Stretch = False
   ImageTask_Window(3).Stretch = False
   ImageTask_Window(4).Stretch = False
   ImageTask_Window(5).Stretch = False
   ImageTask_Window(6).Stretch = False
   ImageTask_Window(7).Stretch = False
   ImageTask_Window(8).Stretch = False
   ImageTask_TitleBarButton(0).Stretch = False
   ImageTask_TitleBarButton(1).Stretch = False
   ImageTask_TitleBarButton(2).Stretch = False
   ImageTask_TitleBarButton(3).Stretch = False
   ImageTask_TitleBarButton(4).Stretch = False
   'ImageTask_Window(0).Stretch = True
   ImageTask_Window(1).Stretch = True
   'ImageTask_Window(2).Stretch = True
   ImageTask_Window(3).Stretch = True
   ImageTask_Window(4).Stretch = True
   'ImageTask_Window(5).Stretch = True
   ImageTask_Window(6).Stretch = True
   'ImageTask_Window(7).Stretch = True
   ImageTask_Window(8).Stretch = True
   ImageTask_Window(0).Top = 0
   ImageTask_Window(1).Top = 0
   ImageTask_Window(2).Top = 0
   ImageTask_Window(0).Left = 0
   ImageTask_Window(3).Left = 0
   ImageTask_Window(5).Left = 0
End Sub

Private Sub cleanPicture()
    Set ImageTask_Window(0).Picture = Nothing
    Set ImageTask_Window(1).Picture = Nothing
    Set ImageTask_Window(2).Picture = Nothing
    Set ImageTask_Window(3).Picture = Nothing
    Set ImageTask_Window(4).Picture = Nothing
    Set ImageTask_Window(5).Picture = Nothing
    Set ImageTask_Window(6).Picture = Nothing
    Set ImageTask_Window(7).Picture = Nothing
    Set ImageTask_Window(8).Picture = Nothing
    Set ImageTask_TitleBarButton(0).Picture = Nothing
    Set ImageTask_TitleBarButton(1).Picture = Nothing
    Set ImageTask_TitleBarButton(2).Picture = Nothing
    Set ImageTask_TitleBarButton(3).Picture = Nothing
    Set ImageTask_TitleBarButton(4).Picture = Nothing
    Set ImageTask_Icon_Close(0).Picture = Nothing
    Set ImageTask_Icon_Close(1).Picture = Nothing
    Set ImageTask_Icon_Restore(0).Picture = Nothing
    Set ImageTask_Icon_Restore(1).Picture = Nothing
    Set ImageTask_Icon_Restore(2).Picture = Nothing
    Set ImageTask_Icon_Restore(3).Picture = Nothing
    Set ImageTask_Icon_Minimize(0).Picture = Nothing
    Set ImageTask_Icon_Minimize(1).Picture = Nothing
    Set ImageTask_Icon_Help(0).Picture = Nothing
    Set ImageTask_Icon_Help(1).Picture = Nothing
    Set ImageTask_Icon_OnTop(0).Picture = Nothing
    Set ImageTask_Icon_OnTop(1).Picture = Nothing
    Set ImageTask_Icon_OnTop(2).Picture = Nothing
    Set ImageTask_Icon_OnTop(3).Picture = Nothing
End Sub

Public Sub Refresh()
   Call sizeHeight(UserControl.Height)
   Call sizeWidth(UserControl.Width)
   Call sizeBackgroundPicture
End Sub

'=============================================================================
'---                        Eventos de los controles                       ---
'=============================================================================

Private Sub ImageTask_Window_Click(index As Integer)
   If (index = 0) Or (index = 1) Or (index = 2) Then
      RaiseEvent BarTitleClick
   End If
End Sub

Private Sub ImageTask_Window_DblClick(index As Integer)
   If (index = 0) Or (index = 1) Or (index = 2) Then
      RaiseEvent BarTitleDblClick
   End If
End Sub

Private Sub ImageTask_TitleBarButton_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (index = 0) Then
      ' Cerrar
      ImageTask_TitleBarButton(0) = ImageTask_Icon_Close(1)
   ElseIf (index = 1) Then
      ' Restaurar
      If (UserControl.Parent.WindowState = vbNormal) Then
         ImageTask_TitleBarButton(1) = ImageTask_Icon_Restore(1)
      ElseIf (UserControl.Parent.WindowState = vbMaximized) Then
         ImageTask_TitleBarButton(1) = ImageTask_Icon_Restore(3)
      End If
   ElseIf (index = 2) Then
      ' Minimizar
      ImageTask_TitleBarButton(2) = ImageTask_Icon_Minimize(1)
   ElseIf (index = 3) Then
      ' Ayuda
      ImageTask_TitleBarButton(3) = ImageTask_Icon_Help(1)
   ElseIf (index = 4) Then
      ' OnTop
      If (WindowTopState = False) Then
         ImageTask_TitleBarButton(4) = ImageTask_Icon_OnTop(1)
      Else
         ImageTask_TitleBarButton(4) = ImageTask_Icon_OnTop(3)
      End If
   End If
End Sub

Private Sub ImageTask_TitleBarButton_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (index = 0) Then
      ' Cerrar
      ImageTask_TitleBarButton(0) = ImageTask_Icon_Close(0)
   ElseIf (index = 1) Then
      ' Restaurar
      If (UserControl.Parent.WindowState = vbNormal) Then
         ImageTask_TitleBarButton(1) = ImageTask_Icon_Restore(0)
      ElseIf (UserControl.Parent.WindowState = vbMaximized) Then
         ImageTask_TitleBarButton(1) = ImageTask_Icon_Restore(2)
      End If
   ElseIf (index = 2) Then
      ' Minimizar
      ImageTask_TitleBarButton(2) = ImageTask_Icon_Minimize(0)
   ElseIf (index = 3) Then
      ' Ayuda
      ImageTask_TitleBarButton(3) = ImageTask_Icon_Help(0)
   ElseIf (index = 4) Then
      ' OnTop
      If (WindowTopState = False) Then
         ImageTask_TitleBarButton(4) = ImageTask_Icon_OnTop(0)
      Else
         ImageTask_TitleBarButton(4) = ImageTask_Icon_OnTop(2)
      End If
   End If
End Sub

Private Sub ImageTask_TitleBarButton_Click(index As Integer)
   If (index = 0) Then
      ' Cerrar
      RaiseEvent CloseClick
   ElseIf (index = 1) Then
      ' Restaurar
      RaiseEvent RestoreClick
      If (UserControl.Parent.WindowState = vbNormal) Then
         ImageTask_TitleBarButton(1) = ImageTask_Icon_Restore(0)
      ElseIf (UserControl.Parent.WindowState = vbMaximized) Then
         ImageTask_TitleBarButton(1) = ImageTask_Icon_Restore(2)
      End If
   ElseIf (index = 2) Then
      ' Minimizar
      RaiseEvent MinimizeClick
   ElseIf (index = 3) Then
      ' Ayuda
      RaiseEvent HelpClick
   ElseIf (index = 4) Then
      ' OnTop
      RaiseEvent OnTopClick
      If (WindowTopState = False) Then
         ImageTask_TitleBarButton(4) = ImageTask_Icon_OnTop(0)
      Else
         ImageTask_TitleBarButton(4) = ImageTask_Icon_OnTop(2)
      End If
   End If
End Sub

Private Sub ImageTask_Window_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (index = 0) Or (index = 1) Or (index = 2) Then
      RaiseEvent BarTitleMouseMove(Button, Shift, X, Y)
   ElseIf (index = 3) Then
      RaiseEvent WindowMarkLeftMouseMove(Button, Shift, X, Y)
   ElseIf (index = 4) Then
      RaiseEvent WindowMarkRightMouseMove(Button, Shift, X, Y)
   ElseIf (index = 5) Then
      RaiseEvent WindowMarkLeftBottomMouseMove(Button, Shift, X, Y)
   ElseIf (index = 6) Then
      RaiseEvent WindowMarkBottomMouseMove(Button, Shift, X, Y)
   ElseIf (index = 7) Then
      RaiseEvent WindowMarkRightBottomMouseMove(Button, Shift, X, Y)
   End If
End Sub

Private Sub ImageTask_Window_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (index = 0) Or (index = 1) Or (index = 2) Then
      RaiseEvent BarTitleMouseDown(Button, Shift, X, Y)
   ElseIf (index = 3) Then
      RaiseEvent WindowMarkLeftMouseDown(Button, Shift, X, Y)
   ElseIf (index = 4) Then
      RaiseEvent WindowMarkRightMouseDown(Button, Shift, X, Y)
   ElseIf (index = 5) Then
      RaiseEvent WindowMarkLeftBottomMouseDown(Button, Shift, X, Y)
   ElseIf (index = 6) Then
      RaiseEvent WindowMarkBottomMouseDown(Button, Shift, X, Y)
   ElseIf (index = 7) Then
      RaiseEvent WindowMarkRightBottomMouseDown(Button, Shift, X, Y)
   End If
End Sub

Private Sub ImageTask_Window_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (index = 0) Or (index = 1) Or (index = 2) Then
      RaiseEvent BarTitleMouseUp(Button, Shift, X, Y)
   ElseIf (index = 3) Then
      RaiseEvent WindowMarkLeftMouseUp(Button, Shift, X, Y)
   ElseIf (index = 4) Then
      RaiseEvent WindowMarkRightMouseUp(Button, Shift, X, Y)
   ElseIf (index = 5) Then
      RaiseEvent WindowMarkLeftBottomMouseUp(Button, Shift, X, Y)
   ElseIf (index = 6) Then
      RaiseEvent WindowMarkBottomMouseUp(Button, Shift, X, Y)
   ElseIf (index = 7) Then
      RaiseEvent WindowMarkRightBottomMouseUp(Button, Shift, X, Y)
   End If
End Sub

Private Sub LabelTask_Title_Click(index As Integer)
   Call ImageTask_Window_Click(0)
End Sub

Private Sub LabelTask_Title_DblClick(index As Integer)
   Call ImageTask_Window_DblClick(0)
End Sub

Private Sub LabelTask_Title_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call ImageTask_Window_MouseDown(0, Button, Shift, X, Y)
End Sub

Private Sub LabelTask_Title_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call ImageTask_Window_MouseMove(0, Button, Shift, X, Y)
End Sub

Private Sub LabelTask_Title_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call ImageTask_Window_MouseUp(0, Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
   Call restorePictures
End Sub

Private Sub UserControl_Resize()
   On Error Resume Next
   If (UserControl.Width <> WindowWidth) Then
      Call sizeWidth(UserControl.Width)
   End If
   If (UserControl.Height <> WindowHeigth) Then
      Call sizeHeight(UserControl.Height)
   End If
   If (UserControl.Parent.WindowState = vbNormal) Then
      If (ImageTask_TitleBarButton(1).Picture.Handle <> ImageTask_Icon_Restore(0).Picture.Handle) Then
         ImageTask_TitleBarButton(1) = ImageTask_Icon_Restore(0)
      End If
   ElseIf (UserControl.Parent.WindowState = vbMaximized) Then
      If (ImageTask_TitleBarButton(1).Picture.Handle <> ImageTask_Icon_Restore(2).Picture.Handle) Then
         ImageTask_TitleBarButton(1) = ImageTask_Icon_Restore(2)
      End If
   End If
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=LabelTask(0),LabelTask_Title,0,Caption
Public Property Get Caption() As String
   Caption = LabelTask_Title(0).Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   LabelTask_Title(0).Caption() = New_Caption
   LabelTask_Title(1).Caption() = New_Caption
   LabelTask_Title(0).AutoSize = True: LabelTask_Title(0).AutoSize = False
   LabelTask_Title(1).AutoSize = True: LabelTask_Title(1).AutoSize = False
   Call moveTitleBarCaption
   PropertyChanged "Caption"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
   BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   UserControl.BackColor() = New_BackColor
   PropertyChanged "BackColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=LabelTask(0),LabelTask_Title,0,Font
Public Property Get Font() As Font
   Set Font = LabelTask_Title(0).Font
End Property

Public Property Set Font(ByVal New_Font As Font)
   Set LabelTask_Title(0).Font = New_Font
   Set LabelTask_Title(1).Font = New_Font
   PropertyChanged "Font"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=LabelTask(0),LabelTask_Title,0,FontBold
Public Property Get FontBold() As Boolean
   FontBold = LabelTask_Title(0).FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
   LabelTask_Title(0).FontBold() = New_FontBold
   LabelTask_Title(1).FontBold() = New_FontBold
   PropertyChanged "FontBold"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=LabelTask(0),LabelTask_Title,0,FontItalic
Public Property Get FontItalic() As Boolean
   FontItalic = LabelTask_Title(0).FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
   LabelTask_Title(0).FontItalic() = New_FontItalic
   LabelTask_Title(1).FontItalic() = New_FontItalic
   PropertyChanged "FontItalic"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=LabelTask(0),LabelTask_Title,0,FontName
Public Property Get FontName() As String
   FontName = LabelTask_Title(0).FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
   LabelTask_Title(0).FontName() = New_FontName
   LabelTask_Title(1).FontName() = New_FontName
   PropertyChanged "FontName"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=LabelTask(0),LabelTask_Title,0,FontSize
Public Property Get FontSize() As Single
   FontSize = LabelTask_Title(0).FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
   LabelTask_Title(0).FontSize() = New_FontSize
   LabelTask_Title(1).FontSize() = New_FontSize
   PropertyChanged "FontSize"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=LabelTask(0),LabelTask_Title,0,FontUnderline
Public Property Get FontUnderline() As Boolean
   FontUnderline = LabelTask_Title(0).FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
   LabelTask_Title(0).FontUnderline() = New_FontUnderline
   LabelTask_Title(1).FontUnderline() = New_FontUnderline
   PropertyChanged "FontUnderline"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=LabelTask(0),LabelTask_Title,0,ForeColor
Public Property Get ForeColor() As OLE_COLOR
   ForeColor = LabelTask_Title(0).ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
   LabelTask_Title(0).ForeColor() = New_ForeColor
   PropertyChanged "ForeColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,0
Public Property Get State() As TASK_STATE
Attribute State.VB_Description = "Devuelve o establece el estado expandido (True) o contractado (False)."
   State = WindowState
End Property

Public Property Let State(ByVal New_state As TASK_STATE)
   Call changeState(New_state)
   WindowState = New_state
   PropertyChanged "state"
End Property

Public Property Get hasDC() As Boolean
   hasDC = UserControl.hasDC
End Property

Public Property Get hDc() As Long
   hDc = UserControl.hDc
End Property

Public Property Get hwnd() As Long
   hwnd = UserControl.hwnd
End Property

Public Property Get CloseButton() As Boolean
   CloseButton = WindowCloseButton
End Property

Public Property Let CloseButton(ByVal New_CloseButton As Boolean)
   WindowCloseButton = New_CloseButton
   Call moveTitleBarButtons
   PropertyChanged "CloseButton"
End Property

Public Property Get MinimizeButton() As Boolean
   MinimizeButton = WindowMinimizeButton
End Property

Public Property Let MinimizeButton(ByVal New_MinimizeButton As Boolean)
   WindowMinimizeButton = New_MinimizeButton
   Call moveTitleBarButtons
   PropertyChanged "MinimizeButton"
End Property

Public Property Get RestoreButton() As Boolean
   RestoreButton = WindowRestoreButton
End Property

Public Property Let RestoreButton(ByVal New_RestoreButton As Boolean)
   WindowRestoreButton = New_RestoreButton
   Call moveTitleBarButtons
   PropertyChanged "RestoreButton"
End Property

Public Property Get HelpButton() As Boolean
   HelpButton = WindowHelpButton
End Property

Public Property Let HelpButton(ByVal New_HelpButton As Boolean)
   WindowHelpButton = New_HelpButton
   Call moveTitleBarButtons
   PropertyChanged "HelpButton"
End Property

Public Property Get OnTopButton() As Boolean
   OnTopButton = WindowOnTopButton
End Property

Public Property Let OnTopButton(ByVal New_OnTopButton As Boolean)
   WindowOnTopButton = New_OnTopButton
   Call moveTitleBarButtons
   PropertyChanged "OnTopButton"
End Property

Public Property Get MostTop() As Boolean
   MostTop = WindowTopState
End Property

Public Property Let MostTop(ByVal New_MostTop As Boolean)
   WindowTopState = New_MostTop
   If (WindowTopState) Then
      SetWindowPos UserControl.Parent.hwnd, HWND_TOPMOST, UserControl.Parent.Left / 15, UserControl.Parent.Top / 15, UserControl.Parent.Width / 15, UserControl.Parent.Height / 15, SWP_SHOWWINDOW
      ImageTask_TitleBarButton(4) = ImageTask_Icon_OnTop(2)
   Else
      SetWindowPos UserControl.Parent.hwnd, HWND_NOTOPMOST, UserControl.Parent.Left / 15, UserControl.Parent.Top / 15, UserControl.Parent.Width / 15, UserControl.Parent.Height / 15, SWP_NOACTIVATE
      ImageTask_TitleBarButton(4) = ImageTask_Icon_OnTop(0)
   End If
   PropertyChanged "MostTop"
End Property

Public Property Get MinimeHeight() As Long
   MinimeHeight = getWindowMinimeHeight()
End Property

Public Property Get MinimeWidth() As Long
   MinimeWidth = getWindowMinimeWidth()
End Property

Public Property Get Sizable() As Boolean
   Sizable = WindowSizable
End Property

Public Property Let Sizable(ByVal New_Sizable As Boolean)
   WindowSizable = New_Sizable
   Call setMarkCursorType
End Property

Public Property Get BackStyle() As TransparentOpaque
   BackStyle = WindowBackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As TransparentOpaque)
   WindowBackStyle = New_BackStyle
   UserControl.BackStyle = WindowBackStyle
   If (WindowBackStyle = TRANSPARENT) Then
      ImageTask_Window(8).Visible = False
   Else
      ImageTask_Window(8).Visible = True
   End If
End Property

Public Property Get backgroundTop() As Single
   backgroundTop = ImageTask_Window(8).Top
End Property

Public Property Get backgroundLeft() As Single
   backgroundLeft = ImageTask_Window(8).Left
End Property

Public Property Get backgroundHeight() As Single
   backgroundHeight = ImageTask_Window(8).Height
End Property

Public Property Let backgroundHeight(ByVal Height As Single)
   Call sizeHeight(ImageTask_Window(1).Height + Height + ImageTask_Window(6).Height)
End Property

Public Property Get backgroundWidth() As Single
   backgroundWidth = ImageTask_Window(8).Width
End Property

Public Property Let backgroundWidth(ByVal Width As Single)
   Call sizeWidth(ImageTask_Window(3).Width + Width + ImageTask_Window(4).Width)
End Property

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
   If (IsWindow(UserControl.Parent.hwnd) <> 1) Then
      Err.Raise 1, "InitProperties", "El valor hWnd de su formulario no es válido."
   End If
   WindowState = m_def_State
   WindowCloseButton = m_def_CloseButton
   WindowMinimizeButton = m_def_MinimizeButton
   WindowRestoreButton = m_def_RestoreButton
   WindowHelpButton = m_def_HelpButton
   WindowOnTopButton = m_def_OnTopButton
   WindowSizable = m_def_WindowSizable
   WindowBackStyle = OPAQUE
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   LabelTask_Title(0).Caption = PropBag.ReadProperty("Caption", "Window")
   LabelTask_Title(1).Caption = PropBag.ReadProperty("Caption", "Window")
   UserControl.BackColor = PropBag.ReadProperty("BackColor", &HF0F2F2)
   Set LabelTask_Title(0).Font = PropBag.ReadProperty("Font", Ambient.Font)
   LabelTask_Title(0).FontBold = PropBag.ReadProperty("FontBold", 1)
   LabelTask_Title(0).FontItalic = PropBag.ReadProperty("FontItalic", 0)
   LabelTask_Title(0).FontName = PropBag.ReadProperty("FontName", "Verdana")
   LabelTask_Title(0).FontSize = PropBag.ReadProperty("FontSize", 9)
   LabelTask_Title(0).FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
   LabelTask_Title(0).ForeColor = PropBag.ReadProperty("ForeColor", &H0)
   WindowState = PropBag.ReadProperty("State", m_def_State)
   WindowCloseButton = PropBag.ReadProperty("CloseButton", m_def_CloseButton)
   WindowMinimizeButton = PropBag.ReadProperty("MinimizeButton", m_def_MinimizeButton)
   WindowRestoreButton = PropBag.ReadProperty("RestoreButton", m_def_RestoreButton)
   WindowHelpButton = PropBag.ReadProperty("HelpButton", m_def_HelpButton)
   WindowOnTopButton = PropBag.ReadProperty("OnTopButton", m_def_OnTopButton)
   WindowTopState = PropBag.ReadProperty("MostTop", m_def_MostTop)
   WindowSizable = PropBag.ReadProperty("Sizable", m_def_WindowSizable)
   Me.Sizable = WindowSizable
   WindowBackStyle = PropBag.ReadProperty("BackStyle", OPAQUE)
   Me.BackStyle = WindowBackStyle
   If (WindowTopState) Then Me.MostTop = WindowTopState
   Call Refresh
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("Caption", LabelTask_Title(0).Caption, "Window")
   Call PropBag.WriteProperty("Caption", LabelTask_Title(1).Caption, "Window")
   Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HF0F2F2)
   Call PropBag.WriteProperty("Font", LabelTask_Title(0).Font, Ambient.Font)
   Call PropBag.WriteProperty("FontBold", LabelTask_Title(0).FontBold, 1)
   Call PropBag.WriteProperty("FontItalic", LabelTask_Title(0).FontItalic, 0)
   Call PropBag.WriteProperty("FontName", LabelTask_Title(0).FontName, "")
   Call PropBag.WriteProperty("FontSize", LabelTask_Title(0).FontSize, 0)
   Call PropBag.WriteProperty("FontUnderline", LabelTask_Title(0).FontUnderline, 0)
   Call PropBag.WriteProperty("ForeColor", LabelTask_Title(0).ForeColor, &H0&)
   Call PropBag.WriteProperty("State", WindowState, m_def_State)
   Call PropBag.WriteProperty("CloseButton", WindowCloseButton, m_def_CloseButton)
   Call PropBag.WriteProperty("MinimizeButton", WindowMinimizeButton, m_def_MinimizeButton)
   Call PropBag.WriteProperty("RestoreButton", WindowRestoreButton, m_def_RestoreButton)
   Call PropBag.WriteProperty("HelpButton", WindowHelpButton, m_def_HelpButton)
   Call PropBag.WriteProperty("OnTopButton", WindowOnTopButton, m_def_OnTopButton)
   Call PropBag.WriteProperty("MostTop", WindowTopState, m_def_MostTop)
   Call PropBag.WriteProperty("Sizable", WindowSizable, m_def_WindowSizable)
   Call PropBag.WriteProperty("BackStyle", WindowBackStyle, OPAQUE)
End Sub
