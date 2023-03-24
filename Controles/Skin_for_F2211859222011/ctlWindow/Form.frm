VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "SKIN"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NORMAL"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F8F1&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FD9D33&
      Height          =   2760
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin Proyecto1.ctlWindow Window 
      Height          =   405
      Left            =   4080
      Top             =   120
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   714
      Caption         =   "WindowTitle"
      Caption         =   "WindowTitle"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   9
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------------------------------------------------
' TEMPORAL CODE
'
' Please take this to help you to collect skins files.
'
' DELETE DELETE DELETE DELETE DELETE DELETE DELETE DELETE DELETE DELETE DELETE DE
'--------------------------------------------------------------------------------
        ' Var to keep all Skin files (full path)
        ' This is part of: getAllSkinsRecursive
        Dim skins As New Collection
'-----------------------------------------------

                        '
                        '// THESE ARE FOR SYSTEM SKIN OR OWN SKIN IMPLEMENTATION.
                        '// THIS IS UNOFICIAL BUT CAN HELP YOU IN SOME THING...
                        '
                        Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
                        Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
                        Private Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
                        Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
                        Const GWL_STYLE = (-16)
                        Const WS_CAPTION = &HC00000
                        Const WS_MAXIMIZEBOX = &H10000
                        Const WS_MINIMIZEBOX = &H20000
                        Const WS_OVERLAPPED = &H0&
                        Const WS_SYSMENU = &H80000
                        Const WS_THICKFRAME = &H40000
                        Const WS_BORDER = &H800000
                        Const WS_DLGFRAME = &H400000
                        Const WS_POPUP = &H80000000
                        Const WS_VISIBLE = &H10000000
                        Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
                        Const WS_CUSTOM = WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or _
                                          WS_OVERLAPPED Or WS_SYSMENU Or WS_THICKFRAME Or _
                                          WS_BORDER Or WS_DLGFRAME Or WS_POPUP Or WS_VISIBLE
                        '
                        '
                        '// SET SYSTEM STYLE TO WINDOW (Has Bugs, can't redraw properly)
                        Private Sub Command1_Click()
                            Window.Visible = False
                            SetWindowLong Me.hwnd, GWL_STYLE, GetWindowLong(Me.hwnd, GWL_STYLE) Or WS_CUSTOM
                            Window.destructWindowRgn
                            SetWindowRgn Me.hwnd, GetWindowRgn(Me.hwnd, 0), True
                        End Sub
                        '
                        '// SET OUR OWN STYLE TO WINDOW
                        Private Sub Command2_Click()
                            Window.SetStyle
                            Window.Visible = True
                            Window.buildWindowRgn
                        End Sub


'---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------

        '// GET ALL SKIN FILES ON APP.PATH AND SUB DIRS
        Private Sub Form_Activate()
            ' get all skin files on all sub directories inclusive
            getAllSkinsRecursive App.path, -1
            
            ' add skin list to List1 object
            Dim i As Integer, tmp As String
            For i = 1 To skins.Count
                ' cut only file and its folder (no complete path)
                tmp = Mid$(skins.Item(i), InStrRev(skins.Item(i), "\", InStrRev(skins.Item(i), "\", -1, vbBinaryCompare) - 1, vbBinaryCompare) + 1)
                List1.AddItem tmp
            Next i
        End Sub
        
        '// LOAD THE SELECTED SKIN NAME
        Private Sub List1_Click()
            Window.LoadSkin skins.Item(List1.ListIndex + 1)
            Me.BackColor = Window.BackColor
            Window.buildWindowRgn
        End Sub
        
        '// FUNCTION PARAMS:
        '                   path    -   directory where search.
        '                   subDirs -   deep level of sub-directories to process from path. This is;
        '                              -1: all sub-directories of each sub-directory of path
        '                               0: only current path.
        '                               1: 1 sub-directory of path (for each folder on path).
        '                               2: 2 sub-directories of path (for each folder on path) and so.
        '
        ' Feel free to modify this to you needs! (~may be this put on an Bas Module~).
        '
        Private Sub getAllSkinsRecursive(Optional path As String, Optional subDirs As Integer = 1)
            Dim file  As String
            Const SKIN_MAGIC = "#WINDOW-SKIN-1.0"
            ' if not before search
            If (path = "") Then
                ' set default path to search
                path = App.path
            End If
            ' put back-slash directory ("\") at end
            If (Right$(path, 1) <> "\") Then path = path & "\"
            ' init find
            file = Dir(path, vbArchive Or vbDirectory Or vbNormal Or vbReadOnly Or vbHidden Or vbSystem)
            Do
                If (file <> "") And (file <> ".") And (file <> "..") Then
                    ' if is dir AND you want do it "recursive"
                    'On Error Resume Next
                    If ((FileLen(path & file) <= 0) And (subDirs <> 0)) Then
                        ' get files of sub-dir
                        getAllSkinsRecursive path & file, subDirs - 1
                        ' restore Dir position
                        Dim tmp2 As String
                        tmp2 = Dir(path, vbArchive Or vbDirectory Or vbNormal Or vbReadOnly Or vbHidden Or vbSystem)
                        Do Until ((tmp2 = file) Or (tmp2 = ""))
                            tmp2 = Dir
                        Loop
                    ElseIf (FileLen(path & file) > 0) Then
                        ' verify if is an skin file
                        Dim fp As Integer, tmp As String
                        fp = FreeFile
                        Open path & file For Input Access Read As #fp
                            Line Input #fp, tmp
                            If (UCase$(tmp) = SKIN_MAGIC) Then
                                skins.Add path & file
                            End If
                        Close #fp
                    End If
                ElseIf (file = "") Then
                    ' no more files
                    Exit Do
                End If
                ' next file
                file = Dir()
            Loop
        End Sub
'--------------------------------------------------------------------------------
' END TEMPORAL CODE
' DELETE DELETE DELETE DELETE DELETE DELETE DELETE DELETE DELETE DELETE DELETE DE
'--------------------------------------------------------------------------------



'--------------------------------------------------------------------------------
' FORM LOAD/UNLOAD PROCEDURE
'--------------------------------------------------------------------------------
Private Sub Form_Load()
   ' Put Window control on Form
   Window.SetStyle
   Window.LoadSkin App.path & "\skins\vista\skin"
   Window.BackStyle = TRANSPARENT
   Me.BackColor = Window.BackColor
   Window.Top = 0
   Window.Left = 0
   Window.Height = Me.Height
   Window.Width = Me.Width
   Window.Caption = Me.Caption
   Window.ZOrder 0
   ' Window control default
   Window.CloseButton = True
   Window.HelpButton = True
   Window.MinimizeButton = True
   Window.MostTop = False
   Window.OnTopButton = True
   Window.RestoreButton = True
   Window.Sizable = True
   Window.Caption = App.Title
   ' Construct window region
   Window.buildWindowRgn
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '//Call killMe
End Sub

'--------------------------------------------------------------------------------
' WINDOW BUTTON CLICK PROCEDURE
'--------------------------------------------------------------------------------
Public Sub Window_CloseClick()
   Unload Me
End Sub

Private Sub Window_RestoreClick()
   If (Me.WindowState = vbNormal) Then
      Me.WindowState = vbMaximized
   ElseIf (Me.WindowState = vbMaximized) Then
      Me.WindowState = vbNormal
   End If
End Sub

Public Sub Window_MinimizeClick()
   If (Me.WindowState <> vbMinimized) Then
      Me.WindowState = vbMinimized
   End If
End Sub

Private Sub Window_HelpClick()
    MsgBox "Author: Abraham Araon H. L. / abrahamaraon@hotmail.com"
End Sub

Private Sub Window_OnTopClick()
   If (Window.MostTop) Then
      Window.MostTop = False
   Else
      Window.MostTop = True
   End If
End Sub

'--------------------------------------------------------------------------------
' WINDOW SYSTEM MENU PROCEDURE
'--------------------------------------------------------------------------------
Private Sub Window_BarTitleMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button = vbRightButton) Then
      If (Me.WindowState = vbNormal) Then
         Window.popUpSystemMenu
      End If
   End If
End Sub

'--------------------------------------------------------------------------------
' WINDOW MOVE PROCEDURE
'--------------------------------------------------------------------------------
Private Sub Window_BarTitleMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button = vbLeftButton) Then
      If (Me.WindowState <> vbMaximized) Then
         Window.startMove
      End If
   End If
End Sub

'--------------------------------------------------------------------------------
' WINDOW EXPAND/COLAPSE PROCEDURE
'--------------------------------------------------------------------------------
Private Sub Window_BarTitleDblClick()
   If (Me.WindowState = vbNormal) Then
      If (Window.State = COLLAPSED) Then
         Window.State = EXPANDED
         Me.Height = Window.Height
      ElseIf (Window.State = EXPANDED) Then
         Window.State = COLLAPSED
         Me.Height = Window.Height
      End If
   End If
End Sub

'--------------------------------------------------------------------------------
' WINDOW SIZE FORM PROCEDURE
'--------------------------------------------------------------------------------
Private Sub Form_Resize()
   If (Me.WindowState <> vbMinimized) Then
      If (Me.Width < Window.MinimeWidth) Then
         'SendKeys "{ENTER}"
         Me.Width = Window.MinimeWidth
      ElseIf (Me.Height < Window.MinimeHeight) Then
         'SendKeys "{ENTER}"
         Me.Height = Window.MinimeHeight
      Else
         Window.Width = Me.Width
         Window.Height = Me.Height
         Window.buildWindowRgn
      End If
   End If
End Sub

Private Sub Window_WindowMarkBottomMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button = vbLeftButton) Then
      If (Window.Sizable) Then
         If (Me.WindowState <> vbMaximized) Then
            Window.startSize TS_BOTTOM
         End If
      End If
   End If
End Sub

Private Sub Window_WindowMarkLeftBottomMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button = vbLeftButton) Then
      If (Window.Sizable) Then
         If (Me.WindowState <> vbMaximized) Then
            Window.startSize TS_LEFT_BOTTOM
         End If
      End If
   End If
End Sub

Private Sub Window_WindowMarkLeftMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button = vbLeftButton) Then
      If (Window.Sizable) Then
         If (Me.WindowState <> vbMaximized) Then
            Window.startSize TS_LEFT
         End If
      End If
   End If
End Sub

Private Sub Window_WindowMarkRightBottomMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button = vbLeftButton) Then
      If (Window.Sizable) Then
         If (Me.WindowState <> vbMaximized) Then
            Window.startSize TS_RIGHT_BOTTOM
         End If
      End If
   End If
End Sub

Private Sub Window_WindowMarkRightMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button = vbLeftButton) Then
      If (Window.Sizable) Then
         If (Me.WindowState <> vbMaximized) Then
            Window.startSize TS_RIGHT
         End If
      End If
   End If
End Sub

