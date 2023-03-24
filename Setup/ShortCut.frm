VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmShortCut 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregando entradas de registro"
   ClientHeight    =   690
   ClientLeft      =   3255
   ClientTop       =   1920
   ClientWidth     =   4620
   ForeColor       =   &H8000000F&
   Icon            =   "ShortCut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ckbRunProgram 
      Caption         =   "(e) Run program on starting Window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   390
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4530
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.CommandButton cmdReboot 
      Caption         =   "Reboot"
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
      Left            =   2490
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Close Windows and restart"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CheckBox ckbExplorerMenu 
      Caption         =   "(d) Explorer menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   375
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4185
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CheckBox ckbProgramsMenu 
      Caption         =   "(c) Programs menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3600
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3870
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.CheckBox ckbStartMenu 
      Caption         =   "(b) Start menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1905
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3870
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.CheckBox ckbDesktop 
      Caption         =   "(a) Desktop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   375
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3870
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      Left            =   3525
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
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
      Left            =   405
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Include program name"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      Left            =   1470
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Exclude program name"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Height          =   2100
      Left            =   345
      TabIndex        =   0
      Top             =   1695
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox txtArgList 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         TabIndex        =   14
         Top             =   1620
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtTitleRef 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   7
         Top             =   465
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.CommandButton cmdDialogFileSpec 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5580
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "File dialog"
         Top             =   1230
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtExecutableFileSpec 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   330
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   5595
      End
      Begin VB.Label lblArgList 
         Caption         =   "Arguments, if any:"
         Height          =   225
         Left            =   1800
         TabIndex        =   15
         Top             =   1650
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblTitleRef 
         Caption         =   "Title Ref/Shortcut name:"
         Height          =   225
         Left            =   210
         TabIndex        =   8
         Top             =   210
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Label lblExecutableFileSpec 
         Caption         =   "Path and name of executable file:"
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   930
         Visible         =   0   'False
         Width           =   2805
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1785
      Top             =   5865
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Por favor espere..."
      Height          =   210
      Left            =   120
      TabIndex        =   20
      Top             =   375
      Width           =   4440
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Agregando las entradas de registro y finalizando la instalación."
      Height          =   210
      Left            =   135
      TabIndex        =   19
      Top             =   120
      Width           =   4440
   End
   Begin VB.Label LblPath 
      Caption         =   "Label2"
      Height          =   195
      Left            =   2520
      TabIndex        =   18
      Top             =   7485
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   $"ShortCut.frx":000C
      Height          =   405
      Left            =   510
      TabIndex        =   13
      Top             =   6840
      Visible         =   0   'False
      Width           =   5625
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      X1              =   270
      X2              =   6330
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Visible         =   0   'False
      X1              =   270
      X2              =   6330
      Y1              =   7290
      Y2              =   7290
   End
End
Attribute VB_Name = "frmShortCut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VB6ShortCut.frm
'
' By Herman Liu
'
' (1) To create shortcuts linking a program from "Desktop", "Start" menu, and/or "Programs"
' menu, or to remove them  - This code is specially for VB6 users using VB6SKKIT.DLL for
' VB6 instead of STKIT432.DLL for older versions.
'
' (2) To add or delete a program name in "Explorer" menu (that menu invoked with a right
' click on the Start button). Instead of dropping a shortcut file in a directory, we set
' entries in the registry.
'
' (3) To enable/cancel running a program on starting the Windows.

Option Explicit

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SecurityAttributes
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type


Private Declare Function GetWindowsDirectory Lib "KERNEL32" Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetShortPathName Lib "KERNEL32" Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, _
    ByVal dwReserved As Long) As Long
Private Declare Function OSfCreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" _
    (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, _
    ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String, _
    ByVal fPrivate As Long, ByVal sParent As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal mKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal mKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
    ' --------------------------------------------------------------------------------
    ' Re RegSetValueEx: If you declare the lpData parameter as String, you must
    ' pass it By inValue.
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal mKey As Long, ByVal lpValueName As String, ByVal reserved As Long, _
    ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    ' --------------------------------------------------------------------------------
Private Declare Function RegSetValueExByte Lib "Advapi32" Alias "RegSetValueExA" _
    (ByVal mKey As Long, ByVal szValuename As String, ByVal lpReserved As Long, _
    ByVal dwValuetype As Long, bData As Byte, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "Advapi32" Alias "RegSetValueExA" _
    (ByVal mKey As Long, ByVal szValuename As String, ByVal lpReserved As Long, _
    ByVal dwValuetype As Long, dwData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExString Lib "Advapi32" Alias "RegSetValueExA" _
    (ByVal mKey As Long, ByVal szValuename As String, ByVal lpReserved As Long, _
    ByVal dwValuetype As Long, ByVal szData As String, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal mKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
    (ByVal mKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal mKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "Advapi32" Alias "RegCreateKeyExA" _
    (ByVal mKey As Long, ByVal szSubkey As String, ByVal lpReserved As Long, _
    ByVal szClass As String, ByVal dwOptions As Long, ByVal dwDesiredAccess As Long, _
    lpSecurityAttributes As SecurityAttributes, lphResult As Long, _
    lpdwDisposition As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
    (ByVal mKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
    lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, _
    lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal mKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
    lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, _
    lpcbData As Long) As Long
    
Private Const OPTION_NON_VOLATILE = &H0    ' Info is stored in a file and is preserved

Private Const FILE_ATTRIBUTE_NORMAL = &H80
 
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006

' Reg key security attribute
Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_SET_VALUE = &H2&
Private Const KEY_ALL_ACCESS = &H3F
Private Const KEY_CREATE_SUBKEY = &H4&
Private Const KEY_ENUMERATE_SUBKEY = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const KEY_CREATE_LINK = &H20
Private Const READ_CONTROL = &H20000
Private Const WRITE_OWNER = &H80000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or _
      KEY_ENUMERATE_SUBKEY Or KEY_NOTIFY
Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUBKEY

Private Const REG_NONE = 0&
Private Const REG_SZ = 1&                ' Unicode null terminated string
Private Const REG_BINARY = 3             ' Binary
Private Const REG_DWORD = 4              ' 32-bit number
Private Const REG_DWORD_BIG_ENDIAN = 5

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_BADDB = 1009&
Private Const ERROR_BADKEY = 1010&
Private Const ERROR_CANTOPEN = 1011&
Private Const ERROR_CANTREAD = 1012&
Private Const ERROR_CANTWRITE = 1013&
Private Const ERROR_OUTOFMEMORY = 14&
Private Const ERROR_INVALID_PARAMETER = 87&
Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const ERROR_MORE_DATA = 234&

Private Const mExplorerMenuKeyName = "DIRECTORY\SHELL"
Dim mWindowsDir As String
Dim mDeskTopPath As String
Dim mStartMenuPath As String
Dim mProgramsPath As String
Dim mDeskTopPathAbsolute As String
Dim mStartMenuPathAbsolute As String
Dim mProgramsPathAbsolute As String
Dim mTitleRef As String
Dim mKeyHandle As Long
Dim mresult



Private Function GetWinDir() As String
    On Error Resume Next
    Dim mBuffer As String * 256
    Dim mDir As String
    mresult = GetWindowsDirectory(mBuffer, 256)
    If Len(mresult) = 0 Then
         MsgBox "Failed to get windows dir"
         Exit Function
    End If
    mDir = Mid$(mBuffer, 1, mresult)
    GetWinDir = mDir
End Function



Private Sub cmdDialogFileSpec_Click()
     On Error GoTo errHandler
     Dim gcdg As Object
     Set gcdg = CommonDialog1
     gcdg.Filter = "(*.exe)|*.exe|(*.com)|*.com|(*.*)|*.*|"
     gcdg.FilterIndex = 1
     gcdg.DefaultExt = "exe"
     gcdg.Flags = cdlOFNFileMustExist
     gcdg.FileName = ""
     gcdg.CancelError = True
     gcdg.ShowOpen
     If gcdg.FileName = "" Then
         Exit Sub
     End If
     txtExecutableFileSpec.Text = gcdg.FileName
     Exit Sub
errHandler:
     If Err <> 32755 Then
         ErrMsgProc "cmdDialogFileSpec_Click"
     End If
End Sub



Sub cmdAdd_Click()
    
    If Not AtLeastOneCheckBox Then
         MsgBox "No check box ticked yet"
         Exit Sub
    End If
    If Trim(txtTitleRef.Text) = "" Then
         MsgBox "No title ref entered yet"
         Exit Sub
    ElseIf Trim(txtExecutableFileSpec.Text) = "" Then
         MsgBox "No executable file spec entered yet"
         Exit Sub
    End If
    If IsFileThere(Trim(txtExecutableFileSpec.Text)) = False Then
         MsgBox Trim(txtExecutableFileSpec.Text) & " not found"
         Exit Sub
    End If
    If Not doAddShortCut Then
         Exit Sub
    End If
End Sub



Private Function AtLeastOneCheckBox() As Boolean
    AtLeastOneCheckBox = False
    If ckbDesktop.value = 1 Then
        AtLeastOneCheckBox = True
        Exit Function
    ElseIf ckbStartMenu.value = 1 Then
        AtLeastOneCheckBox = True
        Exit Function
    ElseIf ckbProgramsMenu.value = 1 Then
        AtLeastOneCheckBox = True
        Exit Function
    ElseIf ckbExplorerMenu.value = 1 Then
        AtLeastOneCheckBox = True
        Exit Function
    ElseIf ckbRunProgram.value = 1 Then
        AtLeastOneCheckBox = True
        Exit Function
    End If
End Function



Private Function doAddShortCut() As Boolean
    
    On Error GoTo errHandler
    doAddShortCut = False
    Dim mExeFileSpec As String
    Dim mDestPath As String
    Dim mArgList As String
    Dim mTitleRef As String
    Dim mPrivate As Boolean
    Dim mParent As String

    mExeFileSpec = Trim(txtExecutableFileSpec.Text)
    mArgList = Trim(txtArgList.Text)
    mTitleRef = Trim(txtTitleRef.Text)
    mPrivate = True
    mParent = "$(Programs)"
    If ckbDesktop.value = 1 Then
        mDestPath = mDeskTopPath
        OSfCreateShellLink mDestPath, mTitleRef, mExeFileSpec, mArgList, mPrivate, mParent
    End If
    If ckbStartMenu.value = 1 Then
        mDestPath = mStartMenuPath
        OSfCreateShellLink mDestPath, mTitleRef, mExeFileSpec, mArgList, mPrivate, mParent
    End If
    If ckbProgramsMenu.value = 1 Then
        mDestPath = mProgramsPath
        OSfCreateShellLink mDestPath, mTitleRef, mExeFileSpec, mArgList, mPrivate, mParent
    End If
    doAddShortCut = True
    Exit Function
errHandler:
End Function



Private Function doAddRegistry() As Boolean
    On Error GoTo errHandler
    doAddRegistry = False
    Dim mSubKeySpec As String
    Dim mSubSub As String
    Dim mKey As Long
    Dim mExeFileSpec As String
    Dim DispBuffer As Long
    Dim typSA As SecurityAttributes
    
    typSA.lpSecurityDescriptor = KEY_ALL_ACCESS
    mKeyHandle = HKEY_CLASSES_ROOT
    mTitleRef = Trim(txtTitleRef.Text)
    mSubKeySpec = mExplorerMenuKeyName & "\" & mTitleRef
    
    mExeFileSpec = Trim(txtExecutableFileSpec.Text)
    mresult = RegOpenKeyEx(mKeyHandle, mSubKeySpec, 0, KEY_READ, mKey)
    If mresult <> 0 Then           ' Not exist yet
        If RegCreateKeyEx(mKeyHandle, mSubKeySpec, 0, "", _
             OPTION_NON_VOLATILE, KEY_ALL_ACCESS, typSA, mKey, DispBuffer) <> 0 Then
             MsgBox "Unable to create " & mSubKeySpec
             Exit Function
        End If
    End If
       ' Enter the value as "default" value
    If SetRegEntry(mKeyHandle, mSubKeySpec, "", mTitleRef) = 0 Then
         MsgBox "Unable to enter " & mTitleRef & " to " & mSubKeySpec
         Exit Function
    End If
    RegCloseKey mKey
       ' Now create a "Command" subkey under mSubKeySpec
    mSubSub = mSubKeySpec & "\Command"
    If RegCreateKeyEx(mKeyHandle, mSubSub, 0, "", _
         OPTION_NON_VOLATILE, KEY_ALL_ACCESS, typSA, mKey, DispBuffer) <> 0 Then
         MsgBox "Unable to create " & mSubSub
         Exit Function
    End If
       ' Enter the value as "default" value
    If SetRegEntry(mKeyHandle, mSubSub, "", mExeFileSpec) = 0 Then
         MsgBox "Unable to enter " & mTitleRef & " to " & mSubSub
         Exit Function
    End If
    RegCloseKey mKey
       ' Now create a "DefaultIcon" subkey under mSubKeySpec
    mSubSub = mSubKeySpec & "\DefaultIcon"
    If RegCreateKeyEx(mKeyHandle, mSubSub, 0, "", _
        OPTION_NON_VOLATILE, KEY_ALL_ACCESS, typSA, mKey, DispBuffer) <> 0 Then
         MsgBox "Unable create " & mSubSub
         Exit Function
    End If
    RegCloseKey mKey
    If SetRegEntry(mKeyHandle, mSubSub, "", mExeFileSpec & ",1") = 0 Then
         MsgBox "Unable to set value to " & mSubSub & " (as default icon)"
         Exit Function
    End If
    RegCloseKey mKey
    doAddRegistry = True
    Exit Function
errHandler:
End Function



Private Function GetRegEntry(ByVal inMainKey As Long, ByVal inSubKey As String, ByVal inEntry As String) As String
    Dim mKey As Long
    Dim mBuffer As String * 255
    Dim mBufSize As Long
    mresult = RegOpenKeyEx(inMainKey, inSubKey, 0, KEY_READ, mKey)
    If mresult = 0 Then
          mBufSize = Len(mBuffer)
          mresult = RegQueryValueEx(mKey, inEntry, 0, REG_SZ, mBuffer, mBufSize)
          If mresult = 0 Then
                If mBuffer <> "" Then
                     GetRegEntry = Mid$(mBuffer, 1, mBufSize)
                End If
                RegCloseKey mKey
          Else         ' inValue may be simply not exist, not an error
                GetRegEntry = ""
          End If
    Else
          MsgBox "Unable to open " & inSubKey
          GetRegEntry = ""
    End If
End Function



Private Function SetRegEntry(ByVal inMainKey As Long, ByVal inSubKey As String, ByVal inEntry As String, ByVal inValue As String) As Boolean
    Dim mKey As Long
    mresult = RegOpenKeyEx(inMainKey, inSubKey, 0, KEY_WRITE, mKey)
    If mresult <> 0 Then
         SetRegEntry = False
         Exit Function
    End If
        ' Here we set value as REG_SZ type, you may set it to other type, e.g.
        ' if the type is REG_DWORD: mresult = RegSetValueExLong(mKey, inEntry,
        ' 0, REG_DWORD, inValue, 4)
    mresult = RegSetValueExString(mKey, inEntry, 0, REG_SZ, inValue, Len(inValue))
    If mresult <> 0 Then
         MsgBox "Unable to set value of " & inValue & " to subkey " & inEntry
    End If
    RegCloseKey mKey
    SetRegEntry = (mresult = 0)
End Function



Private Sub cmdDelete_Click()
    If Not AtLeastOneCheckBox Then
         MsgBox "No check box ticked yet"
         Exit Sub
    End If
    If ckbDesktop.value = 1 Or ckbStartMenu.value = 1 Or _
        ckbProgramsMenu.value = 1 Or ckbExplorerMenu.value = 1 Then
        If Trim(txtTitleRef.Text) = "" Then
             MsgBox "No title ref entered yet"
             Exit Sub
        End If
        txtExecutableFileSpec.Text = ""
    ElseIf ckbRunProgram.value = 1 Then
        txtTitleRef.Text = ""
        txtExecutableFileSpec.Text = ""
    End If
    
    If MsgBox("Proceed?", vbYesNo + vbQuestion) <> vbYes Then
         Exit Sub
    End If
    doDelShortCut
    If ckbExplorerMenu.value = 1 Then
         If Not doDelRegistry Then
             Exit Sub
         End If
    End If
    If ckbRunProgram.value = 1 Then
         If Not doRemoveRunProgram Then
             Exit Sub
         End If
    End If
    MsgBox "Deletion(s) done"
End Sub




Private Function doDelShortCut()
    On Error Resume Next
    Dim mDestPath As String
    Dim mTitleRef As String
    mTitleRef = Trim(txtTitleRef.Text) & ".lnk"
    If ckbDesktop.value = 1 Then
        mDestPath = mDeskTopPathAbsolute & "\" & mTitleRef
        If IsFileThere(mDestPath) Then
            Kill mDestPath
        End If
    End If
    If ckbStartMenu.value = 1 Then
        mDestPath = mStartMenuPathAbsolute & "\" & mTitleRef
        If IsFileThere(mDestPath) Then
            Kill mDestPath
        End If
    End If
    If ckbProgramsMenu.value = 1 Then
        mDestPath = mProgramsPathAbsolute & "\" & mTitleRef
        If IsFileThere(mDestPath) Then
            Kill mDestPath
        End If
    End If
End Function



Private Function LongToShort(inSpec) As String
    Dim i
    Dim ShortSpec As String
    Dim mBuffer As String
    Dim mBufLen As Long
    mBufLen = 164
    mBuffer = String(mBufLen, 0)
    i = GetShortPathName(inSpec, mBuffer, mBufLen)
    LongToShort = Left$(mBuffer, i)
End Function



Private Function doDelRegistry() As Boolean
    On Error GoTo errHandler
    doDelRegistry = False
    Dim mKey As Long
    Dim mSubKeySpec As String
    Dim mTitleRef As String
    mTitleRef = Trim(txtTitleRef.Text)
    mSubKeySpec = mExplorerMenuKeyName & "\" & mTitleRef
    mKeyHandle = HKEY_CLASSES_ROOT
    mresult = RegOpenKeyEx(mKeyHandle, mSubKeySpec, 0, KEY_READ, mKey)
    If mresult <> 0 Then
         Exit Function
    End If
    RegCloseKey mKey
    If Not DoDeleteRegKey(mKeyHandle, mSubKeySpec) Then
         MsgBox "Failed to complete deletion of " & mSubKeySpec
         Exit Function
    End If
    doDelRegistry = True
    Exit Function
errHandler:
End Function



Private Function DoDeleteRegKey(ByVal inMainKey As Long, ByVal inSubKey As String) As Boolean
    On Error GoTo errHandler
    Dim One_Level_Up As String
    Dim mKey As Long
    Dim mPos As Integer

    If Right$(inSubKey, 1) = "\" Then
         inSubKey = Left$(inSubKey, Len(inSubKey) - 1)
    End If
       ' Delete the inSubkey's own subkeys first
    If DeleteSubkeys(inMainKey, inSubKey) = False Then
         GoTo errHandler
    End If
    
       ' Get the parent of inSubkey
    mPos = InStrRev(inSubKey, "\")
    If mPos = 0 Then
           ' This is a top-level key, delete it from the inMainKey.
         RegDeleteKey inMainKey, inSubKey
    Else
           ' Find the parent key within inSubKey itself.
         One_Level_Up = Left$(inSubKey, mPos - 1)
         inSubKey = Mid$(inSubKey, mPos + 1)
         mresult = RegOpenKeyEx(inMainKey, One_Level_Up, 0, KEY_ALL_ACCESS, mKey)
         If mresult = 0 Then
              RegDeleteKey mKey, inSubKey
              RegCloseKey mKey
         End If
    End If
    DoDeleteRegKey = True
    Exit Function
errHandler:
    DoDeleteRegKey = False
End Function



  ' Delete all the subkey's subkeys.
Private Function DeleteSubkeys(ByVal inMainKey As Long, ByVal inSubKey As String) As Boolean
    On Error GoTo errHandler
    Dim mKey As Long
    Dim colSubKeys As Collection
    Dim mClassBuffer As String * 255
    Dim mClassBufSize As Long
    Dim typLastWriteTime As FILETIME
    Dim mIndex As Long
    Dim mBufSize As Long
    Dim mSubKeyName As String

    mresult = RegOpenKeyEx(inMainKey, inSubKey, 0, KEY_ALL_ACCESS, mKey)
    If mresult <> 0 Then
         MsgBox "Unable to open " & inSubKey
         GoTo errHandler
    End If

    Set colSubKeys = New Collection
    mIndex = 0
    Do
         ' lpClassBuffer is a pointer to a buffer that receives the null-terminated
         ' class string of the enumerated subkey. No classes are currently defined;
         ' hence this parameter can be NULL.
         ' lpClassBufSize is a pointer to a variable that specifies the size of
         ' lpClassBuffer, including the terminating null character. When the function
         ' returns, it contains the number of characters stored in the buffer.
         ' The count returned does not include the terminating null character.
        mClassBuffer = ""
        mClassBufSize = 0
        mBufSize = 256
        mSubKeyName = Space$(mBufSize)
        mresult = RegEnumKeyEx(mKey, mIndex, mSubKeyName, mBufSize, 0, mClassBuffer, _
                mClassBufSize, typLastWriteTime)
    
        If mresult <> 0 Then                    ' No more
             DeleteSubkeys = True
             Exit Function
        End If
        mIndex = mIndex + 1
        mSubKeyName = Left$(mSubKeyName, InStr(mSubKeyName, Chr$(0)) - 1)
        colSubKeys.Add mSubKeyName
    Loop
    For mIndex = 1 To colSubKeys.Count
          ' Delete the subkey's colSubKeys first
        DeleteSubkeys inMainKey, inSubKey & "\" & colSubKeys(mIndex)
          ' Effect delete of the subkey itself
        RegDeleteKey mKey, colSubKeys(mIndex)
    Next mIndex

    RegCloseKey mKey
    
    DeleteSubkeys = True
    Exit Function
    
errHandler:
    DeleteSubkeys = False
End Function



Private Function doAddRunProgram() As Boolean
    On Error GoTo errHandler
    doAddRunProgram = False
    Dim mSubKeySpec As String
    Dim mSubSub As String
    Dim mExeFileSpec As String
    Dim mKey As Long
    Dim DispBuffer As Long
    Dim typSA As SecurityAttributes
    typSA.lpSecurityDescriptor = KEY_ALL_ACCESS
    mKeyHandle = HKEY_CURRENT_USER
    mSubSub = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    mExeFileSpec = Trim(txtExecutableFileSpec.Text)
    If RegCreateKeyEx(mKeyHandle, mSubSub, 0, "", _
         OPTION_NON_VOLATILE, KEY_ALL_ACCESS, typSA, mKey, DispBuffer) <> 0 Then
         MsgBox "Unable to create " & mSubSub
         Exit Function
    End If
       ' Enter the value as "default" value
    If SetRegEntry(mKeyHandle, mSubSub, "", mExeFileSpec) = 0 Then
         MsgBox "Unable to enter " & mExeFileSpec & " to " & mSubSub
         Exit Function
    End If
    RegCloseKey mKey
    doAddRunProgram = True
    Exit Function
errHandler:
End Function



Private Function doRemoveRunProgram() As Boolean
    On Error GoTo errHandler
    doRemoveRunProgram = False
    Dim mSubKeySpec As String
    Dim mSubSub As String
    Dim mKey As Long
    mKeyHandle = HKEY_CURRENT_USER
    mSubSub = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    mresult = RegOpenKeyEx(mKeyHandle, mSubSub, 0, KEY_WRITE, mKey)
    If mresult = 0 Then
           ' Simply enter "" value to (Default)
         mresult = RegSetValueExString(mKey, "", 0, REG_SZ, "", 0)
         If mresult = 0 Then
             RegCloseKey mKey
         Else
             MsgBox "Unable to write a value to Default entry of " & mSubSub
             Exit Function
         End If
    Else
         MsgBox "Unable to delete entry of " & mSubSub
         Exit Function
    End If
    doRemoveRunProgram = True
    Exit Function
errHandler:
End Function



Private Sub cmdReboot_Click()
    If MsgBox("Proceed to reboot?", vbYesNo + vbQuestion) <> vbYes Then
         Exit Sub
    End If
    Const EWX_SHUTDOWN = 1
    Const EWX_REBOOT = 2
    Const EWX_LOGOFF = 0
    Const EWX_FORCE = 4
    ExitWindowsEx EWX_REBOOT, 0
End Sub



Private Sub cmdExit_Click()
    End
End Sub


Private Sub Form_Load()

Dim APData As StpIniFile
Dim mEXEname As String, mEXEDesc As String
Dim mREADname As String, mREADDesc As String
Dim MSubProgDir As String
Dim ExPath As String

DoEvents

If IsEnglishWin = True Then
    mWindowsDir = GetWinDir
    mDeskTopPath = "..\..\Desktop"
    mStartMenuPath = ".."
    mProgramsPath = mStartMenuPath & "\Programs"
    mDeskTopPathAbsolute = mWindowsDir & "\Desktop"
    mStartMenuPathAbsolute = mWindowsDir & "\Start Menu"
    mProgramsPathAbsolute = mStartMenuPathAbsolute & "\Programs"
Else
    mWindowsDir = GetWinDir
    mDeskTopPath = "..\..\Escritorio"
    mStartMenuPath = ".."
    mProgramsPath = mStartMenuPath & "\Programas"
    mDeskTopPathAbsolute = mWindowsDir & "\Escritorio"
    mStartMenuPathAbsolute = mWindowsDir & "\Menu Inicio"
    mProgramsPathAbsolute = mStartMenuPathAbsolute & "\Programas"
End If


'//////////////////////////////////////////////////////////
'cargamos los datos del archivo ini
APData = ReadIniFile(App.Path, 1)
'----
mEXEname = Trim(Desencriptar(DefPass, APData.APPEXEName))
mEXEDesc = Trim(Desencriptar(DefPass, APData.APPEXEDesc))
mREADname = Trim(Desencriptar(DefPass, APData.APPReadmeName))
mREADDesc = Trim(Desencriptar(DefPass, APData.APPReadmeDesc))
MSubProgDir = Trim(Desencriptar(DefPass, APData.AppCompany))
'//////////////////////////////////////////////////////////

'creamos un directorio
On Error Resume Next
MkDir mProgramsPathAbsolute & "\" & MSubProgDir
mProgramsPath = mProgramsPath & "\" & MSubProgDir

LblPath.Caption = FrmCopy.Lb1.Caption
ExPath = Trim(LblPath.Caption)
Unload FrmCopy

'add the first entry
txtTitleRef.Text = mEXEDesc
txtExecutableFileSpec.Text = ExPath & "\" & mEXEname
ckbProgramsMenu.value = 1
ckbDesktop.value = 1
Call cmdAdd_Click

'add the second entry
txtTitleRef.Text = mREADDesc
txtExecutableFileSpec.Text = ExPath & "\" & mREADname
ckbProgramsMenu.value = 1
ckbDesktop.value = 0
Call cmdAdd_Click

MsgBox "Instalación Completada con éxito!."
End

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub



Function IsFileThere(inFileSpec As String) As Boolean
    On Error Resume Next
    Dim mFile
    Dim i
    mFile = LongToShort(inFileSpec)
    i = FreeFile
    Open mFile For Input As i
    If Err Then
        IsFileThere = False
    Else
        Close i
        IsFileThere = True
    End If
End Function



Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub



