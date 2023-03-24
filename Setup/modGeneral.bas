Attribute VB_Name = "ModBrowse"
'Copyright (c) 1987-2002 Creaciones Digitales Inc.

Option Explicit

Public Type SHITEMID    'Browse Dialog
   cb             As Long
   abID           As Byte
End Type

Public Type ITEMIDLIST  'Browse Dialog
   mkid           As SHITEMID
End Type

Public Type BROWSEINFO  'Browse Dialog
   hOwner         As Long
   pidlRoot       As Long
   pszDisplayName As String
   lpszTitle      As String
   ulFlags        As Long
   lpfn           As Long
   lParam         As Long
   iImage         As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = &H1 'Browse Dialog
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Public Cancelled As Boolean

'Opens Browse for folder dialog box
Public Function BrowseForFolder(Optional Title As String) As String
   
   Dim bi As BROWSEINFO
   Dim pidl As Long
   Dim nRet As Long
   Dim szPath As String
   
   szPath = Space$(512)
   
   bi.hOwner = 0&
   bi.pidlRoot = 0&
   
   bi.lpszTitle = IIf(Title = "", "Directory", Title)
   bi.ulFlags = BIF_RETURNONLYFSDIRS
   
   'Display the dialog and get the selected path
   pidl& = SHBrowseForFolder(bi)
   SHGetPathFromIDList ByVal pidl&, ByVal szPath
   
   'variable de retorno
   BrowseForFolder = Trim$(szPath)
   
End Function

