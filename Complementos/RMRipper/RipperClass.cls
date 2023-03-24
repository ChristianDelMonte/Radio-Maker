VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RipperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get PlugInName() As String
    PlugInName = "RM CD Ripper..."
End Property

Public Function GetPlugInState() As Long

If frmMain.WindowState = 0 Then
    GetPlugInState = 1      'estado abierto
Else
    GetPlugInState = 0      'estado cerrado
End If
    
End Function

Public Property Get PlugInInfo() As String
    PlugInInfo = "RMRipper_dll for Only Radiomaker"
End Property

Public Property Get VersionInfo() As String
    VersionInfo = App.Major & "." & App.Minor & "." & App.Revision
End Property

Sub LoadControl(WAuthorCode As String)

Dim Txtin As String
Dim Txtout As String
Dim Passw As String

Txtin = WAuthorCode
Passw = "Radiomaker_dll"

Txtout = DecipherData(Passw, Txtin)

If Trim(Txtout) = "Radiomaker_dll_InitComponent" Then
    frmMain.Show
Else
    MsgBox "Codigo de Autorización Inválido.", vbCritical, "RmRipper.dll"
End If

End Sub

Public Sub UnloadControl()

Unload frmMain

End Sub

