VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddonClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function PlugIsRunning() As Boolean

If Pstate = False Then
    PlugIsRunning = False
Else
    PlugIsRunning = True
End If

End Function

Public Function GetPlugInState() As Long

If MainForm.WindowState = 0 Then
    GetPlugInState = 1      'estado abierto
Else
    GetPlugInState = 0      'estado cerrado
End If
    
End Function

Public Property Get WTemperature() As String
Dim Temp As String

Temp = GetTR
If Temp = "err" Then
    WTemperature = "N/A"
Else
    WTemperature = Temp
End If

End Property

Public Property Get WHumedad() As String
Dim Hume As String

Hume = GetHU
If Hume = "err" Then
    WHumedad = "N/A"
Else
    WHumedad = Hume
End If

End Property

Public Property Get PlugInInfo() As String
    PlugInInfo = "RadioMaker Voice Player DLL"
End Property

Public Property Get PlugInName() As String
    PlugInName = "RMVoice"
End Property

Public Property Get DLLName() As String
    DLLName = "RMVoice"
End Property

Public Property Get ClassName() As String
    ClassName = "AddonClass"
End Property

Public Property Get VersionInfo() As String
    VersionInfo = App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Function LoadSilent() As Integer

LoadSilent = 1

End Function

Public Function LoadControl(WAuthorCode As String, WAction As String) As Integer

Dim Txtin As String
Dim Txtout As String
Dim Passw As String

Txtin = WAuthorCode
Passw = "Radiomaker_dll"

'verificamos la autenticidad de radiomaker
Txtout = DecipherData(Passw, Txtin)
If Trim(Txtout) = "Radiomaker_dll_InitComponent" Then
    LoadControl = 1
    'xxx nothing... se continua con la carga del plugin
Else
    DisplayMsg GetComLng_ByID(LNGDef, "1011"), " in LoadControl. ", err.Number, True    '1011 = codigo de autorizacion invalido
    LoadControl = 0
    Exit Function
End If

Select Case WAction
    Case "ShowConfig"
        MainForm.Show
    Case "ShowMain"
        MainForm.Show
    Case "SayHora"
        Call InitHora
    Case "SayMinutos"
        Call InitMinutos
    Case "SayTemperatura"
        Call InitTemperatura
    Case "SayHumedad"
        Call InitHumedad
    Case "LoadSilent"
        'xxxx
    Case Else
        'xxx... nothing... solo cargar
End Select

End Function

Public Sub UnloadControl()

Unload MainForm

End Sub

