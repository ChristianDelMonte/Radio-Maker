VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddonClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetPlugInState() As Long

If Main.WindowState = 0 Then
    GetPlugInState = 1      'estado abierto
Else
    GetPlugInState = 0      'estado cerrado
End If
    
End Function

Public Property Get PlugInInfo() As String
    PlugInInfo = "RadioMaker Mini Player DLL"
End Property

Public Property Get PlugInName() As String
    PlugInName = "RMPlayer"
End Property

Public Property Get DLLName() As String
    DLLName = "RMPlayer"
End Property

Public Property Get ClassName() As String
    ClassName = "AddonClass"
End Property

Public Property Get VersionInfo() As String
    VersionInfo = App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Sub SetBuffLevel(ByVal WLeft As Long, ByVal WRight As Long)

'SUB para visualizar un gráfico de 25 cuadros
'al compas de la musica.

On Error Resume Next

Dim l, Lft As Long
Dim r, Rgt As Long

Lft = WLeft / 1000
Rgt = WRight / 1000

If Lft >= 24 Then
    Lft = 24
    Main.PicAnmL.Picture = Main.PicLL.GraphicCell(Lft)
Else
    Main.PicAnmL.Picture = Main.PicLL.GraphicCell(Lft)
End If

If Rgt >= 24 Then
    Rgt = 24
    Main.PicAnmR.Picture = Main.PicRR.GraphicCell(Rgt)
Else
    Main.PicAnmR.Picture = Main.PicRR.GraphicCell(Rgt)
End If

End Sub

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
    MsgBox "Codigo de Autorización Inválido. Consulte a su proveedor de software", vbCritical, "RmPlayer.dll"
    LoadControl = 0
    Exit Function
End If

Select Case WAction
    Case "ShowConfig"
        Main.Show
    Case "ShowMain"
        Main.Show
    Case Else
        'xxx nothing... solo cargar
End Select

End Function

Public Sub SetNextTime(WFormatedTime As String)

'SOLO PARA Mini-RM
'el wformatedtime debe estar formateado = 00:00:00

Dim N1, N2, N3, N4, N5, N6
Dim LenNum

LenNum = Len(WFormatedTime)

If LenNum <> 8 Then
    Exit Sub
End If

If LenNum = 8 Then
    WFormatedTime = LTrim(RTrim(WFormatedTime))
    N1 = Left$(WFormatedTime, 1)
    N2 = Mid$(WFormatedTime, 2, 1)
    N3 = Mid$(WFormatedTime, 4, 1)
    N4 = Mid$(WFormatedTime, 5, 1)
    N5 = Mid$(WFormatedTime, 7, 1)
    N6 = Right$(WFormatedTime, 1)
        
    If N1 = 0 Then Main.Tp1.Picture = Main.SmallClip.GraphicCell(0)
    If N1 = 1 Then Main.Tp1.Picture = Main.SmallClip.GraphicCell(1)
    If N1 = 2 Then Main.Tp1.Picture = Main.SmallClip.GraphicCell(2)
    If N1 = 3 Then Main.Tp1.Picture = Main.SmallClip.GraphicCell(3)
    If N1 = 4 Then Main.Tp1.Picture = Main.SmallClip.GraphicCell(4)
    If N1 = 5 Then Main.Tp1.Picture = Main.SmallClip.GraphicCell(5)
    If N1 = 6 Then Main.Tp1.Picture = Main.SmallClip.GraphicCell(6)
    If N1 = 7 Then Main.Tp1.Picture = Main.SmallClip.GraphicCell(7)
    If N1 = 8 Then Main.Tp1.Picture = Main.SmallClip.GraphicCell(8)
    If N1 = 9 Then Main.Tp1.Picture = Main.SmallClip.GraphicCell(9)
            
    If N2 = 0 Then Main.Tp2.Picture = Main.SmallClip.GraphicCell(0)
    If N2 = 1 Then Main.Tp2.Picture = Main.SmallClip.GraphicCell(1)
    If N2 = 2 Then Main.Tp2.Picture = Main.SmallClip.GraphicCell(2)
    If N2 = 3 Then Main.Tp2.Picture = Main.SmallClip.GraphicCell(3)
    If N2 = 4 Then Main.Tp2.Picture = Main.SmallClip.GraphicCell(4)
    If N2 = 5 Then Main.Tp2.Picture = Main.SmallClip.GraphicCell(5)
    If N2 = 6 Then Main.Tp2.Picture = Main.SmallClip.GraphicCell(6)
    If N2 = 7 Then Main.Tp2.Picture = Main.SmallClip.GraphicCell(7)
    If N2 = 8 Then Main.Tp2.Picture = Main.SmallClip.GraphicCell(8)
    If N2 = 9 Then Main.Tp2.Picture = Main.SmallClip.GraphicCell(9)

    Main.Tp3.Picture = Main.SmallClip.GraphicCell(11) '= :
        
    If N3 = 0 Then Main.Tp4.Picture = Main.SmallClip.GraphicCell(0)
    If N3 = 1 Then Main.Tp4.Picture = Main.SmallClip.GraphicCell(1)
    If N3 = 2 Then Main.Tp4.Picture = Main.SmallClip.GraphicCell(2)
    If N3 = 3 Then Main.Tp4.Picture = Main.SmallClip.GraphicCell(3)
    If N3 = 4 Then Main.Tp4.Picture = Main.SmallClip.GraphicCell(4)
    If N3 = 5 Then Main.Tp4.Picture = Main.SmallClip.GraphicCell(5)
    If N3 = 6 Then Main.Tp4.Picture = Main.SmallClip.GraphicCell(6)
    If N3 = 7 Then Main.Tp4.Picture = Main.SmallClip.GraphicCell(7)
    If N3 = 8 Then Main.Tp4.Picture = Main.SmallClip.GraphicCell(8)
    If N3 = 9 Then Main.Tp4.Picture = Main.SmallClip.GraphicCell(9)
    
    If N4 = 0 Then Main.Tp5.Picture = Main.SmallClip.GraphicCell(0)
    If N4 = 1 Then Main.Tp5.Picture = Main.SmallClip.GraphicCell(1)
    If N4 = 2 Then Main.Tp5.Picture = Main.SmallClip.GraphicCell(2)
    If N4 = 3 Then Main.Tp5.Picture = Main.SmallClip.GraphicCell(3)
    If N4 = 4 Then Main.Tp5.Picture = Main.SmallClip.GraphicCell(4)
    If N4 = 5 Then Main.Tp5.Picture = Main.SmallClip.GraphicCell(5)
    If N4 = 6 Then Main.Tp5.Picture = Main.SmallClip.GraphicCell(6)
    If N4 = 7 Then Main.Tp5.Picture = Main.SmallClip.GraphicCell(7)
    If N4 = 8 Then Main.Tp5.Picture = Main.SmallClip.GraphicCell(8)
    If N4 = 9 Then Main.Tp5.Picture = Main.SmallClip.GraphicCell(9)
        
    Main.Tp6.Picture = Main.SmallClip.GraphicCell(11) '= :
    
    If N5 = 0 Then Main.Tp7.Picture = Main.SmallClip.GraphicCell(0)
    If N5 = 1 Then Main.Tp7.Picture = Main.SmallClip.GraphicCell(1)
    If N5 = 2 Then Main.Tp7.Picture = Main.SmallClip.GraphicCell(2)
    If N5 = 3 Then Main.Tp7.Picture = Main.SmallClip.GraphicCell(3)
    If N5 = 4 Then Main.Tp7.Picture = Main.SmallClip.GraphicCell(4)
    If N5 = 5 Then Main.Tp7.Picture = Main.SmallClip.GraphicCell(5)
    If N5 = 6 Then Main.Tp7.Picture = Main.SmallClip.GraphicCell(6)
    If N5 = 7 Then Main.Tp7.Picture = Main.SmallClip.GraphicCell(7)
    If N5 = 8 Then Main.Tp7.Picture = Main.SmallClip.GraphicCell(8)
    If N5 = 9 Then Main.Tp7.Picture = Main.SmallClip.GraphicCell(9)
    
    If N6 = 0 Then Main.Tp8.Picture = Main.SmallClip.GraphicCell(0)
    If N6 = 1 Then Main.Tp8.Picture = Main.SmallClip.GraphicCell(1)
    If N6 = 2 Then Main.Tp8.Picture = Main.SmallClip.GraphicCell(2)
    If N6 = 3 Then Main.Tp8.Picture = Main.SmallClip.GraphicCell(3)
    If N6 = 4 Then Main.Tp8.Picture = Main.SmallClip.GraphicCell(4)
    If N6 = 5 Then Main.Tp8.Picture = Main.SmallClip.GraphicCell(5)
    If N6 = 6 Then Main.Tp8.Picture = Main.SmallClip.GraphicCell(6)
    If N6 = 7 Then Main.Tp8.Picture = Main.SmallClip.GraphicCell(7)
    If N6 = 8 Then Main.Tp8.Picture = Main.SmallClip.GraphicCell(8)
    If N6 = 9 Then Main.Tp8.Picture = Main.SmallClip.GraphicCell(9)
End If

End Sub

Public Sub SetTime(WTime As String, WType As String)

'-------- SetDigClock dims ------------------
Dim Minutos As String
Dim M1, M2 As Integer
Dim Segundos As String
Dim S1, S2 As Integer
Dim Resultado As String

'formatea el tiempo de salida de los temas
'para mostrarlos en el display
'WTime debe ser: 00:00:00 y el resultado es: 00:00 or -00:00
'WType debe ser: Normal or Restante

'formateamos el tiempo y separamos los minutos de los segundos
Resultado = Trim(Right$(WTime, 5))
Minutos = Left$(Resultado, 2)
Segundos = Right$(Resultado, 2)
M1 = CInt(Left$(Minutos, 1)): M2 = CInt(Right$(Minutos, 1))
S1 = CInt(Left$(Segundos, 1)): S2 = CInt(Right$(Segundos, 1))

'seteamos el display con los numeros correspondientes
Select Case WType
    Case "Normal"
        Main.tr1.Picture = Main.SmallClip.GraphicCell(10)  '= nada
        Main.tr2.Picture = Main.SmallClip.GraphicCell(M1)
        Main.tr3.Picture = Main.SmallClip.GraphicCell(M2)
        Main.tr4.Picture = Main.SmallClip.GraphicCell(11) '= :
        Main.tr5.Picture = Main.SmallClip.GraphicCell(S1)
        Main.tr6.Picture = Main.SmallClip.GraphicCell(S2)
    Case "Restante"
        Main.tr1.Picture = Main.SmallClip.GraphicCell(13)  '= signo menos
        Main.tr2.Picture = Main.SmallClip.GraphicCell(M1)
        Main.tr3.Picture = Main.SmallClip.GraphicCell(M2)
        Main.tr4.Picture = Main.SmallClip.GraphicCell(11) '= :
        Main.tr5.Picture = Main.SmallClip.GraphicCell(S1)
        Main.tr6.Picture = Main.SmallClip.GraphicCell(S2)
End Select

End Sub

Public Sub DrawMiniFFT(Wdata() As Single, sSize)

On Error Resume Next

Dim D() As Single
Dim TopLevel As Long
Static Peak(128) As Long
Dim a As Long
Dim sCount As Long
Dim f As String
Dim z As Long
Dim sLeft As Integer
    
ReDim D(256) As Single  '513
    
    D = Wdata
    sLeft = Main.Picfft1.ScaleWidth - 128
    sLeft = sLeft / 2
    ' This should be even number between 2 and 10
    Main.Picfft1.Cls
    For a = 0 To 128 Step sSize
        z = (D(a) * 1000)
        If z > 10 Then z = z + VisGain
        If z > 40 Then z = 40
        If z > Peak(a) Then
            Peak(a) = z
        Else
            If Peak(a) > 45 Then
                Peak(a) = Peak(a) - 2
            Else
                Peak(a) = Peak(a) - 1
            End If
        End If
        ' Draw blue
         Main.Picfft1.Line (sLeft + a, 45)-(sLeft + a + (sSize / 2), 45 - z), &HC0C000, BF
        ' Draw blue/red
        TopLevel = 10
        If z > TopLevel Then
             Main.Picfft1.Line (sLeft + a, 45 - TopLevel)-(sLeft + a + (sSize / 2), 45 - z), &HC0C000, BF
        End If
        ' Draw red
        TopLevel = 20
        If z > TopLevel Then
             Main.Picfft1.Line (sLeft + a, 45 - TopLevel)-(sLeft + a + (sSize / 2), 45 - z), &HC0C000, BF              '&HFFFF&, BF
        End If
        ' Draw red/orange
        TopLevel = 30
        If z > TopLevel Then
             Main.Picfft1.Line (sLeft + a, 45 - TopLevel)-(sLeft + a + (sSize / 2), 45 - z), &HC0C000, BF            '&H69CAFE, BF
        End If
        ' Draw Orange
        TopLevel = 35
        If z > TopLevel Then
             Main.Picfft1.Line (sLeft + a, 45 - TopLevel)-(sLeft + a + (sSize / 2), 45 - z), &HC0C000, BF            '&H80FF&, BF   'top
        End If
        ' Draw orange/yellow
        TopLevel = 40
        If z > TopLevel Then
             Main.Picfft1.Line (sLeft + a, 45 - TopLevel)-(sLeft + a + (sSize / 2), 45 - z), &HC0C000, BF
        End If
         Main.Picfft1.Line (sLeft + a, 45 - Peak(a))-(sLeft + a + (sSize / 2), 45 - Peak(a)), &H808000, BF
        sCount = sCount + 1
    Next

End Sub

Public Sub SetStatusText(WStatText As String)

Main.Label2.Caption = WStatText

End Sub

Public Sub SetStreamName(WStrName As String)

Main.Label1.Caption = WStrName

End Sub

Public Sub UnloadControl()

Main.TmrClock.Interval = 0
Main.TmrClock.Enabled = False

Unload Main

End Sub

