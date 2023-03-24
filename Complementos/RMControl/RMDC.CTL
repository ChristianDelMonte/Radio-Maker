VERSION 5.00
Begin VB.UserControl RMDC 
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2160
   ScaleHeight     =   870
   ScaleWidth      =   2160
   ToolboxBitmap   =   "RMDC.ctx":0000
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      Picture         =   "RMDC.ctx":0312
      ScaleHeight     =   600
      ScaleWidth      =   2130
      TabIndex        =   0
      Top             =   0
      Width           =   2130
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "DATA CONTROL"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   630
      Width           =   2175
   End
End
Attribute VB_Name = "RMDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'****************** RMDC ********************
'       RADIO MAKER DATA CONTROL OCX
'COPYRIGHT (C) 2001 CREACIONES DIGITALES INC.
'********************************************

Option Explicit

Private Const MIN_ASC = 1
Private Const MAX_ASC = 255
Private Const NUM_ASC = MAX_ASC - MIN_ASC + 1

Function DecipherData(WPass As String, InText As String) As String

Dim offset As Long
Dim str_len As Integer
Dim i As Integer
Dim ch As Integer
Dim OutText As String

    'chequeos necesarios
    If WPass = "" Or WPass = " " Then
        DecipherData = ""
        Exit Function
    End If
    If InText = "" Or InText = " " Then
        DecipherData = ""
        Exit Function
    End If
    
    'Inicializar el generador de numeros aleatorios
    OutText = ""
    offset = CN(WPass)
    Rnd -1
    Randomize offset
    
    'Desencriptar el texto
    str_len = Len(InText)
    For i = 1 To str_len
        ch = Asc(Mid$(InText, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch - offset) Mod NUM_ASC)
            If ch < 0 Then ch = ch + NUM_ASC
            ch = ch + MIN_ASC
            OutText = OutText & Chr$(ch)
        End If
    Next i

    DecipherData = OutText

End Function

Function CipherData(WPass As String, InText As String) As String

Dim offset As Long
Dim str_len As Integer
Dim i As Integer
Dim ch As Integer
Dim OutText As String

    'chequeos necesarios
    If WPass = "" Or WPass = " " Then
        CipherData = ""
        Exit Function
    End If
    If InText = "" Or InText = " " Then
        CipherData = ""
        Exit Function
    End If

    'Inicializar el generador de numeros aleatorios
    OutText = ""
    offset = CN(WPass)
    Rnd -1
    Randomize offset
    
    'Encriptar el texto
    str_len = Len(InText)
    For i = 1 To str_len
        ch = Asc(Mid$(InText, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch + offset) Mod NUM_ASC)
            ch = ch + MIN_ASC
            OutText = OutText & Chr$(ch)
        End If
    Next i

    CipherData = OutText
    
End Function

Private Function CN(ByVal WnPass As String) As Long

Dim value As Long
Dim ch As Long
Dim shift1 As Long
Dim shift2 As Long
Dim i As Integer
Dim str_len As Integer

    str_len = Len(WnPass)
    For i = 1 To str_len
        ch = Asc(Mid$(WnPass, i, 1))
        value = value Xor (ch * 2 ^ shift1)
        value = value Xor (ch * 2 ^ shift2)
        
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    
    CN = value

End Function

Function GetInfoData(WInfoType As String, WInfoVal As String) As String

'Esta function devuelve un tipo especifico de dato
'de acuerdo a lo especificado por el usuario y de
'acuerdo a la clave de acceso utilizada.
'Los datos contenidos en esta funcion deben estar
'encriptados para seguridad.
Dim DataCh As String
Dim DataDch As String

'chequeos necesarios
If WInfoType = "" Or WInfoType = " " Then Exit Function
If WInfoVal = "" Or WInfoVal = " " Then Exit Function

'seleccion de WInfoType: ¿Que dato desea obtener el usuario?
Select Case WInfoType
    Case "x1"
        'DataCh = "xxx"  'dato encryptado
        'DecipherData WInfoVal, DataCh, DataDch  'desencriptamos
        GetInfoData = "NotImplemented yet..."   'devolvemos el dato desencriptado al usuario
    Case "x2"
        'xxx
    Case "x3"
        'xxx
    Case "x4"
        'xxx
    Case "x5"
        'xxx
    Case "x6"
        'xxx
    Case "x7"
        'xxx
    Case "x8"
        'xxx
    Case "x9"
        'xxx
    Case Else
        GetInfoData = "NotImplemented Yet..."
        'xxx
End Select

End Function
