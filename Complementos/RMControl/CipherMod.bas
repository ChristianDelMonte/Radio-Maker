Attribute VB_Name = "ChipherMod"
' Copyright (c) 2001 Creaciones Digitales Inc.

Option Explicit

Public Const Rm100Ver = "1.001"
Public Const Rm100Name = "Radio Maker OCX Control"
Public Const Rm100Autor = "Christian A. Del Monte"
Public Const Rm100Copyr = "(c) 2001 Creaciones Digitales Inc."
Public Const WWDefPass = "1.001a"

Private Const MIN_ASC = 1
Private Const MAX_ASC = 255
Private Const NUM_ASC = MAX_ASC - MIN_ASC + 1

Sub Desencriptar(ByVal Password As String, ByVal TxtOrigen As String, TxtDestino As String)

Dim offset As Long
Dim str_len As Integer
Dim i As Integer
Dim ch As Integer

    'Inicializar el generador de numeros aleatorios
    offset = ClaveNumerica(Password)
    Rnd -1
    Randomize offset
    
    'Desencriptar el texto
    str_len = Len(TxtOrigen)
    For i = 1 To str_len
        ch = Asc(Mid$(TxtOrigen, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch - offset) Mod NUM_ASC)
            If ch < 0 Then ch = ch + NUM_ASC
            ch = ch + MIN_ASC
            TxtDestino = TxtDestino & Chr$(ch)
        End If
    Next i

End Sub

Sub Encriptar(ByVal Password As String, ByVal TxtOrigen As String, TxtDestino As String)

Const MIN_ASC = 1
Const MAX_ASC = 255
Const NUM_ASC = MAX_ASC - MIN_ASC + 1

Dim offset As Long
Dim str_len As Integer
Dim i As Integer
Dim ch As Integer

    'Inicializar el generador de numeros aleatorios
    offset = ClaveNumerica(Password)
    Rnd -1
    Randomize offset
    
    'Encriptar el texto
    str_len = Len(TxtOrigen)
    For i = 1 To str_len
        ch = Asc(Mid$(TxtOrigen, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch + offset) Mod NUM_ASC)
            ch = ch + MIN_ASC
            TxtDestino = TxtDestino & Chr$(ch)
        End If
    Next i

End Sub

Function ClaveNumerica(ByVal Password As String) As Long

Dim value As Long
Dim ch As Long
Dim shift1 As Long
Dim shift2 As Long
Dim i As Integer
Dim str_len As Integer

    str_len = Len(Password)
    For i = 1 To str_len
        ch = Asc(Mid$(Password, i, 1))
        value = value Xor (ch * 2 ^ shift1)
        value = value Xor (ch * 2 ^ shift2)
        
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    ClaveNumerica = value

End Function

