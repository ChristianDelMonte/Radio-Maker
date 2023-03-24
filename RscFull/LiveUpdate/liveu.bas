Attribute VB_Name = "LiveUpd"
'****************************************************************
'
' Live Program Update Code
'
' Written by:  Blake B. Pell
'              blakepell@hotmail.com
'              bpell@indiana.edu
'              http://www.blakepell.com
'              December 7, 2000
'
' This code is open source, I would appreciate that anybody using
' this is a released application to e-mail or get in contact with
' me.  I hope this makes someone's day easier or helps them learn
' a bit.
'
'
'****************************************************************

Global myVer As String, myPath As String
Global status$
Global UpdateTime As Integer

Public Function GetInternetFile(Inet1 As Inet, myURL As String, DestDIR As String) As Boolean
    ' Written by: Blake Pell
    
    On Local Error GoTo er
    
    Dim myData() As Byte
    If Inet1.StillExecuting = True Then Exit Function
    myData() = Inet1.OpenURL(myURL, icByteArray)


    For X = Len(myURL) To 1 Step -1
        If Left$(Right$(myURL, X), 1) = "/" Then RealFile$ = Right$(myURL, X - 1)
    Next X
    
    myFile$ = DestDIR + "\" + RealFile$
    
    Open myFile$ For Binary Access Write As #12
    Put #12, , myData()
    Close #12
    
    GetInternetFile = True
    Exit Function

' error handler
er:
X = MsgBox("Ha ocurrido un error al intentar descargar el archivo.  Por favor intente nuevamente mas tarde.", vbInformation)
GetInternetFile = False
End Function
