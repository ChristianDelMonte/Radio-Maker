Attribute VB_Name = "Localize"

'***********************************************'
'Modulo de Nacionalizacion para ONLY RMLocalize '
'-----------------------------------------------'
'   Este componente forma parte del modulo de   '
' nacionalizacion de sistema de ONLY Radiomaker '
' para la traduccion de componentes a distintos '
'idiomas. Compatibles con productos ONLY unicos '
'-----------------------------------------------'
'       Copyright (c) ONLY development 2008     '
'***********************************************'
' Rev: 21-02-2008 '

Option Explicit

Public Const CFG_File = "\CFG_Data.dat"         'archivo de configuracion
Public Const LNG_File = "\RMLocalize_Data.lng"         'archivos de datos de nacionalizacion

Public Type CFG_Tipo
    Id As Integer
    Lenguaje As String * 50    'lenguaje
    LNG_Predet As Integer       '=1 si es predeterminado o =0 si no lo es
End Type

Public Type LNG_Tipo
    Id As Integer               'identificador
    LNG_ID As Integer           'identificador de lenguaje (vease LNG_Config file)
    PRG_ID As Integer           'identificador de programa o componente a nacionalizar
    PRG_Desc As String * 360    'componente o descripcion en lenguaje ya establecido
    LNG_Comm As String * 255    'comentarios adicionales
End Type

Public LenguajeData As LNG_Tipo
Public ConfigData As CFG_Tipo
Public lastreg As Integer

'// extraer el numero de ultimo registro del archivo de configuracion
Public Function GetCFGLastReg() As Integer

'/// abrimos el archivo
On Error GoTo err
Open App.Path & CFG_File For Random As #24 Len = Len(ConfigData)

'/// check for the last reg ID
lastreg = LOF(24) \ Len(ConfigData)

GetCFGLastReg = lastreg

'///end the function
Close #24
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en GetCFGLastReg > Module localize - " & err.Number & " - " & err.Description
GetCFGLastReg = -1
Close #24
End Function

'// extraer el numero de ultimo registro del archivo de nacionalizacion
Public Function GetLNGLastReg() As Integer

'/// abrimos el archivo
On Error GoTo err
Open App.Path & LNG_File For Random As #24 Len = Len(LenguajeData)

'/// check for the last reg ID
lastreg = LOF(24) \ Len(LenguajeData)

GetLNGLastReg = lastreg

'///end the function
Close #24
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en GetLNGLastReg > Module localize - " & err.Number & " - " & err.Description
GetLNGLastReg = -1
Close #24
End Function

'funcion para listar los lenguajes del archivo configuracion
Public Function GetCFGConfigList(WOptionalID As Integer) As CFG_Tipo

Dim i As Integer

On Error GoTo err
Open App.Path & CFG_File For Random As #12 Len = Len(ConfigData)

If WOptionalID = 0 Or WOptionalID = -1 Then
    GetCFGConfigList.Id = -1
    Close #12
    Exit Function
Else
    Get #12, WOptionalID, ConfigData
    GetCFGConfigList = ConfigData
    Close #12
    Exit Function
End If

Close #12
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en GetLNGConfigList > Module localize - " & err.Number & " - " & err.Description
Close #12
GetCFGConfigList.Id = -1
End Function

'extraer y devolver el ID de un lenguaje especifico
Public Function GetIDFromCFGConfig(WLng As String) As Integer

Dim i As Integer

On Error GoTo err
Open App.Path & CFG_File For Random As #12 Len = Len(ConfigData)

lastreg = GetCFGLastReg
lastreg = lastreg + 1

For i = 1 To lastreg
    Get #12, i, ConfigData
    If UCase(Trim(ConfigData.Lenguaje)) = UCase(Trim(WLng)) Then
        GetIDFromCFGConfig = ConfigData.Id
        Close #12
        Exit For: Exit Function
    Else
        If i = lastreg Then
            GetIDFromCFGConfig = -1 'ERROR!! no se encontro ninguna coincidencia
            Close #12
            Exit For: Exit Function
        End If
    End If
Next i

Close #12
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en GetIDFromLNGConfig > Module localize - " & err.Number & " - " & err.Description
Close #12
GetIDFromCFGConfig = -1
End Function

'funcion para listar todas las funciones dentro del archivo de nacionalizacion
Public Function GetLNGDataList(WLngId As Integer, WOptionalID As Integer) As LNG_Tipo

Dim i As Integer

'/// abrimos el archivo de datos
On Error GoTo err
Open App.Path & LNG_File For Random As #11 Len = Len(LenguajeData)

'/// check for the ID to load
If WLngId = 0 Or WLngId = -1 Then
    GetLNGDataList.Id = -1  'ERROR!! no se especifico el id del lenguaje. no se puede continuar
    Close #11: Exit Function
Else
    If WOptionalID = 0 Or WOptionalID = -1 Then
        GetLNGDataList.Id = -1  'ERROR!! no se especifico el id de registro. no se puede continuar
        Close #11: Exit Function
    Else
        Get #11, WOptionalID, LenguajeData
        If LenguajeData.LNG_ID = WLngId Then    'verificamos que extraiga solo los datos del idioma predet.
            GetLNGDataList = LenguajeData
            Close #11: Exit Function
        End If
    End If
End If

Close #11
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en GetLNGDataList > Module localize - " & err.Number & " - " & err.Description
Close #11
GetLNGDataList.Id = -1
End Function

'// extraer los datos de nacionalizacion especificos segun el lenguaje seleccionado
Public Function GetLNGData2(WLngId As Integer, WPrgId As Integer) As LNG_Tipo

Dim i As Integer

'/// abrimos el archivo de datos
On Error GoTo err
Open App.Path & LNG_File For Random As #11 Len = Len(LenguajeData)

'/// check for the ID to load
If WLngId = 0 Or WLngId = -1 Then
    GetLNGData2.Id = -1  'ERROR!! no se especifico el id del lenguaje. no se puede continuar
    Close #11: Exit Function
Else
    If WPrgId = 0 Or WPrgId = -1 Then
        GetLNGData2.Id = -1  'ERROR!! no se especifico el id del componente a leer nacionalizacion
        Close #11: Exit Function
    Else
        lastreg = GetLNGLastReg
        lastreg = lastreg + 1
        For i = 1 To lastreg
            Get #11, i, LenguajeData
            If LenguajeData.LNG_ID = WLngId And LenguajeData.PRG_ID = WPrgId Then
                GetLNGData2 = LenguajeData
                Close #11
                Exit For: Exit Function
            Else
                If i = lastreg Then
                    GetLNGData2.Id = -1  'ERROR!! no se encontro ninguna coincidencia de nacionalizacion
                    Close #11
                    Exit For: Exit Function
                End If
            End If
        Next i
    End If
End If

Close #11
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en GetLNGData > Module localize - " & err.Number & " - " & err.Description
Close #11
GetLNGData2.Id = -1
End Function
'// extraer los datos de nacionalizacion especificos segun el lenguaje seleccionado
Public Function GetLNGData(WLngId As Integer, WPrgId As Integer) As String

Dim i As Integer

'/// abrimos el archivo de datos
On Error GoTo err
Open App.Path & LNG_File For Random As #11 Len = Len(LenguajeData)

'/// check for the ID to load
If WLngId = 0 Or WLngId = -1 Then
    GetLNGData = -1  'ERROR!! no se especifico el id del lenguaje. no se puede continuar
    Close #11: Exit Function
Else
    If WPrgId = 0 Or WPrgId = -1 Then
        GetLNGData = -1  'ERROR!! no se especifico el id del componente a leer nacionalizacion
        Close #11: Exit Function
    Else
        lastreg = GetLNGLastReg
        lastreg = lastreg + 1
        For i = 1 To lastreg
            Get #11, i, LenguajeData
            If LenguajeData.LNG_ID = WLngId And LenguajeData.PRG_ID = WPrgId Then
                GetLNGData = Trim(LenguajeData.PRG_Desc)
                Close #11
                Exit For: Exit Function
            Else
                If i = lastreg Then
                    GetLNGData = -1  'ERROR!! no se encontro ninguna coincidencia de nacionalizacion
                    Close #11
                    Exit For: Exit Function
                End If
            End If
        Next i
    End If
End If

Close #11
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en GetLNGData > Module localize - " & err.Number & " - " & err.Description
Close #11
GetLNGData = -1
End Function

'//guardar los datos de nacionalizacion especificos segun el lenguaje especificado
Public Function SaveCFGData(WData As CFG_Tipo, WOptionalID As Integer, IsPredet As Boolean) As Boolean

'/// abrimos el archivo de productos
On Error GoTo err
Open App.Path & CFG_File For Random As #12 Len = Len(ConfigData)

'/// check for the ID to save
If WOptionalID = 0 Or WOptionalID = -1 Then
    lastreg = LOF(12) \ Len(ConfigData)
    lastreg = lastreg + 1
Else
    lastreg = WOptionalID
End If

'/// seteamos los datos del inventario a guardar
ConfigData.Id = lastreg
ConfigData.Lenguaje = WData.Lenguaje
If IsPredet = True Then
    ConfigData.LNG_Predet = 1
Else
    ConfigData.LNG_Predet = 0
End If

'/// guardamos
Put #12, lastreg, ConfigData
Close #12

SaveCFGData = True
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en SaveCFGData > Module localize - " & err.Number & " - " & err.Description
Close #12
SaveCFGData = False
End Function

'//guardar los datos de nacionalizacion especificos segun el lenguaje especificado
Public Function SaveLNGData(WData As LNG_Tipo, WOptionalID As Integer) As Boolean

'/// abrimos el archivo de productos
On Error GoTo err
Open App.Path & LNG_File For Random As #12 Len = Len(LenguajeData)

'/// check for the ID to save
If WOptionalID = 0 Or WOptionalID = -1 Then
    lastreg = LOF(12) \ Len(LenguajeData)
    lastreg = lastreg + 1
Else
    lastreg = WOptionalID
End If

'/// seteamos los datos del inventario a guardar
LenguajeData.Id = lastreg
LenguajeData.LNG_ID = WData.LNG_ID
LenguajeData.PRG_ID = WData.PRG_ID
LenguajeData.PRG_Desc = Trim(WData.PRG_Desc)
LenguajeData.LNG_Comm = Trim(WData.LNG_Comm)

'/// guardamos
Put #12, lastreg, LenguajeData
Close #12

SaveLNGData = True
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en SaveLNGData > Module localize - " & err.Number & " - " & err.Description
Close #12
SaveLNGData = False
End Function

Public Function ImportTextFile(WFileName As String, IDPredet As Integer) As Boolean

Dim DataId As String, Data As String

On Error GoTo err
Open WFileName For Input As #30
Do Until EOF(30)
Line Input #30, Data
    DataId = Left$(Data, 4)
    Data = Mid$(Data, 5, Len(Data))
    'MsgBox "ID=" & Trim(DataId) & " - data=" & Trim(Data)
    
    LenguajeData.LNG_ID = IDPredet
    LenguajeData.PRG_ID = CInt(DataId)
    LenguajeData.PRG_Desc = Trim(Data)
    LenguajeData.LNG_Comm = "vacio"
    
    SaveLNGData LenguajeData, 0
    
Loop
Close #30
Exit Function

err:
ImportTextFile = False
End Function
