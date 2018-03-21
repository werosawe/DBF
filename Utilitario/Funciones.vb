Option Strict On
Option Explicit On
Imports System
Imports System.Web.Script.Services
Imports System.Web.Script.Serialization
Imports System.Runtime.Serialization.Json
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.IO
Imports System.Diagnostics
Imports System.Web.UI.WebControls
Imports System.Reflection

Public Enum UTILITARIO_MUESTRA_FECHA As Integer
    CLIENTE_FECHA_24HORA = 1
    CLIENTE_FECHA_12HORA
    CLIENTE_FECHA
    CLIENTE_12HORA
    CLIENTE_24HORA
    CLIENTE_DIA_NUMERO
    CLIENTE_MES_NUMERO
    CLIENTE_ANNO_NUMERO
    CLIENTE_DIA_NOMBRE
    CLIENTE_MES_NOMBRE
End Enum

Public Class Funciones
    'Implements IComparer(Of ListItem)
    'Public Function Compare(ByVal x As ListItem, ByVal y As ListItem) As Integer _
    '    Implements IComparer(Of ListItem).Compare
    '    Dim c As New CaseInsensitiveComparer
    '    Return c.Compare(x.Text, y.Text)
    'End Function
    Public Shared Function Traslado(ByVal XstrCad As String) As String
        Dim strResp As String = ""
        If Not XstrCad Is System.DBNull.Value And Not XstrCad Is Nothing Then
            strResp = XstrCad
            strResp = strResp.Replace("á", "&#225;")
            strResp = strResp.Replace("é", "&#233;")
            strResp = strResp.Replace("í", "&#237;")
            strResp = strResp.Replace("ó", "&#243;")
            strResp = strResp.Replace("ú", "&#250;")
            strResp = strResp.Replace("ñ", "&#241;")
            strResp = strResp.Replace("Ñ", "&#209;")
            strResp = strResp.Replace("ü", "&#252;")
            strResp = strResp.Replace("Ü", "&#220;")
            strResp = strResp.Replace("'", "&#96;")
            strResp = strResp.Replace("Á", "&#193;")
            strResp = strResp.Replace("É", "&#201;")
            strResp = strResp.Replace("Í", "&#205;")
            strResp = strResp.Replace("Ó", "&#211;")
            strResp = strResp.Replace("Ú", "&#218;")
            strResp = strResp.Replace("°", "&#176;")
            strResp = strResp.Replace("º", "&#186;")
        Else
            strResp = ""
        End If
        Return strResp
    End Function
    Public Shared Function Traslado_Invertido(ByVal XstrCad As String) As String
        Dim strResp As String = ""
        If Not XstrCad Is System.DBNull.Value Then ''
            strResp = XstrCad
            strResp = strResp.Replace("&#225;", "á")
            strResp = strResp.Replace("&#233;", "é")
            strResp = strResp.Replace("&#237;", "í")
            strResp = strResp.Replace("&#243;", "ó")
            strResp = strResp.Replace("&#250;", "ú")
            strResp = strResp.Replace("&#241;", "ñ")
            strResp = strResp.Replace("&#209;", "Ñ")
            strResp = strResp.Replace("&#252;", "ü")
            strResp = strResp.Replace("&#220;", "Ü")
            strResp = strResp.Replace("&#96;", "'")
            strResp = strResp.Replace("&#211;", "Ó")
            strResp = strResp.Replace("&#193;", "Á")
            strResp = strResp.Replace("&#201;", "É")
            strResp = strResp.Replace("&#205;", "Í")
            strResp = strResp.Replace("&#218;", "Ú")
            strResp = strResp.Replace("&#176;", "°")
            strResp = strResp.Replace("&#186;", "º")
        End If
        Return strResp
    End Function

    Public Shared Function Dame_Texto(ByVal XobjValue As Object) As String
        'If XobjValue Is System.DBNull.Value Then
        '    Return Nothing
        'ElseIf XobjValue Is Nothing Then
        '    Return Nothing
        'ElseIf XobjValue.ToString.Trim.Equals("") Then
        '    Return Nothing
        'End If
        'Return Convert.ToString(XobjValue)
        If XobjValue Is System.DBNull.Value Then
            Return ""
        ElseIf XobjValue Is Nothing Then
            Return ""
        ElseIf XobjValue.ToString.Trim.Equals("") Then
            Return ""
        End If
        Return Convert.ToString(XobjValue)
    End Function

    Public Shared Function Dame_Entero_Decimal(ByVal XobjValue As Object) As Double
        If XobjValue Is System.DBNull.Value Then
            Return 0
        ElseIf XobjValue Is Nothing Then
            Return 0
        ElseIf XobjValue.ToString.Trim.Equals("") Then
            Return 0
        End If
        
        Return CType(XobjValue, Double)
    End Function
    Public Shared Function Dame_Entero(ByVal XobjValue As Object) As Integer
        If XobjValue Is System.DBNull.Value Then
            Return 0
        ElseIf XobjValue Is Nothing Then
            Return 0
        ElseIf XobjValue.ToString.Trim.Equals("") Then
            Return 0
        End If
        Return Convert.ToInt32(XobjValue)
    End Function

    'Public Shared Function Convert_dd_mm_yyyy(ByVal XobjValue As String) As Nullable(Of Date)

    '    If XobjValue Is System.DBNull.Value Then
    '        Return Nothing
    '    ElseIf XobjValue Is Nothing Then
    '        Return Nothing
    '    ElseIf XobjValue.ToString.Trim.Equals("") Then
    '        Return Nothing
    '    End If
    '    Try
    '        Dim arrMonth() As String = {"ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"}
    '        Dim iMonth As Integer = arrMonth.ToList().IndexOf(XobjValue.Substring(3, 3).ToUpper) + 1
    '        Dim f As String = XobjValue.Substring(0, 2) + "/" + iMonth.ToString.PadLeft(2, CChar("0")) + "/" + XobjValue.Substring(7, 4)
    '        Return CDate(f)
    '    Catch ex As Exception
    '        Return CDate(XobjValue)
    '    Finally

    '    End Try


    'End Function

    Public Shared Function Dame_Fecha_dd_MMM_yyyy(ByVal XobjValue As Object) As Nullable(Of Date)

        If XobjValue Is System.DBNull.Value Then
            Return Nothing
        ElseIf XobjValue Is Nothing Then
            Return Nothing
        ElseIf XobjValue.ToString.Trim.Equals("") Then
            Return Nothing
        End If

        Dim f As String = String.Format("{0:dd-MMM-yyy}", XobjValue)

        Dim dtResult As Date, xdate As Date

        Try
            If DateTime.TryParse(f, dtResult) Then
                xdate = CDate(dtResult.ToString("f", System.Globalization.CultureInfo.InvariantCulture))
            End If

            Return xdate
        Catch ex As Exception
            Return CDate(XobjValue)
        Finally

        End Try

    End Function

    Public Shared Function Dame_Boolean(ByVal XobjValue As Object) As Boolean
        If XobjValue Is System.DBNull.Value Then
            Return Nothing
        ElseIf XobjValue Is Nothing Then
            Return Nothing
        ElseIf XobjValue.ToString.Trim.Equals("") Then
            Return Nothing
        ElseIf XobjValue.ToString.ToUpper = "FALSE" Then
            Return False
        ElseIf XobjValue.ToString.ToUpper = "TRUE" Then
            Return True
        ElseIf Convert.ToInt32(XobjValue) = 0 Then
            Return False
        ElseIf Convert.ToInt32(XobjValue) = 1 Then
            Return True
        End If
        Return CType(XobjValue, Boolean)
    End Function

    Public Shared Function Dame_ListCount(ByVal XobjValue As Object) As Integer

        If XobjValue Is System.DBNull.Value Then
            Return 0
        ElseIf XobjValue Is Nothing Then
            Return 0
        ElseIf XobjValue.ToString.Trim.Equals("") Then
            Return 0
        End If
        Return DirectCast(XobjValue, IList).Count
    End Function

    Public Shared Function Fecha_BD(ByVal XobjValue As Object) As String
        If XobjValue Is System.DBNull.Value Then
            Return Nothing
        ElseIf XobjValue Is Nothing Then
            Return Nothing
        ElseIf XobjValue.ToString.Trim.Equals("") = True Then
            Return Nothing
        Else
            Return XobjValue.ToString.Substring(0, 2) & "/" & XobjValue.ToString.Substring(3, 2) & "/" & XobjValue.ToString.Substring(6, 4)
        End If
    End Function

    Public Shared Function Hora_Sistema() As String
        Dim tx_hora As String = Date.Now.Hour.ToString("00")
        Dim txampm As String = Date.Now.ToString("tt").Replace(".", "").ToUpper
        Dim fl_cambio As Boolean = False
        If tx_hora.ToString.Equals("13") = True Then tx_hora = "01" : fl_cambio = True
        If tx_hora.ToString.Equals("14") = True Then tx_hora = "02" : fl_cambio = True
        If tx_hora.ToString.Equals("15") = True Then tx_hora = "03" : fl_cambio = True
        If tx_hora.ToString.Equals("16") = True Then tx_hora = "04" : fl_cambio = True
        If tx_hora.ToString.Equals("17") = True Then tx_hora = "05" : fl_cambio = True
        If tx_hora.ToString.Equals("18") = True Then tx_hora = "06" : fl_cambio = True
        If tx_hora.ToString.Equals("19") = True Then tx_hora = "07" : fl_cambio = True
        If tx_hora.ToString.Equals("20") = True Then tx_hora = "08" : fl_cambio = True
        If tx_hora.ToString.Equals("21") = True Then tx_hora = "09" : fl_cambio = True
        If tx_hora.ToString.Equals("22") = True Then tx_hora = "10" : fl_cambio = True
        If tx_hora.ToString.Equals("23") = True Then tx_hora = "11" : fl_cambio = True
        If tx_hora.ToString.Equals("00") = True Then tx_hora = "12" : fl_cambio = True
        If fl_cambio = True Then
            tx_hora = tx_hora & ":" & Date.Now.Minute.ToString("00") & txampm
        Else
            tx_hora = tx_hora & ":" & Date.Now.Minute.ToString("00") & txampm
        End If
        Return tx_hora
    End Function
    Public Shared Function Fecha_Hora_BD(ByVal XobjValueFecha As Object, Optional ByVal XobjValueHora As Object = Nothing) As String
        If XobjValueFecha Is System.DBNull.Value And XobjValueHora Is System.DBNull.Value Then
            Return ""
        ElseIf XobjValueFecha Is Nothing And XobjValueHora Is Nothing Then
            Return ""
        ElseIf XobjValueFecha.ToString.Trim.Length = 10 And XobjValueHora Is Nothing Then
            Return XobjValueFecha.ToString.Substring(6, 4) & XobjValueFecha.ToString.Substring(3, 2) & XobjValueFecha.ToString.Substring(0, 2) & "000000"
        ElseIf XobjValueFecha.ToString.Trim.Equals("") And XobjValueHora.ToString.Trim.Equals("") Then
            Return ""
        ElseIf XobjValueHora.ToString.Trim.Equals("") = True And XobjValueFecha.ToString.Trim.Length = 10 Then
            Return XobjValueFecha.ToString.Substring(6, 4) & XobjValueFecha.ToString.Substring(3, 2) & XobjValueFecha.ToString.Substring(0, 2) & "000000"
        ElseIf XobjValueHora.ToString.Trim.Length <> 7 And XobjValueFecha.ToString.Trim.Length = 10 Then
            Return XobjValueFecha.ToString.Substring(6, 4) & XobjValueFecha.ToString.Substring(3, 2) & XobjValueFecha.ToString.Substring(0, 2) & "000000"
        ElseIf XobjValueFecha.ToString.Trim.Length = 10 And XobjValueHora.ToString.Trim.Length = 7 Then
            Dim tx_hora As String = XobjValueHora.ToString.Substring(0, 2)
            Dim tx_minuto As String = XobjValueHora.ToString.Substring(3, 2)
            Dim tx_segundo As String = "00"
            Dim tx_AMPM As String = XobjValueHora.ToString.Substring(5, 2)
            If XobjValueHora.ToString.Trim.Length = 10 Then
                tx_segundo = XobjValueHora.ToString.Substring(6, 2)
            End If
            If tx_AMPM.Equals("PM") Then
                If tx_hora.Equals("01") = True Then tx_hora = "13"
                If tx_hora.Equals("02") = True Then tx_hora = "14"
                If tx_hora.Equals("03") = True Then tx_hora = "15"
                If tx_hora.Equals("04") = True Then tx_hora = "16"
                If tx_hora.Equals("05") = True Then tx_hora = "17"
                If tx_hora.Equals("06") = True Then tx_hora = "18"
                If tx_hora.Equals("07") = True Then tx_hora = "19"
                If tx_hora.Equals("08") = True Then tx_hora = "20"
                If tx_hora.Equals("09") = True Then tx_hora = "21"
                If tx_hora.Equals("10") = True Then tx_hora = "22"
                If tx_hora.Equals("11") = True Then tx_hora = "23"
                'If tx_hora.Equals("12") = True Then tx_hora = "00"
            ElseIf tx_AMPM.Equals("AM") Then
                If tx_hora.Equals("12") = True Then tx_hora = "00"
            End If
            Return XobjValueFecha.ToString.Substring(6, 4) & XobjValueFecha.ToString.Substring(3, 2) & XobjValueFecha.ToString.Substring(0, 2) & tx_hora & tx_minuto & tx_segundo
        Else
            Return ""
        End If
    End Function

    Public Shared Function Fecha(ByVal XobjValue As Object, Optional ByVal _Mostrar As UTILITARIO_MUESTRA_FECHA = UTILITARIO_MUESTRA_FECHA.CLIENTE_FECHA_12HORA, Optional ByVal FL_SEGUNDO As Boolean = False) As String
        If XobjValue Is System.DBNull.Value Then Return ""
        If XobjValue Is Nothing Then Return ""
        If XobjValue.ToString.Equals("") = True Then Return ""
        If _Mostrar = Nothing Then Return ""

        If XobjValue.ToString.Trim.Length = 14 Then
            Dim tx_dia As String = XobjValue.ToString.Substring(6, 2)
            Dim tx_mes As String = XobjValue.ToString.Substring(4, 2)
            Dim tx_anno As String = XobjValue.ToString.Substring(0, 4)

            Dim tx_fecha_hora As String = ""
            Dim tx_hora As String = XobjValue.ToString.Substring(8, 2)
            Dim tx_minuto As String = XobjValue.ToString.Substring(10, 2)
            Dim tx_segundo As String = XobjValue.ToString.Substring(12, 2)

            If _Mostrar = UTILITARIO_MUESTRA_FECHA.CLIENTE_24HORA Or _Mostrar = UTILITARIO_MUESTRA_FECHA.CLIENTE_FECHA_24HORA Then
                If FL_SEGUNDO = True Then
                    tx_fecha_hora = tx_hora & ":" & tx_minuto & ":" & tx_segundo
                Else
                    tx_fecha_hora = tx_hora & ":" & tx_minuto
                End If
            ElseIf _Mostrar = UTILITARIO_MUESTRA_FECHA.CLIENTE_12HORA Or _Mostrar = UTILITARIO_MUESTRA_FECHA.CLIENTE_FECHA_12HORA Then
                Dim TX_AM_PM As String = "AM"
                If tx_hora.Equals("12") = True Then TX_AM_PM = "PM"
                If tx_hora.Equals("13") = True Then tx_hora = "01" : TX_AM_PM = "PM"
                If tx_hora.Equals("14") = True Then tx_hora = "02" : TX_AM_PM = "PM"
                If tx_hora.Equals("15") = True Then tx_hora = "03" : TX_AM_PM = "PM"
                If tx_hora.Equals("16") = True Then tx_hora = "04" : TX_AM_PM = "PM"
                If tx_hora.Equals("17") = True Then tx_hora = "05" : TX_AM_PM = "PM"
                If tx_hora.Equals("18") = True Then tx_hora = "06" : TX_AM_PM = "PM"
                If tx_hora.Equals("19") = True Then tx_hora = "07" : TX_AM_PM = "PM"
                If tx_hora.Equals("20") = True Then tx_hora = "08" : TX_AM_PM = "PM"
                If tx_hora.Equals("21") = True Then tx_hora = "09" : TX_AM_PM = "PM"
                If tx_hora.Equals("22") = True Then tx_hora = "10" : TX_AM_PM = "PM"
                If tx_hora.Equals("23") = True Then tx_hora = "11" : TX_AM_PM = "PM"
                If tx_hora.Equals("00") = True Then tx_hora = "12"



                If FL_SEGUNDO = True Then
                    tx_fecha_hora = tx_hora & ":" & tx_minuto & ":" & tx_segundo & TX_AM_PM
                Else
                    tx_fecha_hora = tx_hora & ":" & tx_minuto & TX_AM_PM
                End If


            End If

            If _Mostrar = UTILITARIO_MUESTRA_FECHA.CLIENTE_FECHA Then
                Return tx_dia & "/" & tx_mes & "/" & tx_anno
            ElseIf _Mostrar = UTILITARIO_MUESTRA_FECHA.CLIENTE_FECHA_12HORA Then
                Return tx_dia & "/" & tx_mes & "/" & tx_anno & " " & tx_fecha_hora
            ElseIf _Mostrar = UTILITARIO_MUESTRA_FECHA.CLIENTE_FECHA_24HORA Then
                Return tx_dia & "/" & tx_mes & "/" & tx_anno & " " & tx_fecha_hora
            ElseIf _Mostrar = UTILITARIO_MUESTRA_FECHA.CLIENTE_DIA_NUMERO Then
                Return tx_dia
            ElseIf _Mostrar = UTILITARIO_MUESTRA_FECHA.CLIENTE_MES_NUMERO Then
                Return tx_dia
            ElseIf _Mostrar = UTILITARIO_MUESTRA_FECHA.CLIENTE_ANNO_NUMERO Then
                Return tx_dia
            ElseIf _Mostrar = UTILITARIO_MUESTRA_FECHA.CLIENTE_MES_NOMBRE Then
                Dim ARR_MES() As String = {"Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"}
                Return ARR_MES(CType(tx_mes, Integer) - 1)
            ElseIf _Mostrar = UTILITARIO_MUESTRA_FECHA.CLIENTE_12HORA Then
                Return tx_fecha_hora
            ElseIf _Mostrar = UTILITARIO_MUESTRA_FECHA.CLIENTE_24HORA Then
                Return tx_fecha_hora
            End If
        ElseIf XobjValue.ToString.Trim.Length = 8 Then
            Dim tx_dia As String = XobjValue.ToString.Substring(6, 2)
            Dim tx_mes As String = XobjValue.ToString.Substring(4, 2)
            Dim tx_anno As String = XobjValue.ToString.Substring(0, 4)
            Return tx_dia & "/" & tx_mes & "/" & tx_anno
        End If
        Return Nothing
    End Function

    Public Shared Function esNulo(ByVal XobjValue As Object) As Boolean
        Dim blnRes As Boolean = False
        If XobjValue Is System.DBNull.Value Then
            blnRes = True
        ElseIf XobjValue Is Nothing Then
            blnRes = True
        ElseIf XobjValue.ToString.Length = 0 Then
            blnRes = True
        End If
        Return blnRes
    End Function

    Public Shared Function CloneObject(ByVal obj As Object) As Object
        Try
            Dim ms As New System.IO.MemoryStream
            Dim bf As New BinaryFormatter(Nothing, New StreamingContext(StreamingContextStates.Clone))
            bf.Serialize(ms, obj)
            ms.Seek(0, SeekOrigin.Begin)
            Dim o As Object = bf.Deserialize(ms)
            ms.Close()
            Return o
        Catch ex As Exception
            Throw ex
        Finally
        End Try
    End Function


    Private Declare Auto Function SetProcessWorkingSetSize Lib "kernel32.dll" (ByVal procHandle As IntPtr, ByVal min As Int32, ByVal max As Int32) As Boolean
    Public Shared Sub ClearMemory()
        Try
            Dim Mem As Process
            Mem = Process.GetCurrentProcess()
            SetProcessWorkingSetSize(Mem.Handle, -1, -1)
        Catch ex As Exception
            'Control de errores
        End Try
    End Sub

    Public Shared Function Asignar_Propiedad_JSON(ByVal string_json As String, ByVal id_propiedad As String, ByVal id_valor As String) As String
        Dim il As IList = string_json.Split(Convert.ToChar(","))
        Dim I As Integer = 0
        For I = 0 To il.Count - 1
            If il(I).ToString.Split(Convert.ToChar(":")).GetValue(0).ToString.Equals(id_propiedad) Then
                il(I) = id_propiedad & ":" & id_valor
            End If
        Next

        Dim sb As New System.Text.StringBuilder("")
        For I = 0 To il.Count - 1
            sb.Append(il(I).ToString & ",")
        Next
        Return sb.ToString
    End Function
    Public Shared Function JSON(ByVal XstrCad As String) As String
        Dim sb As New System.Text.StringBuilder(XstrCad)
        Dim I As Integer = 0
        Dim n As Integer = 0
        Dim chars As String = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZÑñáéíóúÁÉÍÓÚ@-_."

        For I = 0 To sb.ToString.Length - 1
            'n = Convert.ToInt32(sb.Chars(I))
            If chars.IndexOf(sb.Chars(I)) < 0 Then
                sb.Chars(I) = Convert.ToChar(" ")
            End If
        Next
        Return sb.ToString
    End Function

    Public Shared Sub pInicializaVariables(ByVal obj As Object)
        Dim t As Type = obj.GetType
        Try
            Dim pr As PropertyInfo() = t.GetProperties
            For Each m As PropertyInfo In pr
                If DirectCast(m.PropertyType, System.Type).Name = "String" Then
                    If m.GetValue(obj, Nothing) Is Nothing Then
                        m.SetValue(obj, "", Nothing)
                    End If
                ElseIf DirectCast(m.PropertyType, System.Type).Name = "Int32" Then
                    If m.GetValue(obj, Nothing) Is Nothing Then
                        m.SetValue(obj, 0, Nothing)
                    End If
                End If
            Next
        Catch ex As Exception
            Dim o As Object = ex
        End Try
    End Sub

    Public Shared Sub pDestruyeVariables(ByVal obj As Object)
        Dim t As Type = obj.GetType
        Try
            Dim pr As PropertyInfo() = t.GetProperties
            For Each m As PropertyInfo In pr
                If DirectCast(m.PropertyType, System.Type).Name = "String" Or _
                   DirectCast(m.PropertyType, System.Type).Name = "Int32" Then
                    m.SetValue(obj, Nothing, Nothing)
                End If
            Next
        Catch ex As Exception
            Dim o As Object = ex
        End Try
    End Sub

    'Public Shared Function fDesSerializar(Of T)(ByVal LST As List(Of T), ByVal TX_DATA As String) As List(Of T)
    '    Dim ms As System.IO.MemoryStream = Nothing
    '    Dim serializer As DataContractJsonSerializer = Nothing
    '    Try
    '        ms = New System.IO.MemoryStream(Text.Encoding.Unicode.GetBytes(TX_DATA))
    '        serializer = New DataContractJsonSerializer(LST.GetType)
    '        Return CType(serializer.ReadObject(ms), List(Of T))
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        ms.Close()
    '        ms = Nothing
    '        serializer = Nothing
    '    End Try
    'End Function
    'Public Shared Function fDesSerializar(Of T)(ByVal LST As T, ByVal TX_DATA As String) As T
    '    Dim ms As System.IO.MemoryStream = Nothing
    '    Dim serializer As DataContractJsonSerializer = Nothing
    '    Try
    '        ms = New System.IO.MemoryStream(Text.Encoding.Unicode.GetBytes(TX_DATA))
    '        serializer = New DataContractJsonSerializer(LST.GetType)
    '        Return CType(serializer.ReadObject(ms), T)
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        ms.Close()
    '        ms = Nothing
    '        serializer = Nothing
    '    End Try
    'End Function

    'Public Shared Function fSerializar(Of T)(ByVal LST As T) As String
    '    Dim ms As New System.IO.MemoryStream
    '    Dim serializer As DataContractJsonSerializer
    '    Try
    '        serializer = New DataContractJsonSerializer(LST.GetType)
    '        serializer.WriteObject(ms, LST)
    '        Return Text.Encoding.UTF8.GetString(ms.ToArray)
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        ms.Close()
    '        ms = Nothing
    '    End Try
    'End Function

    'Public Shared Function fSerializar(Of T)(ByVal LST As List(Of T)) As String
    '    Dim ms As New System.IO.MemoryStream
    '    Dim serializer As DataContractJsonSerializer
    '    Try
    '        serializer = New DataContractJsonSerializer(LST.GetType)
    '        serializer.WriteObject(ms, LST)
    '        Return Text.Encoding.UTF8.GetString(ms.ToArray)
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        ms.Close()
    '        ms = Nothing
    '    End Try
    'End Function

    Public Shared Function fCambiaMayusculas(ByVal TX_ As Object) As String
        If Not Dame_Texto(TX_) Is Nothing Then
            Dim TX_BUSCA As String = CType(TX_, String).ToUpper.Trim
            TX_BUSCA = TX_BUSCA.Replace("Á", "A")
            TX_BUSCA = TX_BUSCA.Replace("É", "E")
            TX_BUSCA = TX_BUSCA.Replace("Í", "I")
            TX_BUSCA = TX_BUSCA.Replace("Ó", "O")
            TX_BUSCA = TX_BUSCA.Replace("Ú", "U")
            'ÀÈÌÒÙ
            TX_BUSCA = TX_BUSCA.Replace("À", "A")
            TX_BUSCA = TX_BUSCA.Replace("È", "E")
            TX_BUSCA = TX_BUSCA.Replace("Ì", "I")
            TX_BUSCA = TX_BUSCA.Replace("Ò", "O")
            TX_BUSCA = TX_BUSCA.Replace("Ù", "U")
            Return TX_BUSCA
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function Solo_Numeros(ByVal strNumber As String) As Boolean
        Solo_Numeros = True
        If strNumber = "" Then
            Solo_Numeros = False
        End If
        Dim i As Integer = 0
        i = 0
        Do While (i <= strNumber.Length - 1)
            If "1234567890".IndexOf(strNumber.Substring(i, 1)) < 0 Then
                'Error
                Solo_Numeros = False
            End If
            i = i + 1
        Loop

        Return Solo_Numeros
    End Function

    Public Shared Function OnlyGoodCharacters(ByVal texto As String, Optional ByVal Longitud As Integer = 0) As Boolean

        If texto.Length = 0 Then
            Return True
        End If

        If Longitud > 0 Then
            If texto.Length > Longitud Then
                Return False
            End If
        End If

        OnlyGoodCharacters = True
        
        Dim i As Integer = 0
        i = 0
        Do While (i <= texto.Length - 1)
            If "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_@-.+".IndexOf(texto.Substring(i, 1)) < 0 Then
                'Error
                OnlyGoodCharacters = False
                Exit Do
            End If
            i = i + 1
        Loop

        Return OnlyGoodCharacters
    End Function

    Public Shared Function ListToDataTable(Of T)(ByVal elements As System.Collections.Generic.IList(Of T)) As System.Data.DataTable
        Dim arrPropInfo As System.Reflection.PropertyInfo() = GetType(T).GetProperties()
        Dim dt As System.Data.DataTable = New DataTable()
        Dim dr As System.Data.DataRow

        For Each pInfo As System.Reflection.PropertyInfo In arrPropInfo
            If pInfo.PropertyType.Name.IndexOf("Nullable") < 0 Then
                dt.Columns.Add(pInfo.Name, pInfo.PropertyType)
            End If

        Next
        For Each elem As Object In elements
            dr = dt.NewRow()
            For Each pInfo As System.Reflection.PropertyInfo In arrPropInfo
                If pInfo.PropertyType.Name.IndexOf("Nullable") < 0 Then
                    dr(pInfo.Name) = pInfo.GetValue(elem, Nothing)
                End If

            Next
            dt.Rows.Add(dr)
        Next

        Return dt
    End Function
    Public Shared Function Valid_DNI(ByVal strDNI As String) As Boolean
        Valid_DNI = True
        Dim i As Integer = 0
        If strDNI.Length = 8 Then
            i = 0
            Do While (i <= strDNI.Length - 1)
                If "1234567890".IndexOf(strDNI.Substring(i, 1)) < 0 Then
                    'DNI Error
                    Valid_DNI = False
                End If
                i = i + 1
            Loop
        Else
            'DNI Error
            Valid_DNI = False
        End If
        Return Valid_DNI
    End Function

    Public Shared Function Extraer_Solo_Cifras(ByVal strDNI As String) As String
        Dim myChars() As Char = strDNI.ToCharArray()
        Dim ret As String = ""
        For Each ch As Char In myChars
            If Char.IsDigit(ch) Then
                ret = ret & ch.ToString
            End If
        Next
        Return ret
    End Function

    Public Shared Function Fecha_EnLetras(ByVal Fecha As Date) As String
        Fecha_EnLetras = Fecha.ToLongDateString()
        Fecha_EnLetras = Fecha_EnLetras.Replace("lunes, ", "")
        Fecha_EnLetras = Fecha_EnLetras.Replace("martes, ", "")
        Fecha_EnLetras = Fecha_EnLetras.Replace("miercoles, ", "")
        Fecha_EnLetras = Fecha_EnLetras.Replace("jueves, ", "")
        Fecha_EnLetras = Fecha_EnLetras.Replace("viernes, ", "")
        Fecha_EnLetras = Fecha_EnLetras.Replace("sábado, ", "")
        Fecha_EnLetras = Fecha_EnLetras.Replace("domingo, ", "")
    End Function

    'Public Shared Function Dame_Valor_WebConfig(ByVal param As String) As String

    '    Return System.Configuration.ConfigurationManager.AppSettings(param)
    'End Function

    Public Shared Function Es_Registrador_Provincia(ByVal Cod_Perfil As String) As Boolean
        If Cod_Perfil = "08" Or Cod_Perfil = "19" Then
            Return True
        Else
            Return False
        End If
    End Function

End Class
