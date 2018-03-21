'Es necesario cargar el dll:
'c:\Archivos de Programas\Microsoft.NET\OracleClient.NET\System.Data.OracleClient.dll
Imports System.Data.SqlClient
Imports System
Imports System.Configuration
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Data.OracleClient
Imports System.Data.oledb
Imports System.Data.odbc
Imports System.Xml
Imports System.Text
Imports System.Security.Cryptography
Imports System.Reflection

'http://msdn.microsoft.com/msdnmag/issues/02/05/ASPSec2/default.aspx <--SEGURIDAD Has demo app, download at top
Public Class Utilitario
    Public retorno As Integer = 0
    Public varOutput As Object
    Public strRuta As String

    Public Function oraConexion() As OracleConnection
        'Dim strPassword = "desarrollo"
        Dim strConn As String
        strConn = ConfigurationSettings.GetConfig("appSettings")("ConnectionString")
        oraConexion = New OracleConnection(strConn)

        'oraConexion = New OracleConnection("Data Source=SROR;User id=orgpol;password=desarrollo")
        'oraConexion = New OracleConnection("Data Source=desaopint;User id=orgpolv2;password=desarrollo")
        Ruta = "/OrgPol/"
    End Function
    Public Function odbcConexion(ByVal strRuta As String) As OdbcConnection

        'Dim strConexion As String = ConfigurationSettings.GetConfig("appSettings")("DBFConnectionString")
        Dim strConexion As String = "Driver={Microsoft Visual FoxPro Driver};SourceType=DBF;SourceDB="

        odbcConexion = New OdbcConnection(strConexion + strRuta)

    End Function
    Public Function oleConexion(ByVal strRuta As String) As OleDbConnection
        oleConexion = New OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;data source=" & strRuta & ";Extended Properties=dbase IV;User ID=Admin;Password=")

    End Function
    Private Class Pair
        Public Entry As spEntry
        Public pmt As OracleClient.OracleParameter
        Sub New()
            MyBase.new()
        End Sub
    End Class
    Public Function CallSP(ByVal spName As String, ByVal cn As OracleConnection, Optional ByVal InputEntries() As spEntry = Nothing, Optional ByRef OutputEntries() As spEntry = Nothing, Optional ByVal InOutEntries() As spEntry = Nothing, Optional ByRef dtr As OracleClient.OracleDataReader = Nothing, Optional ByVal ReturnReader As Integer = 0, Optional ByRef Tran As OracleTransaction = Nothing, Optional ByRef xml As System.xml.XmlReader = Nothing) As Integer
        Dim cm As New OracleClient.OracleCommand(spName, cn)
        Dim Entry As spEntry
        Dim p As Integer
        Dim Pairs() As Pair
        Try
            cm.CommandText = spName
            cm.CommandType = CommandType.StoredProcedure

            If Not InputEntries Is Nothing Then
                For Each Entry In InputEntries
                    Dim pmt As OracleParameter = New OracleParameter(Entry.sVarName, Entry.vType, Entry.iSize)
                    pmt.Direction = ParameterDirection.Input
                    cm.Parameters.Add(pmt).Value = Entry.oValue

                Next
            End If

            If Not OutputEntries Is Nothing Then
                For Each Entry In OutputEntries
                    Dim pmt As OracleClient.OracleParameter
                    If Entry.vType = OracleType.Cursor Then
                        pmt = New OracleClient.OracleParameter(Entry.sVarName, Entry.vType)
                    Else
                        pmt = New OracleClient.OracleParameter(Entry.sVarName, Entry.vType, Entry.iSize)
                    End If

                    pmt.Direction = ParameterDirection.Output
                    cm.Parameters.Add(pmt)
                    If Pairs Is Nothing Then
                        p = 0
                    Else
                        p = UBound(Pairs) + 1
                    End If
                    ReDim Preserve Pairs(p)
                    Pairs(p) = New Pair
                    Pairs(p).Entry = Entry
                    Pairs(p).pmt = pmt

                Next
            End If
            If Not Tran Is Nothing Then
                cm.Transaction = Tran
            End If

            Select Case ReturnReader
                Case 0
                    cm.Parameters.Add("o_return", OracleType.Int16, 4).Direction = ParameterDirection.Output
                    cm.ExecuteNonQuery()
                    retorno = Convert.ToInt32(cm.Parameters("o_return").Value)
                Case 1
                    cm.Parameters.Add(New OracleClient.OracleParameter("R_CURSOR", OracleType.Cursor)).Direction = ParameterDirection.Output
                    dtr = cm.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
                Case 2
                    'xml = cm.ExecuteXmlReader
            End Select

            If Not OutputEntries Is Nothing Then
                Dim i As Integer
                For i = 0 To UBound(Pairs)
                    varOutput = Pairs(i).pmt.Value
                    OutputEntries(i).oValue = Pairs(i).pmt.Value
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function Decrypt(ByVal strText As String, ByVal sDecrKey As String) As String
        Dim byKey() As Byte = {}
        Dim IV() As Byte = {&H12, &H34, &H56, &H78, &H90, &HAB, &HCD, &HEF}
        Dim inputByteArray(strText.Length) As Byte

        Try
            byKey = System.Text.Encoding.UTF8.GetBytes(Left(sDecrKey, 8))
            Dim des As New DESCryptoServiceProvider
            inputByteArray = Convert.FromBase64String(strText)
            Dim ms As New MemoryStream
            Dim cs As New CryptoStream(ms, des.CreateDecryptor(byKey, IV), CryptoStreamMode.Write)

            cs.Write(inputByteArray, 0, inputByteArray.Length)
            cs.FlushFinalBlock()
            Dim encoding As System.Text.Encoding = System.Text.Encoding.UTF8

            Return encoding.GetString(ms.ToArray())

        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Sub CreaDataTable(ByVal strSQLorStoredProc As String, ByVal IsSP As Boolean, ByVal InputEntries() As spEntry, ByRef dt As DataTable)
        Dim _DT As New DataTable
        Dim Entry As spEntry
        Dim conOP As OracleConnection = oraConexion()
        conOP.Open()

        Dim cmd As OracleCommand = New OracleClient.OracleCommand(strSQLorStoredProc, conOP)

        If IsSP = True Then
            cmd.CommandText = strSQLorStoredProc
            cmd.CommandType = CommandType.StoredProcedure
            If Not InputEntries Is Nothing Then
                For Each Entry In InputEntries
                    cmd.Parameters.Add(Entry.sVarName, Entry.oValue)
                Next
            End If
            cmd.Parameters.Add(New OracleClient.OracleParameter("R_CURSOR", OracleType.Cursor)).Direction = ParameterDirection.Output
        Else
            cmd.CommandText = strSQLorStoredProc
            cmd.CommandType = CommandType.Text
        End If

        Dim daAdapter As New OracleDataAdapter
        daAdapter.SelectCommand = cmd
        dt = New DataTable
        daAdapter.Fill(_DT)
        dt = _DT
        conOP.Close()
    End Sub

    Public Function CreateDataTable(ByVal strSQLorStoredProc As String, ByVal IsSP As Boolean, ByVal InputEntries() As spEntry) As DataTable
        Dim _DT As New DataTable
        Dim Entry As spEntry
        Dim conOP As OracleConnection = oraConexion()
        conOP.Open()

        Dim cmd As OracleCommand = New OracleClient.OracleCommand(strSQLorStoredProc, conOP)

        If IsSP = True Then
            cmd.CommandText = strSQLorStoredProc
            cmd.CommandType = CommandType.StoredProcedure
            If Not InputEntries Is Nothing Then
                For Each Entry In InputEntries
                    cmd.Parameters.Add(Entry.sVarName, Entry.oValue)
                Next
            End If
            cmd.Parameters.Add(New OracleClient.OracleParameter("R_CURSOR", OracleType.Cursor)).Direction = ParameterDirection.Output
        Else
            cmd.CommandText = strSQLorStoredProc
            cmd.CommandType = CommandType.Text
        End If

        Dim daAdapter As New OracleDataAdapter
        daAdapter.SelectCommand = cmd

        daAdapter.Fill(_DT)

        conOP.Close()
        Return _DT
    End Function
    
    Public Function get_Apellidos_2(ByRef Nomb As String, ByRef ApPat As String, ByRef ApMat As String, ByVal Cadena As String)
        Dim i As Integer
        If Not Cadena = "" Then
            i = Cadena.IndexOf("*")
            If i > 0 Then
                Nomb = Cadena.Substring(0, i).Trim
                Cadena = Cadena.Substring(i + 1)
                i = Cadena.IndexOf("*")
                ApPat = Cadena.Substring(0, i).Trim
                Cadena = Cadena.Substring(i + 1).Trim
                ApMat = Cadena
            End If
        End If
    End Function

    Public Shared Function Dame_Texto(ByVal XobjValue As Object) As String
        
        If XobjValue Is System.DBNull.Value Then
            Return ""
        ElseIf XobjValue Is Nothing Then
            Return ""
        ElseIf XobjValue.ToString.Trim.Equals("") Then
            Return ""
        End If
        Return Convert.ToString(XobjValue)
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


    Public Function OnlyNumbers(ByVal str As String) As Boolean
        OnlyNumbers = True
        Dim i As Integer = 0

        i = 0
        Do While (i <= str.Length - 1)
            If "1234567890".IndexOf(str.Substring(i, 1)) < 0 Then
                ' Error
                OnlyNumbers = False
            End If
            i = i + 1
        Loop

        Return OnlyNumbers
    End Function


    Public Function SoloNumeros_Ascii(ByVal Keyascii As Short) As Short
        If InStr("1234567890", Chr(Keyascii)) = 0 Then
            SoloNumeros_Ascii = 0
        Else
            SoloNumeros_Ascii = Keyascii
        End If
        Select Case Keyascii
            Case 8
                SoloNumeros_Ascii = Keyascii
            Case 13
                SoloNumeros_Ascii = Keyascii
        End Select
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

    Public Property Ruta() As String
        Get
            Return strRuta
        End Get
        Set(ByVal Valor1 As String)
            strRuta = Valor1
        End Set
    End Property

End Class
Public Class spEntry
    Public sVarName As String
    Public oValue As Object
    Public vType As OracleType
    Public iSize As Integer
    Public Sub New(ByVal VarName As String, Optional ByVal Value As Object = Nothing, Optional ByVal oType As OracleType = OracleType.VarChar, Optional ByVal Size As Integer = -1)
        sVarName = VarName
        oValue = Value
        vType = oType
        iSize = Size
    End Sub
End Class

Public Class MyLista
    Inherits CollectionBase

    Sub Add(ByVal il As Object)
        Me.List.Add(il)
    End Sub

    Public Property Item(ByVal index As Integer) As Object
        Get
            Return Me.List(index)
        End Get
        Set(ByVal Value As Object)
            Me.List(index) = Value
        End Set
    End Property
    Sub Amayuscula_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim myDelegate As New MatchEvaluator(AddressOf MatchHandler)
        Dim sb As New System.Text.StringBuilder
        Dim bodyOfText As String = "el texto a convertir"

        Dim pattern As String = "\b(\w)(\w+)?\b"
        Dim re As New Regex(pattern, RegexOptions.Multiline Or RegexOptions.IgnoreCase)
        Dim newString As String = re.Replace(bodyOfText, myDelegate)
        ' newString es el string que contiene las primeras letras en mayuscula        
    End Sub
    Private Function MatchHandler(ByVal m As Match) As String
        Return m.Groups(1).Value.ToUpper() & m.Groups(2).Value
    End Function


    
    Private Function Raise_Confirm() As Integer

        Dim end1 As String = "</"
        Dim strScript As String = "<script language=""JavaScript"">" & vbCrLf & "<!-- " & vbCrLf & "MsgOkCancel();// --> " & end1 & "script>"

        Dim sb As New System.Text.StringBuilder("")
        sb.Append(strScript)

    
    End Function

End Class


