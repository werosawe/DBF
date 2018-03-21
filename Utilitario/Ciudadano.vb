Imports System.Data.OracleClient
Public Class Ciudadano
    Private OP As New Utilitario
    Private mCod_DNI As String = ""
    Private mNombres As String = "-"
    Private mApePat As String = "-"
    Private mApeMat As String = "-"
    Private mCodSex As Integer = 0
    Private mUbigeoDNI As String = ""
    Private mEncontrado As Boolean = True

    Public Sub New(ByVal sCod_DNI As String)
        mCod_DNI = sCod_DNI
    End Sub
    Public Sub Get_Ciudadano_Nombres()
        Dim dr As OracleDataReader
        Dim ospEntry(0) As spEntry
        Dim conOP As OracleConnection = OP.oraConexion()
        conOP.Open()
        ospEntry(0) = New spEntry("i_cod_dni", Me.Cod_DNI, OracleType.Char, 8)
        OP.CallSP("pkg_USERS.SP_Listar_Apellidos", conOP, ospEntry, , , dr, 1)

        If dr.HasRows = True Then
            While (dr.Read())
                OP.get_Apellidos_2(mNombres, mApePat, mApeMat, dr.Item("nombres").ToString())
                If dr.Item("cod_sexo") Is DBNull.Value Then
                    mCodSex = 0
                Else
                    mCodSex = CType(dr.Item("cod_sexo"), Integer)
                End If
                mUbigeoDNI = dr.Item("ubigeo_dni")
            End While
        Else
            mNombres = "-"
            mApePat = "-"
            mApeMat = "-"
            mEncontrado = False
        End If
        conOP.Close()
        conOP.Dispose()
        dr.Close()
    End Sub
    Public ReadOnly Property Cod_DNI() As String
        Get
            Return mCod_DNI
        End Get
    End Property
    Public ReadOnly Property Nombres() As String
        Get
            Return mNombres
        End Get
    End Property
    Public ReadOnly Property ApePat() As String
        Get
            Return mApePat
        End Get
    End Property
    Public ReadOnly Property ApeMat() As String
        Get
            Return mApeMat
        End Get
    End Property
    Public ReadOnly Property CodSex() As Integer
        Get
            Return mCodSex
        End Get
    End Property
    Public ReadOnly Property UbigeoDNI() As String
        Get
            Return mUbigeoDNI
        End Get
    End Property
    Public ReadOnly Property NombreCompleto() As String
        Get
            Return mApePat + " " + mApeMat + " " + mNombres
        End Get
    End Property
    Public ReadOnly Property Encontrado() As Boolean
        Get
            Return mEncontrado
        End Get
    End Property

End Class
