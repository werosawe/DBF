Imports System.Data.OracleClient
Public Class PadronOP
    Private OP As New Utilitario
    Private mCod_OP As Integer = 0
    Private mDes_OP As String = ""
    Private mDes_Tipo_OP As String = ""
    Public Property UBigeo_OP As String = ""
    Private mUltima_Entrega As Integer = 0
    Private mNro_Entrega As Integer = 0
    Private mFec_Present As String = ""
    Private mYear_Padron_Afil As Integer = 0
    Private mPadronCancelatorio As Integer = 0
    Private mFechaCarga As String = ""
    Private mMensajeDeReporte As String = ""
    Private mdt As New DataTable

    Public Property Cod_Exp_MTD() As String
    Public Property Cod_Doc_MTD() As String

    Public Sub New(ByVal sCod_OP As Integer)
        mCod_OP = sCod_OP
    End Sub

    Public Sub Get_Datos_Padron()
        Dim dr As OracleDataReader
        Dim ospEntry(1) As spEntry
        Dim conOP As OracleConnection = OP.oraConexion()
        conOP.Open()
        ospEntry(0) = New spEntry("i_cod_op", Me.Cod_OP, OracleType.Int32, 4)
        ospEntry(1) = New spEntry("i_nro_entrega", Me.Nro_Entrega, OracleType.Int32, 4)
        OP.CallSP("pkg_Afiliados.SP_get_datos_padron", conOP, ospEntry, , , dr, 1)

        If dr.HasRows = True Then
            While (dr.Read())
                mDes_OP = dr.Item("des_op").ToString
                mDes_Tipo_OP = dr.Item("des_tipo_op").ToString
                If dr.Item("fec_present") Is DBNull.Value Then
                    mFec_Present = ""
                Else
                    mFec_Present = String.Format("{0:dd/MM/yyyy HH:mm:ss tt}", dr.Item("fec_present"))
                End If
                If dr.Item("fec_present") Is DBNull.Value Then
                    mYear_Padron_Afil = 0
                Else
                    mYear_Padron_Afil = CType(dr.Item("fec_present"), Date).Year
                End If

                mPadronCancelatorio = dr.Item("flg_cancelatorio").ToString

                If dr.Item("fecha_carga") Is DBNull.Value Then
                    mFechaCarga = ""
                Else
                    mFechaCarga = String.Format("{0:dd/MM/yyyy HH:mm:ss tt}", dr.Item("fecha_carga"))
                End If

                Me.UBigeo_OP = Utilitario.Dame_Texto(dr.Item("ubigeo_op"))

                Cod_Exp_MTD = dr.Item("Cod_exp_mtd").ToString
                Cod_Doc_MTD = dr.Item("Cod_doc_mtd").ToString

            End While
        Else
            mFec_Present = ""
            mYear_Padron_Afil = 0
            mPadronCancelatorio = 0
            mFechaCarga = ""
        End If
        conOP.Close()
        conOP.Dispose()
        dr.Close()
    End Sub

    Public Sub Get_Ult_Entrega()
        Dim dr As OracleDataReader
        Dim ospEntry(0) As spEntry
        Dim conOP As OracleConnection = OP.oraConexion()
        conOP.Open()
        ospEntry(0) = New spEntry("i_cod_op", Me.Cod_OP, OracleType.Int32, 4)
        OP.CallSP("pkg_Afiliados.SP_get_ult_entrega", conOP, ospEntry, , , dr, 1)

        If dr.HasRows = True Then
            While (dr.Read())
                mUltima_Entrega = CType(dr.Item("nro_entrega"), Integer)
                'mFec_Present = String.Format("{0:dd-MMM-yyyy}", dr.Item("fec_present"))
                mFec_Present = String.Format("{0:dd/MM/yyyy} HH:mi:ss tt", dr.Item("fec_present"))

            End While
        Else
            mUltima_Entrega = 0
            mFec_Present = ""
        End If
        conOP.Close()
        conOP.Dispose()
        dr.Close()
    End Sub

    Public Sub Variables_Informe_Carga(ByRef o_afil_pivot As Integer, _
                                       ByRef o_total_rec_dbf As Integer, _
                                       ByRef o_resultado_vf_N As Integer, _
                                       ByRef o_resultado_vf_D As Integer, _
                                       ByRef o_resultado_vf_A As Integer, _
                                       ByRef o_dni_invalidos As Integer, _
                                       ByRef o_dni_nopadronelec As Integer, _
                                       ByRef o_dni_nombresdist As Integer, _
                                       ByRef o_ubigeo_distinto As Integer, _
                                       ByRef o_dni_obs_rop As Integer, _
                                       ByRef o_dni_obs_rop_tot As Integer, _
                                       ByRef o_dni_aptos As Integer, _
                                       ByRef o_padron_cancelatorio As String, _
                                       ByRef o_desafil_entrega As Integer, _
                                       ByRef o_enotrasop As Integer, _
                                       ByRef o_afil_valid_entrega As Integer, _
                                       ByRef o_repres_valid As Integer, _
                                       ByRef o_dni_vuelven_informar As Integer, _
                                       ByRef o_afil_valid_total As Integer, _
                                       ByRef o_desafil_total As Integer)

        Dim conOP As OracleConnection = OP.oraConexion()
        Dim ospEntry(1) As spEntry
        ospEntry(0) = New spEntry("i_cod_op", Me.Cod_OP, OracleType.Int32, 4)
        ospEntry(1) = New spEntry("i_nro_entrega", Me.Nro_Entrega, OracleType.Int16, 2)

        Dim ospOutPut(19) As spEntry

        ospOutPut(0) = New spEntry("o_afil_pivot", 0, OracleType.Int32, 4)
        ospOutPut(1) = New spEntry("o_total_rec_dbf", 0, OracleType.Int32, 4)
        ospOutPut(2) = New spEntry("o_resultado_vf_N", 0, OracleType.Int32, 4)
        ospOutPut(3) = New spEntry("o_resultado_vf_D", 0, OracleType.Int32, 4)
        ospOutPut(4) = New spEntry("o_resultado_vf_A", 0, OracleType.Int32, 4)
        ospOutPut(5) = New spEntry("o_dni_invalidos", 0, OracleType.Int32, 4)
        ospOutPut(6) = New spEntry("o_dni_nopadronelec", 0, OracleType.Int32, 4)
        ospOutPut(7) = New spEntry("o_dni_nombresdist", 0, OracleType.Int32, 4)
        ospOutPut(8) = New spEntry("o_ubigeo_distinto", 0, OracleType.Int32, 4)
        ospOutPut(9) = New spEntry("o_dni_obs_rop", 0, OracleType.Int32, 4)
        ospOutPut(10) = New spEntry("o_dni_obs_rop_tot", 0, OracleType.Int32, 4)
        ospOutPut(11) = New spEntry("o_dni_aptos", 0, OracleType.Int32, 4)
        ospOutPut(12) = New spEntry("o_padron_cancelatorio", 0, OracleType.VarChar, 2)
        ospOutPut(13) = New spEntry("o_desafil_entrega", 0, OracleType.Int32, 4)
        ospOutPut(14) = New spEntry("o_enotrasop", 0, OracleType.Int32, 4)
        ospOutPut(15) = New spEntry("o_afil_valid_entrega", 0, OracleType.Int32, 4)
        ospOutPut(16) = New spEntry("o_repres_valid", 0, OracleType.Int32, 4)
        ospOutPut(17) = New spEntry("o_dni_vuelven_informar", 0, OracleType.Int32, 4)
        ospOutPut(18) = New spEntry("o_afil_valid_total", 0, OracleType.Int32, 4)
        ospOutPut(19) = New spEntry("o_desafil_total", 0, OracleType.Int32, 4)

        conOP.Open()
        OP.CallSP("pkg_Afiliados.sp_rpt_dni_informe_carga", conOP, ospEntry, ospOutPut, , , 0)
        conOP.Close()
        conOP.Dispose()

        o_afil_pivot = ospOutPut(0).oValue
        o_total_rec_dbf = ospOutPut(1).oValue
        o_resultado_vf_N = ospOutPut(2).oValue
        o_resultado_vf_D = ospOutPut(3).oValue
        o_resultado_vf_A = ospOutPut(4).oValue
        o_dni_invalidos = ospOutPut(5).oValue
        o_dni_nopadronelec = ospOutPut(6).oValue
        o_dni_nombresdist = ospOutPut(7).oValue
        o_ubigeo_distinto = ospOutPut(8).oValue
        o_dni_obs_rop = ospOutPut(9).oValue
        o_dni_obs_rop_tot = ospOutPut(10).oValue
        o_dni_aptos = ospOutPut(11).oValue
        o_padron_cancelatorio = ospOutPut(12).oValue
        o_desafil_entrega = ospOutPut(13).oValue
        o_enotrasop = ospOutPut(14).oValue
        o_afil_valid_entrega = ospOutPut(15).oValue
        o_repres_valid = ospOutPut(16).oValue
        o_dni_vuelven_informar = ospOutPut(17).oValue
        o_afil_valid_total = ospOutPut(18).oValue
        o_desafil_total = ospOutPut(19).oValue

    End Sub

    Public ReadOnly Property Cod_OP() As Integer
        Get
            Return mCod_OP
        End Get
    End Property

    Public ReadOnly Property Des_OP() As String
        Get
            Return mDes_OP
        End Get
    End Property

    Public ReadOnly Property Des_Tipo_OP() As String
        Get
            Return mDes_Tipo_OP
        End Get
    End Property

    Public ReadOnly Property Ultima_Entrega() As Integer
        Get
            Return mUltima_Entrega
        End Get
    End Property

    Public ReadOnly Property Fec_Present() As String
        Get
            Return mFec_Present
        End Get
    End Property

    Public ReadOnly Property Year_Padron_Afil() As Integer
        Get
            Return mYear_Padron_Afil
        End Get
    End Property

    Public Property Nro_Entrega() As Integer
        Get
            Return mNro_Entrega
        End Get
        Set(ByVal Value As Integer)
            mNro_Entrega = Value
        End Set
    End Property

    Public Property PadronCancelatorio() As Integer
        Get
            Return mPadronCancelatorio
        End Get
        Set(ByVal Value As Integer)
            mPadronCancelatorio = Value
        End Set
    End Property

    Public Property Fecha_Carga() As String
        Get
            Return mFechaCarga
        End Get
        Set(ByVal Value As String)
            mFechaCarga = Value
        End Set
    End Property

    Public Property MensajeDeReporte() As String
        Get
            Return mMensajeDeReporte
        End Get
        Set(ByVal Value As String)
            mMensajeDeReporte = Value
        End Set
    End Property

    Public Property dt() As DataTable
        Get
            Return mdt
        End Get
        Set(ByVal Value As DataTable)
            mdt = Value
        End Set
    End Property

End Class
