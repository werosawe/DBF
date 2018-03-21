Imports System.Data.oracleclient
'Imports OrgPol.oropwebservice

Public Class Candidato
    Private mAno_Elecc As Integer = 0
    Private mCod_Tipo_Elecc As String = ""
    Private mCod_OP As Integer = 0
    Private mUbiReg As Integer = 0
    Private mUbiProv As Integer = 0
    Private mUbiDist As Integer = 0

    Private mRegion As String = ""
    Private mProvincia As String = ""
    Private mDistrito As String = ""

    Public arrOP As New ArrayList
    Public arrOPElecc As ArrayList
    Private sCod_DNI As String
    Private dtDatosCandidato As DataTable
    Private clasesOP As New clasesOP
    Public Sub New(ByVal Cod_DNI As String)
        sCod_DNI = Cod_DNI
    End Sub
    Public Sub Cargar_OP()
        Dim oRowDatosCandidato As DataRow

        dtDatosCandidato = New DataTable
        dtDatosCandidato.Columns.Add(New DataColumn("codigoorop", GetType(Integer)))
        dtDatosCandidato.Columns.Add(New DataColumn("cargo", GetType(String)))
        dtDatosCandidato.Columns.Add(New DataColumn("posicion", GetType(String)))
        dtDatosCandidato.Columns.Add(New DataColumn("estado_candidato", GetType(String)))
        dtDatosCandidato.Columns.Add(New DataColumn("ambito", GetType(String)))
        dtDatosCandidato.Columns.Add(New DataColumn("region", GetType(String)))
        dtDatosCandidato.Columns.Add(New DataColumn("provincia", GetType(String)))
        dtDatosCandidato.Columns.Add(New DataColumn("distrito", GetType(String)))

        Dim ws As New ws_Candidato
        Dim ds As DataSet = ws.ds_CandidatoRegLoc(sCod_DNI)
        Dim dt As New DataTable
        dt = ds.Tables(0)

        For Each oRow As DataRow In dt.Rows
            oRowDatosCandidato = dtDatosCandidato.NewRow
            oRowDatosCandidato("codigoorop") = oRow("codigoorop")
            oRowDatosCandidato("cargo") = oRow("cargo")
            oRowDatosCandidato("posicion") = oRow("posicion")
            oRowDatosCandidato("estado_candidato") = oRow("estado_candidato")
            oRowDatosCandidato("ambito") = oRow("ambito")
            oRowDatosCandidato("region") = oRow("region")
            oRowDatosCandidato("provincia") = oRow("provincia")
            oRowDatosCandidato("distrito") = oRow("distrito")
            dtDatosCandidato.Rows.Add(oRowDatosCandidato)
        Next

        ds = ws.ds_CandidatoGen(sCod_DNI)
        dt = ds.Tables(0)

        For Each oRow As DataRow In dt.Rows
            oRowDatosCandidato = dtDatosCandidato.NewRow
            oRowDatosCandidato("codigoorop") = oRow("codigoorop")
            oRowDatosCandidato("cargo") = oRow("cargo")
            oRowDatosCandidato("posicion") = oRow("posicion")
            oRowDatosCandidato("estado_candidato") = oRow("estado_candidato")
            oRowDatosCandidato("ambito") = oRow("ambito")
            oRowDatosCandidato("region") = oRow("region")
            oRowDatosCandidato("provincia") = oRow("provincia")
            oRowDatosCandidato("distrito") = oRow("distrito")
            dtDatosCandidato.Rows.Add(oRowDatosCandidato)
        Next

        'dtDatosCandidato = ds.Tables(0)
        arrOP.Clear()

        For Each oRow As DataRow In dtDatosCandidato.Rows
            If indice_op_arrOP(oRow("codigoorop")) = -1 Then
                ' Si oRow("codigoorop") no existe en arrOP, agregarlo a arrOP
                ' solo se carga cod_op unicos
                arrOP.Add(New OPol(oRow("codigoorop")))
            End If
        Next

    End Sub
    Public Sub Cargar_OPElecc(ByVal iCod_OP As Integer)
        Dim oDatosElecc As DatosElecc
        Dim op_datos As OPol
        arrOPElecc = New ArrayList
        For Each oRow As DataRow In dtDatosCandidato.Rows
            If oRow("codigoorop") = iCod_OP Then
                oDatosElecc = New DatosElecc
                oDatosElecc.IsCandidato_Elecc = True
                oDatosElecc.Cargo_Elecc = oRow("cargo")
                oDatosElecc.Pos_Cargo = IIf(IsDBNull(oRow("posicion")), "-", oRow("posicion"))
                oDatosElecc.Estado_Candidato = oRow("estado_candidato")
                oDatosElecc.Tipo_Elecc = oRow("ambito")
                oDatosElecc.Anno_Elecc = 2006  ' ojo fijo

                Select Case oRow("ambito")
                    Case "PRESIDENCIAL", "PARLAMENTO ANDINO"
                        oDatosElecc.Ubigeo_Elecc = "-"
                    Case "CONGRESAL"
                        oDatosElecc.Ubigeo_Elecc = oRow("region")
                    Case "REGIONAL"
                        oDatosElecc.Ubigeo_Elecc = oRow("region")
                    Case "PROVINCIAL"
                        oDatosElecc.Ubigeo_Elecc = oRow("region") + " - " + oRow("provincia")
                    Case "DISTRITAL"
                        oDatosElecc.Ubigeo_Elecc = oRow("region") + " - " + oRow("provincia") + " - " + oRow("distrito")
                End Select

                oDatosElecc.Cod_OP = oRow("codigoorop")

                op_datos = New OPol(oDatosElecc.Cod_OP)
                oDatosElecc.Nombre_OP = op_datos.Des_OP
                oDatosElecc.Tipo_OP = op_datos.Des_Tipo_OP
                oDatosElecc.Ubigeo_OP = op_datos.Ubigeo_OP
                oDatosElecc.Estado_Insc = op_datos.Des_Estado_Insc

                arrOPElecc.Add(oDatosElecc)

            End If

        Next
    End Sub
    Public Function ListaCandidatos() As DataTable
        Dim dtListaCandidatos As New DataTable

        _Ubigeo(Me.UbiReg, Me.UbiProv, Me.UbiDist)

        Dim ws As New ws_Candidato
        Dim ds As DataSet = ws.ds_EleccRegLocOP(Me.iCod_OP, Me.sCod_Tipo_Elecc, Me.UbiReg, Me.UbiProv, Me.UbiDist)
        If ds Is Nothing Then
        Else
            dtListaCandidatos = ds.Tables(0)

        End If


        'dr.Item("posicion").ToString()
        'dr.Item("dni").ToString()
        'dr.Item("cargo").ToString()
        'dr.Item("nombre_completo").ToString()

        'dtListaCandidatos.Columns.Add(New DataColumn("Cod_DNI", GetType(String)))
        'dtListaCandidatos.Columns.Add(New DataColumn("ApePat", GetType(String)))
        'dtListaCandidatos.Columns.Add(New DataColumn("ApeMat", GetType(String)))
        'dtListaCandidatos.Columns.Add(New DataColumn("Nombre", GetType(String)))
        'dtListaCandidatos.Columns.Add(New DataColumn("Cargo", GetType(String)))
        'dtListaCandidatos.Columns.Add(New DataColumn("Posicion", GetType(String)))
        'Dim sCargo1 As String = ""
        'Dim sCargo2 As String = ""
        'Dim sCargo3 As String = ""
        'Select Case Me.sCod_Tipo_Elecc
        '    Case "01"  'Presidencial
        '        sCargo1 = "PRESIDENTE"
        '        sCargo2 = "VICE-PRESIDENTE"
        '        sCargo3 = "2do VICE-PRESIDENTE"
        '    Case "02"  'Congresista
        '        sCargo1 = "CONGRESISTA"
        '        sCargo2 = "CONGRESISTA"
        '        sCargo3 = "CONGRESISTA"
        '    Case "03"  'Regional
        '        sCargo1 = "PRESIDENTE REGIONAL"
        '        sCargo2 = "Otro Cargo"
        '        sCargo3 = "Otro Cargo"
        '    Case "04"  'Provincial
        '        sCargo1 = "ALCALDE"
        '        sCargo2 = "REGIDOR"
        '        sCargo3 = "REGIDOR"
        '    Case "05"  'Distrital
        '        sCargo1 = "ALCALDE"
        '        sCargo2 = "REGIDOR"
        '        sCargo3 = "REGIDOR"
        'End Select

        'Dim oRow As DataRow
        'oRow = dtListaCandidatos.NewRow
        'oRow("Cod_DNI") = "13584423"
        'oRow("ApePat") = "VALVERDE"
        'oRow("ApeMat") = "DELGADO"
        'oRow("Nombre") = "ALBERTO"
        'oRow("Cargo") = sCargo1
        'oRow("Posicion") = ""
        'dtListaCandidatos.Rows.Add(oRow)

        'oRow = dtListaCandidatos.NewRow
        'oRow("Cod_DNI") = "32984512"
        'oRow("ApePat") = "ASCENCIOS"
        'oRow("ApeMat") = "HEREDIA"
        'oRow("Nombre") = "RICARDO"
        'oRow("Cargo") = sCargo2
        'oRow("Posicion") = "1"
        'dtListaCandidatos.Rows.Add(oRow)

        'oRow = dtListaCandidatos.NewRow
        'oRow("Cod_DNI") = "40984513"
        'oRow("ApePat") = "GALVEZ"
        'oRow("ApeMat") = "SAN ROMAN"
        'oRow("Nombre") = "CARLOS"
        'oRow("Cargo") = sCargo3
        'oRow("Posicion") = "2"
        'dtListaCandidatos.Rows.Add(oRow)
        ListaCandidatos = dtListaCandidatos

    End Function
    Public Function ListaOP() As DataTable
        arrOP.Clear()
        Dim dt As New DataTable
        Dim dtListaOP As New DataTable
        dtListaOP.Columns.Add(New DataColumn("Cod_OP", GetType(Integer)))
        dtListaOP.Columns.Add(New DataColumn("Des_OP", GetType(String)))
        dtListaOP.Columns.Add(New DataColumn("Cod_Tipo_OP", GetType(String)))
        dtListaOP.Columns.Add(New DataColumn("Des_Tipo_OP", GetType(String)))

        _Ubigeo(Me.UbiReg, Me.UbiProv, Me.UbiDist)

        Dim ws As New ws_Candidato
        Dim ds As DataSet = ws.ds_EleccRegLocTipo(Me.sCod_Tipo_Elecc, Me.UbiReg, Me.UbiProv, Me.UbiDist)
        If ds Is Nothing Then
        Else
            dt = ds.Tables(0)
        End If

        Dim oOP As OPol
        Dim oRowListaOP As DataRow

        For Each oRow As DataRow In dt.Rows
            If indice_op_arrOP(oRow("codigoorop")) = -1 Then
                ' Si oRow("codigoorop") no existe en arrOP                
                oOP = New OPol(oRow("codigoorop"))
                arrOP.Add(New OPol(oRow("codigoorop")))
                oRowListaOP = dtListaOP.NewRow
                oRowListaOP("Cod_OP") = oOP.Cod_OP
                oRowListaOP("Des_OP") = oOP.Des_OP
                oRowListaOP("Cod_Tipo_OP") = oOP.Cod_Tipo_OP
                oRowListaOP("Des_Tipo_OP") = oOP.Des_Tipo_OP
                dtListaOP.Rows.Add(oRowListaOP)
            End If
        Next

        ListaOP = dtListaOP
    End Function

    Private Function indice_op_arrOP(ByVal iCod_OP As Integer) As Integer
        Dim nItemsOP As Integer = arrOP.Count
        If nItemsOP > 0 Then
            For i As Integer = 0 To nItemsOP - 1
                If arrOP(i).Cod_OP = iCod_OP Then
                    Return i
                End If
            Next
        End If
        Return -1
    End Function
    Private Function _Ubigeo(ByVal iUbiRegion As Integer, ByVal iUbiProv As Integer, ByVal iUbiDist As Integer)
        Dim Cad As String = "" ' Cadena de Ubigeo
        Dim ospEntry(2) As spEntry
        Dim conOP As OracleConnection = clasesOP.oraConexion()
        conOP.Open()
        ospEntry(0) = New spEntry("i_ubireg", iUbiRegion, OracleType.Int16, 2)
        ospEntry(1) = New spEntry("i_ubiprov", iUbiProv, OracleType.Int16, 2)
        ospEntry(2) = New spEntry("i_ubidist", iUbiDist, OracleType.Int16, 2)

        Dim ospOutPut(0) As spEntry
        ospOutPut(0) = New spEntry("o_mensaje", 0, OracleType.VarChar, 200)
        clasesOP.CallSP("PKG_UBIGEO.sp_get_ubigeo", conOP, ospEntry, ospOutPut, , , 0)
        conOP.Close()
        conOP.Dispose()

        Cad = CType(clasesOP.varOutput, String)
        Cad = Cad.Trim
        'Region:
        Dim x As Integer = InStr(1, Cad, "Dpto:") + Len("Dpto:")
        Dim y As Integer = InStr(1, Cad, ", Provincia:") - 1
        Me.Region = Cad.Substring(x, y - x)

        'Provincia:
        x = InStr(1, Cad, "Provincia:") + Len("Provincia:")
        y = InStr(1, Cad, ", Distrito:") - 1
        Me.Provincia = Cad.Substring(x, y - x)

        'Distrito:
        x = InStr(1, Cad, "Distrito:") + Len("Distrito:")
        y = Len(Cad)
        Me.Distrito = Cad.Substring(x, y - x)
    End Function
    Private Class DatosElecc
        Private mLeido As Boolean = False
        Private mIsCandidato_Elecc As Boolean = False
        Private mCargo_Elecc As String = "-"
        Private mPos_Cargo As String = "-"
        Private mEstado_Candidato As String = ""
        Private mTipo_Elecc As String = "-"
        Private mAnno_Elecc As Integer = 0
        Private mUbigeo_Elecc As String = "-"
        Private mCod_OP As Integer
        Private mNombre_OP As String = "-"
        Private mTipo_OP As String = "-"
        Private mUbigeo_OP As String = "-"
        Private mEstado_Insc As String = "-"

        Public Sub New()

        End Sub

        Public Property Leido() As Boolean
            Get
                Return mLeido
            End Get
            Set(ByVal Value As Boolean)
                mLeido = Value
            End Set
        End Property

        Public Property IsCandidato_Elecc() As Boolean
            Get
                Return mIsCandidato_Elecc
            End Get
            Set(ByVal Value As Boolean)
                mIsCandidato_Elecc = Value
            End Set
        End Property

        Public Property Cargo_Elecc() As String
            Get
                Return mCargo_Elecc
            End Get
            Set(ByVal Value As String)
                mCargo_Elecc = Value
            End Set
        End Property

        Public Property Pos_Cargo() As String
            Get
                Return mPos_Cargo
            End Get
            Set(ByVal Value As String)
                mPos_Cargo = Value
            End Set
        End Property

        Public Property Estado_Candidato() As String
            Get
                Return mEstado_Candidato
            End Get
            Set(ByVal Value As String)
                mEstado_Candidato = Value
            End Set
        End Property

        Public Property Tipo_Elecc() As String
            Get
                Return mTipo_Elecc
            End Get
            Set(ByVal Value As String)
                mTipo_Elecc = Value
            End Set
        End Property

        Public Property Anno_Elecc() As Integer
            Get
                Return mAnno_Elecc
            End Get
            Set(ByVal Value As Integer)
                mAnno_Elecc = Value
            End Set
        End Property

        Public Property Ubigeo_Elecc() As String
            Get
                Return mUbigeo_Elecc
            End Get
            Set(ByVal Value As String)
                mUbigeo_Elecc = Value
            End Set
        End Property
        Public Property Cod_OP() As Integer
            Get
                Return mCod_OP
            End Get
            Set(ByVal Value As Integer)
                mCod_OP = Value
            End Set
        End Property
        Public Property Nombre_OP() As String
            Get
                Return mNombre_OP
            End Get
            Set(ByVal Value As String)
                mNombre_OP = Value
            End Set
        End Property
        Public Property Tipo_OP() As String
            Get
                Return mTipo_OP
            End Get
            Set(ByVal Value As String)
                mTipo_OP = Value
            End Set
        End Property
        Public Property Ubigeo_OP() As String
            Get
                Return mUbigeo_OP
            End Get
            Set(ByVal Value As String)
                mUbigeo_OP = Value
            End Set
        End Property
        Public Property Estado_Insc() As String
            Get
                Return mEstado_Insc
            End Get
            Set(ByVal Value As String)
                mEstado_Insc = Value
            End Set
        End Property
    End Class

    Public Property iAno_Elecc() As Integer
        Get
            Return mAno_Elecc
        End Get
        Set(ByVal Value As Integer)
            mAno_Elecc = Value
        End Set
    End Property

    Public Property sCod_Tipo_Elecc() As String
        Get
            Return mCod_Tipo_Elecc
        End Get
        Set(ByVal Value As String)
            mCod_Tipo_Elecc = Value
        End Set
    End Property

    Public Property iCod_OP() As Integer
        Get
            Return mCod_OP
        End Get
        Set(ByVal Value As Integer)
            mCod_OP = Value
        End Set
    End Property

    Public Property UbiReg() As Integer
        Get
            Return mUbiReg
        End Get
        Set(ByVal Value As Integer)
            mUbiReg = Value
        End Set
    End Property

    Public Property UbiProv() As Integer
        Get
            Return mUbiProv
        End Get
        Set(ByVal Value As Integer)
            mUbiProv = Value
        End Set
    End Property

    Public Property UbiDist() As Integer
        Get
            Return mUbiDist
        End Get
        Set(ByVal Value As Integer)
            mUbiDist = Value
        End Set
    End Property

    Public Property Region() As String
        Get
            Return mRegion
        End Get
        Set(ByVal Value As String)
            mRegion = Value
        End Set
    End Property

    Public Property Provincia() As String
        Get
            Return mProvincia
        End Get
        Set(ByVal Value As String)
            mProvincia = Value
        End Set
    End Property

    Public Property Distrito() As String
        Get
            Return mDistrito
        End Get
        Set(ByVal Value As String)
            mDistrito = Value
        End Set
    End Property

End Class
