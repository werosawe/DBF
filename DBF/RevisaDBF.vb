Imports System.Data.Odbc
Imports System.IO
Imports Utilitario
Imports System.Windows.Forms
Imports System.Configuration
Imports System.Text.RegularExpressions

Public Class RevisaDBF

    Private clase As New Utilitario.Utilitario
    Private NumReg As Integer
    Private strPathOnly As String = "C:\Desarrollo\VS2010\DBF\DBF\Output\"

    Private Sub btnRevisaDBF_Click(sender As System.Object, e As System.EventArgs) Handles btnRevisaDBF.Click

        Dim dbfFiles() As String = Get_DBF_files(strPathOnly)


        For Each file As String In dbfFiles

            file = Path.GetFileNameWithoutExtension(file).ToUpper & ".DBF"

            If file.ToUpper = "LIS_ADE.DBF" Then
                Verifica_DBF(file.ToUpper)
            End If

            If file.ToUpper = "ADHDATA.DBF" Then
                Verifica_DBF_ADHDATA(file.ToUpper)
            End If

            If file.ToUpper = "AFILIADO.DBF" Then
                Verifica_DBF_AFILIADO(file.ToUpper)
            End If

            If Es_Comite(file) = True Then
                Verifica_DBF(file)
            End If
        Next

        MsgBox("EL PROCESO HA CULMINADO", MsgBoxStyle.Information, "REVISION DE DBF")
    End Sub

    Private Sub Verifica_DBF(ByVal strFileNameOnly As String)

        Inicia_Timer()

        Dim MySQL As String = ""

        Dim dr As Odbc.OdbcDataReader

        Dim n As Num_Reg = Calcula_NumReg(strFileNameOnly)

        Dim conOP As OdbcConnection = clase.odbcConexion(strPathOnly)

        conOP.Open()
        MySQL = "SELECT num_pag, num_ite, num_ele, ape_pat, ape_mat, nom_ade FROM " + strFileNameOnly
        Dim cm As New Odbc.OdbcCommand(MySQL, conOP)
        dr = cm.ExecuteReader(System.Data.CommandBehavior.CloseConnection)

        Dim Num_Pag As String = ""
        Dim Num_Ite As String = ""
        Dim strDNI As String = ""
        Dim Ape_Pat As String = ""
        Dim Ape_Mat As String = ""
        Dim Nom_Ade As String = ""
        Dim s As String = ""

        If dr.HasRows = True Then

            txtOutput.AppendText("" & Environment.NewLine)
            txtOutput.AppendText("Archivo: " & strFileNameOnly & Environment.NewLine)
            txtOutput.AppendText("Cantidad de Registros (Delete Off): " & n.NumReg_Off & Environment.NewLine)
            txtOutput.AppendText("Cantidad de Registros (Delete On ): " & n.NumReg_On & Environment.NewLine)

            Do While dr.Read

                Num_Pag = dr.Item("num_pag").ToString().Trim
                Num_Ite = dr.Item("num_ite").ToString().Trim
                strDNI = dr.Item("num_ele").ToString().Trim
                Ape_Pat = dr.Item("Ape_pat").ToString().Trim.ToUpper
                Ape_Mat = dr.Item("Ape_mat").ToString().Trim.ToUpper
                Nom_Ade = dr.Item("nom_ade").ToString().Trim.ToUpper

                If Valida_Caracteres(Ape_Pat) = False Then

                    s = "Num_Pag: " & Num_Pag & " | Num_Ite: " & Num_Ite & " | DNI: " & strDNI & " | ApePat: " & Ape_Pat

                    txtOutput.AppendText(s & Environment.NewLine)

                End If

                If Valida_Caracteres(Ape_Mat) = False Then
                    s = "Num_Pag: " & Num_Pag & " | Num_Ite: " & Num_Ite & " | DNI: " & strDNI & " | ApeMat: " & Ape_Mat
                    txtOutput.AppendText(s & Environment.NewLine)
                End If

                If Valida_Caracteres(Nom_Ade) = False Then
                    s = "Num_Pag: " & Num_Pag & " | Num_Ite: " & Num_Ite & " | DNI: " & strDNI & " | Nombre: " & Nom_Ade
                    txtOutput.AppendText(s & Environment.NewLine)
                End If


                Me.NumReg = Me.NumReg - 1
                Application.DoEvents()


            Loop


        End If

        Finaliza_Timer()

    End Sub

    Private Sub Verifica_DBF_ADHDATA(ByVal strFileNameOnly As String)

        Inicia_Timer()

        Dim MySQL As String = ""

        Dim dr As Odbc.OdbcDataReader

        Dim n As Num_Reg = Calcula_NumReg(strFileNameOnly)

        Dim conOP As OdbcConnection = clase.odbcConexion(strPathOnly)

        conOP.Open()
        MySQL = "SELECT num_pagi, num_line, num_elec, ape_pate, ape_mate, nom_bres FROM " + strFileNameOnly
        Dim cm As New Odbc.OdbcCommand(MySQL, conOP)
        dr = cm.ExecuteReader(System.Data.CommandBehavior.CloseConnection)

        Dim Num_Pagi As String = ""
        Dim Num_Line As String = ""
        Dim strDNI As String = ""
        Dim Ape_Pate As String = ""
        Dim Ape_Mate As String = ""
        Dim Nom_Bres As String = ""
        Dim s As String = ""

        If dr.HasRows = True Then

            txtOutput.AppendText("" & Environment.NewLine)
            txtOutput.AppendText("Archivo: " & strFileNameOnly & Environment.NewLine)
            txtOutput.AppendText("Cantidad de Registros (Delete Off): " & n.NumReg_Off & Environment.NewLine)
            txtOutput.AppendText("Cantidad de Registros (Delete On ): " & n.NumReg_On & Environment.NewLine)

            Do While dr.Read

                Num_Pagi = dr.Item("num_pagi").ToString().Trim
                Num_Line = dr.Item("num_line").ToString().Trim
                strDNI = dr.Item("num_elec").ToString().Trim
                Ape_Pate = dr.Item("Ape_pate").ToString().Trim.ToUpper
                Ape_Mate = dr.Item("Ape_mate").ToString().Trim.ToUpper
                Nom_Bres = dr.Item("nom_bres").ToString().Trim.ToUpper

                If Valida_Caracteres(Ape_Pate) = False Then

                    s = "Num_Pagi: " & Num_Pagi & " | Num_Line: " & Num_Line & " | DNI: " & strDNI & " | ApePate: " & Ape_Pate

                    txtOutput.AppendText(s & Environment.NewLine)

                End If

                If Valida_Caracteres(Ape_Mate) = False Then
                    s = "Num_Pagi: " & Num_Pagi & " | Num_Line: " & Num_Line & " | DNI: " & strDNI & " | ApeMate: " & Ape_Mate
                    txtOutput.AppendText(s & Environment.NewLine)
                End If

                If Valida_Caracteres(Nom_Bres) = False Then
                    s = "Num_Pagi: " & Num_Pagi & " | Num_Line: " & Num_Line & " | DNI: " & strDNI & " | Nombre: " & Nom_Bres
                    txtOutput.AppendText(s & Environment.NewLine)
                End If


                Me.NumReg = Me.NumReg - 1
                Application.DoEvents()


            Loop


        End If

        Finaliza_Timer()

    End Sub


    Private Sub Verifica_DBF_AFILIADO(ByVal strFileNameOnly As String)

        Inicia_Timer()

        Dim MySQL As String = ""

        Dim dr As Odbc.OdbcDataReader

        Dim n As Num_Reg = Calcula_NumReg(strFileNameOnly)

        Dim conOP As OdbcConnection = clase.odbcConexion(strPathOnly)
        conOP.Open()
        ''MySQL = "SELECT num_fic, num_ele, ape_pat, ape_mat, nom_ade FROM " + strFileNameOnly
        MySQL = "SELECT num_pag, num_ele, ape_pat, ape_mat, nom_ade FROM " + strFileNameOnly
        Dim cm As New Odbc.OdbcCommand(MySQL, conOP)
        dr = cm.ExecuteReader(System.Data.CommandBehavior.CloseConnection)

        Dim Num_Pag As String = ""
        Dim Num_Ite As String = ""
        Dim strDNI As String = ""
        Dim Ape_Pat As String = ""
        Dim Ape_Mat As String = ""
        Dim Nom_Ade As String = ""
        Dim s As String = ""

        If dr.HasRows = True Then
            txtOutput.AppendText("" & Environment.NewLine)
            txtOutput.AppendText("Archivo: " & strFileNameOnly & Environment.NewLine)
            txtOutput.AppendText("Cantidad de Registros (Delete Off): " & n.NumReg_Off & Environment.NewLine)
            txtOutput.AppendText("Cantidad de Registros (Delete On ): " & n.NumReg_On & Environment.NewLine)

            Do While dr.Read

                ''Num_Pag = dr.Item("num_fic").ToString().Trim
                Num_Pag = dr.Item("num_pag").ToString().Trim
                strDNI = dr.Item("num_ele").ToString().Trim
                Ape_Pat = dr.Item("Ape_pat").ToString().Trim.ToUpper
                Ape_Mat = dr.Item("Ape_mat").ToString().Trim.ToUpper
                Nom_Ade = dr.Item("nom_ade").ToString().Trim.ToUpper

                If Valida_Caracteres(Ape_Pat) = False Then

                    s = "Num_Pag: " & Num_Pag & " | Num_Ite: " & Num_Ite & " | DNI: " & strDNI & " | ApePat: " & Ape_Pat

                    txtOutput.AppendText(s & Environment.NewLine)

                End If

                If Valida_Caracteres(Ape_Mat) = False Then
                    s = "Num_Pag: " & Num_Pag & " | Num_Ite: " & Num_Ite & " | DNI: " & strDNI & " | ApeMat: " & Ape_Mat
                    txtOutput.AppendText(s & Environment.NewLine)
                End If

                If Valida_Caracteres(Nom_Ade) = False Then
                    s = "Num_Pag: " & Num_Pag & " | Num_Ite: " & Num_Ite & " | DNI: " & strDNI & " | Nombre: " & Nom_Ade
                    txtOutput.AppendText(s & Environment.NewLine)
                End If


                Me.NumReg = Me.NumReg - 1
                Application.DoEvents()


            Loop


        End If

        Finaliza_Timer()

    End Sub

    Private Function Calcula_NumReg(ByVal strFileNameOnly As String) As Num_Reg
        Dim conOP As OdbcConnection = clase.odbcConexion(strPathOnly)
        Dim dr1 As Odbc.OdbcDataReader

        Dim n As New Num_Reg

        conOP.Open()
        Dim MySQL As String = "Set Deleted Off;SELECT COUNT(*) As TotalReg FROM " + strFileNameOnly

        Dim cm1 As New Odbc.OdbcCommand(MySQL, conOP)

        dr1 = cm1.ExecuteReader(System.Data.CommandBehavior.CloseConnection)

        If dr1.HasRows = True Then
            Do While dr1.Read
                n.NumReg_Off = dr1.Item("TotalReg").ToString()

            Loop
        End If
        conOP.Close()

        conOP.Open()
        MySQL = "Set Deleted On;SELECT COUNT(*) As TotalReg FROM " + strFileNameOnly

        cm1 = New Odbc.OdbcCommand(MySQL, conOP)

        dr1 = cm1.ExecuteReader(System.Data.CommandBehavior.CloseConnection)

        If dr1.HasRows = True Then
            Do While dr1.Read
                n.NumReg_On = dr1.Item("TotalReg").ToString()

            Loop
        End If
        conOP.Close()
        Return n
    End Function


    Function Valida_Caracteres(ByVal StringToCheck As String) As Boolean
        For i = 0 To StringToCheck.Length - 1
            If Not Char.IsLetter(StringToCheck.Chars(i)) Then
                If Not Char.IsWhiteSpace(StringToCheck.Chars(i)) Then
                    Return False
                End If
            End If

        Next

        Return True 'Return true if all elements are characters or blank space
    End Function

    Function Es_Comite(strFileNameOnly As String) As Boolean

        Dim f As String = Path.GetFileNameWithoutExtension(strFileNameOnly)

        If Regex.IsMatch(f, "^[0-9 ]+$") Then

            Return True
        Else
            Return False
        End If
    End Function

    Private Sub Inicia_Timer()
        Me.lblContador.Text = "... " + Me.NumReg.ToString("###,###") + " registros por procesar"
        lblContador.Visible = True
        Timer1.Interval = 1
        Timer1.Start()

    End Sub

    Private Sub Finaliza_Timer()

        lblContador.Visible = False
        Timer1.Stop()
    End Sub


    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Static alpha As Integer = 255
        Static delta As Integer = -20

        alpha += delta

        If alpha < 0 Then
            delta = 20
            alpha = 0
        ElseIf alpha > 255 Then
            delta = -20
            alpha = 255
        End If

        Me.lblContador.Text = "... " + Me.NumReg.ToString("###,###") + " registros por procesar"
        Me.lblContador.Refresh()
    End Sub


    Private Function Get_DBF_files(ByVal _path As String) As String()

        Dim ArrFiles As New List(Of String)

        Dim strFileSize As String = ""
        Dim di As New IO.DirectoryInfo(_path)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.dbf")
        Dim fi As IO.FileInfo
        For Each fi In aryFi

            'Console.WriteLine("File Full Name: {0}", fi.FullName)
            ArrFiles.Add(fi.FullName)

        Next


        Dim pdfFiles() As String = ArrFiles.ToArray()

        System.Array.Sort(Of String)(pdfFiles)

        Return pdfFiles
    End Function

End Class


Public Class Num_Reg

    Public Property NumReg_Off As Integer = 0
    Public Property NumReg_On As Integer = 0

End Class