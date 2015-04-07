Imports System.IO

Public Class Form1
    Dim Parar As Boolean = False
    Dim Añadido As Integer = 0
    Dim Ignorado As Integer = 0
    Dim Total As Integer = 0
    Dim Inicio As Integer
    Dim Fin As Integer
    Dim UltimoItem As Integer = 0
    Dim UltimoObj As Integer = 0
    Dim Archivo_temp2 As String = "fr.dat" 'frances
    Dim Archivo_temp3 As String = "de.dat" 'aleman
    Dim Archivo_temp6 As String = "es.dat" 'español
    Dim Archivo_temp8 As String = "ru.dat" 'ruso
    Dim Obj_temp As String = "objetos.txt"
    Dim DatosItem As New StreamReader(Application.StartupPath & "\Datos\Item.dat")
    Dim DatosObj As New StreamReader(Application.StartupPath & "\Datos\Obj.dat")


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Contador As Integer = 0
        Dim I As Integer
        Button5.Enabled = False
        Inicio = Val(TextBox2.Text)
        Fin = Val(TextBox3.Text)
        Añadido = 0
        Ignorado = 0
        Total = 0
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        Parar = False
        Button2.Enabled = True
        ProgressBar1.Maximum = Fin + 1 'Fin - Inicio + 1
        ProgressBar1.Value = Inicio
        TextBox4.Text = ""
        TextBox5.Text = ""
        Button1.Enabled = False
        txtlang01.Text = ""
        txtlang02.Text = ""
        txtlang03.Text = ""
        txtlang04.Text = ""
        txtlang05.Text = ""
        txtlang06.Text = ""
        txtlang07.Text = ""
        txtlang08.Text = ""
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
        GroupBox3.Enabled = False
        Dim DatosItem As New StreamReader(Application.StartupPath & "\Datos\Item.dat")
        TextBox2.Text = DatosItem.ReadToEnd
        DatosItem.Close()

        For I = Inicio To Fin
            If Parar = True Then Exit For
            Label1.Text = I
            Me.Refresh()
            Application.DoEvents()
            'frances

            If My.Computer.FileSystem.FileExists(Archivo_temp2) Then My.Computer.FileSystem.DeleteFile(Archivo_temp2)
            'My.Computer.Network.DownloadFile("http://fr.wowhead.com/item=" & I, Archivo_temp2)
            If Descargar("http://fr.wowhead.com/item=" & I, Archivo_temp2) = True Then
                Dim datos As New StreamReader(Archivo_temp2)
                TextBox1.Text = datos.ReadToEnd
                datos.Close()
                Dim SearchWithinThis As String = TextBox1.Text
                Dim SearchForThis As String = "<h1 class="
                Dim SearchForThis2 As String = "</h1>"
                Dim FirstCharacter As Integer = SearchWithinThis.IndexOf(SearchForThis)
                Dim FirstCharacter2 As Integer = SearchWithinThis.IndexOf(SearchForThis2)
                If FirstCharacter > 0 Then
                    txtlang02.Text = Mid(SearchWithinThis, FirstCharacter + 28, FirstCharacter2 - FirstCharacter - 27)
                Else
                    txtlang02.Text = "NULL"
                End If
                txtlang02.Text = Replace(txtlang02.Text, "&#039;", "´")
                If Microsoft.VisualBasic.Left(txtlang02.Text, 1) <> "[" Then
                    'koreano
                    If chkLang01.CheckState = CheckState.Checked Then txtlang01.Text = ""
                    'Aleman
                    TextBox1.Text = ""
                    If chkLang03.CheckState = CheckState.Checked Then
                        If My.Computer.FileSystem.FileExists(Archivo_temp3) Then My.Computer.FileSystem.DeleteFile(Archivo_temp3)
                        My.Computer.Network.DownloadFile("http://de.wowhead.com/item=" & I, Archivo_temp3)
                        ' Call Descargar("http://fr.wowhead.com/item=" & I, Archivo_temp3)
                        Dim datos3 As New StreamReader(Archivo_temp3)
                        TextBox1.Text = datos3.ReadToEnd
                        datos3.Close()
                        SearchWithinThis = TextBox1.Text
                        SearchForThis = "<h1 class="
                        SearchForThis2 = "</h1>"
                        FirstCharacter = SearchWithinThis.IndexOf(SearchForThis)
                        FirstCharacter2 = SearchWithinThis.IndexOf(SearchForThis2)
                        If FirstCharacter > 0 Then
                            txtlang03.Text = Mid(SearchWithinThis, FirstCharacter + 28, FirstCharacter2 - FirstCharacter - 27)
                        Else
                            txtlang03.Text = "NULL"
                        End If
                        txtlang03.Text = Replace(txtlang03.Text, "&#039;", "´")
                    End If
                    'chino
                    If chkLang04.CheckState = CheckState.Checked Then txtlang04.Text = ""
                    'Twanes
                    If chkLang05.CheckState = CheckState.Checked Then txtlang05.Text = ""
                    'Español
                    TextBox1.Text = ""
                    If chkLang06.CheckState = CheckState.Checked Then

                        If My.Computer.FileSystem.FileExists(Archivo_temp6) Then My.Computer.FileSystem.DeleteFile(Archivo_temp6)
                        My.Computer.Network.DownloadFile("http://es.wowhead.com/item=" & I, Archivo_temp6)
                        'Call Descargar("http://fr.wowhead.com/item=" & I, Archivo_temp6)
                        Dim datos6 As New StreamReader(Archivo_temp6)
                        TextBox1.Text = datos6.ReadToEnd
                        datos6.Close()
                        SearchWithinThis = TextBox1.Text
                        SearchForThis = "<h1 class="
                        SearchForThis2 = "</h1>"
                        FirstCharacter = SearchWithinThis.IndexOf(SearchForThis)
                        FirstCharacter2 = SearchWithinThis.IndexOf(SearchForThis2)
                        If FirstCharacter > 0 Then
                            txtlang06.Text = Mid(SearchWithinThis, FirstCharacter + 28, FirstCharacter2 - FirstCharacter - 27)
                        Else
                            txtlang06.Text = "NULL"
                        End If
                        txtlang06.Text = Replace(txtlang06.Text, "&#039;", "´")
                    End If
                    'Español latino
                    If chkLang07.CheckState = CheckState.Checked Then txtlang07.Text = txtlang06.Text
                    'rusia
                    TextBox1.Text = ""
                    If chkLang08.CheckState = CheckState.Checked Then
                        If My.Computer.FileSystem.FileExists(Archivo_temp8) Then My.Computer.FileSystem.DeleteFile(Archivo_temp8)
                        My.Computer.Network.DownloadFile("http://ru.wowhead.com/item=" & I, Archivo_temp8)
                        'Call Descargar("http://fr.wowhead.com/item=" & I, Archivo_temp8)
                        Dim datos8 As New StreamReader(Archivo_temp8)
                        TextBox1.Text = datos8.ReadToEnd
                        datos8.Close()
                        SearchWithinThis = TextBox1.Text
                        SearchForThis = "<h1 class="
                        SearchForThis2 = "</h1>"
                        FirstCharacter = SearchWithinThis.IndexOf(SearchForThis)
                        FirstCharacter2 = SearchWithinThis.IndexOf(SearchForThis2)
                        If FirstCharacter > 0 Then
                            txtlang08.Text = Mid(SearchWithinThis, FirstCharacter + 28, FirstCharacter2 - FirstCharacter - 27)
                        Else
                            txtlang08.Text = "NULL"
                        End If
                        txtlang08.Text = Replace(txtlang08.Text, "&#039;", "´")
                    End If
                End If
                If txtlang02.Text <> txtlang05.Text Then
                    If Microsoft.VisualBasic.Left(txtlang02.Text, 1) <> "[" Then
                        If chkLang02.CheckState = CheckState.Checked Then
                            TextBox4.Text = TextBox4.Text & vbCrLf & "INSERT INTO `locales_item` VALUES (" & I & ",'" & txtlang01.Text & "','" & txtlang02.Text & "','" & txtlang03.Text & "','" & txtlang04.Text & "','" & txtlang05.Text & "','" & txtlang06.Text & "','" & txtlang07.Text & "','" & txtlang08.Text & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL);"
                        Else
                            TextBox4.Text = TextBox4.Text & vbCrLf & "INSERT INTO `locales_item` VALUES (" & I & ",'" & txtlang01.Text & "','','" & txtlang03.Text & "','" & txtlang04.Text & "','" & txtlang05.Text & "','" & txtlang06.Text & "','" & txtlang07.Text & "','" & txtlang08.Text & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL);"
                        End If
                        TextBox5.Text = "ID : " & I & " Name : " & txtlang06.Text
                        ListBox2.Items.Add(I)
                        Añadido = Añadido + 1
                        Label2.Text = "Añadido: " & Añadido
                        ProgressBar1.Value = ProgressBar1.Value + 1
                        ListBox2.SelectedItem = I
                        Me.Refresh()
                        Contador = Contador + 1
                        If Contador >= 2000 Then
                            Total = Contador + I
                            Contador = 0
                            UltimoItem = Total
                            TextBox5.Text = ""
                            TextBox5.Text = "--" & vbCrLf & "--      WOW Translator 2.0 generated the day " & Microsoft.VisualBasic.DateAndTime.DateValue(Now) & " - " & Microsoft.VisualBasic.DateAndTime.TimeValue(Now) & vbCrLf & "-- -----------------------------------------------------------------" & vbCrLf & vbCrLf & "--       Full database: " & Total & ", missing : " & Ignorado & ", added :" & Añadido & "      By Lanthi" & vbCrLf & "TRUNCATE locales_item;" & vbCrLf & TextBox4.Text
                            Call Guardar("Item")
                            TextBox4.Text = ""
                            Inicio = I
                        End If
                    Else
                        txtlang01.Text = ""
                        txtlang02.Text = ""
                        txtlang03.Text = ""
                        txtlang04.Text = ""
                        txtlang05.Text = ""
                        txtlang06.Text = ""
                        txtlang07.Text = ""
                        txtlang08.Text = ""
                        ListBox1.Items.Add(I)
                        Ignorado = Ignorado + 1
                        Label3.Text = "Omitidos : " & Ignorado
                        ProgressBar1.Value = ProgressBar1.Value + 1
                        ListBox1.SelectedItem = I
                        Me.Refresh()
                    End If
                End If
            Else
                txtlang01.Text = ""
                txtlang02.Text = ""
                txtlang03.Text = ""
                txtlang04.Text = ""
                txtlang05.Text = ""
                txtlang06.Text = ""
                txtlang07.Text = ""
                txtlang08.Text = ""
                ListBox1.Items.Add(I)
                Ignorado = Ignorado + 1
                Label3.Text = "Omitidos : " & Ignorado
                ProgressBar1.Value = ProgressBar1.Value + 1
                ListBox1.SelectedItem = I
                Me.Refresh()
            End If
        Next I
        Button1.Enabled = True
        Button2.Enabled = False
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        Label1.Text = 0
        txtlang01.Text = ""
        txtlang02.Text = ""
        txtlang03.Text = ""
        txtlang04.Text = ""
        txtlang05.Text = ""
        txtlang06.Text = ""
        txtlang07.Text = ""
        txtlang08.Text = ""
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ProgressBar1.Value = 0
        Label3.Text = "Omitidos : 0"
        Label2.Text = "Añadido: 0"
        Total = I
        UltimoItem = Total
        TextBox5.Text = " Añadidos : " & Añadido & " Omitidos : " & Ignorado & " Total : " & Total
        GroupBox3.Enabled = True
        If My.Computer.FileSystem.FileExists(Archivo_temp2) Then My.Computer.FileSystem.DeleteFile(Archivo_temp2)
        If My.Computer.FileSystem.FileExists(Archivo_temp3) Then My.Computer.FileSystem.DeleteFile(Archivo_temp3)
        If My.Computer.FileSystem.FileExists(Archivo_temp6) Then My.Computer.FileSystem.DeleteFile(Archivo_temp6)
        If My.Computer.FileSystem.FileExists(Archivo_temp8) Then My.Computer.FileSystem.DeleteFile(Archivo_temp8)
        TextBox5.Text = ""
        TextBox5.Text = "--" & vbCrLf & "--      WOW Translator 2.0 generated the day " & Microsoft.VisualBasic.DateAndTime.DateValue(Now) & " - " & Microsoft.VisualBasic.DateAndTime.TimeValue(Now) & vbCrLf & "-- -----------------------------------------------------------------" & vbCrLf & vbCrLf & "--       Full database: " & Total & ", missing : " & Ignorado & ", added :" & Añadido & "      By Lanthi" & vbCrLf & "TRUNCATE locales_item;" & vbCrLf & TextBox4.Text
        Call Guardar("Item")
    End Sub
    Function Descargar(ByVal Url As String, ByVal Path_Destino As String) As Boolean
        If Url = vbNullString Or Path_Destino = vbNullString Then
            MsgBox("No se indicó la url o el archivo de destino", MsgBoxStyle.Critical, "Error")
        Else
            If Len(Dir(Path_Destino)) <> 0 Then
                'MsgBox("el archivo ya existe.", MsgBoxStyle.Exclamation, "Error")
            Else
                On Error Resume Next
                My.Computer.Network.DownloadFile(Url, Path_Destino)
                If Err.Number = 0 Then
                    Descargar = True
                Else
                    ' MsgBox(Err.Description)
                    Descargar = False
                End If
                Err.Clear()
                Shell(Path_Destino)
            End If
        End If
    End Function

    Function Guardar(Tipo As String) As Boolean
        Dim Ruta As String
        Try
            Ruta = Tipo & " - " & Microsoft.VisualBasic.DateAndTime.Day(Now) & "-" & Microsoft.VisualBasic.DateAndTime.Month(Now) & "-" & Microsoft.VisualBasic.DateAndTime.Year(Now) & " (" & Inicio & " -- " & Total & ").sql"

            Dim escritor As StreamWriter
            Dim escritor1 As StreamWriter
            Dim escritor2 As StreamWriter
            escritor = File.AppendText(Ruta)

            Select Case Tipo
                Case "Item"
                    escritor.Write(TextBox5.Text)
                Case "Obj"
                    escritor.Write(TextBox8.Text)
                Case Else
            End Select
            escritor.Flush()
            escritor.Close()
            If My.Computer.FileSystem.FileExists(Application.StartupPath & "\Datos\Item.dat") Then My.Computer.FileSystem.DeleteFile(Application.StartupPath & "\Datos\Item.dat")
            escritor1 = File.AppendText(Application.StartupPath & "\Datos\Item.dat")
            escritor1.Write(UltimoItem)
            escritor1.Flush()
            escritor1.Close()
            If My.Computer.FileSystem.FileExists(Application.StartupPath & "\Datos\Obj.dat") Then My.Computer.FileSystem.DeleteFile(Application.StartupPath & "\Datos\Obj.dat")
            escritor2 = File.AppendText(Application.StartupPath & "Datos\Obj.dat")
            escritor2.Write(UltimoObj)
            escritor2.Flush()
            escritor2.Close()

            MessageBox.Show("Lote creado con éxito")
            If CheckBox1.CheckState = CheckState.Checked Then
                My.Computer.Network.UploadFile(Ruta, "ftp://lanthipiso.ddns.net/Wow%20Translator%202.0/Wow%20Translator%202.0/bin/Release/" & Ruta, "Lanthi", "3507042", True, 500)
                My.Computer.Network.UploadFile(Ruta, "ftp://lanthipiso.ddns.net/Wow%20Translator%202.0/Wow%20Translator%202.0/bin/Release/Datos/Item.dat", "Lanthi", "3507042", True, 500)
                My.Computer.Network.UploadFile(Ruta, "ftp://lanthipiso.ddns.net/Wow%20Translator%202.0/Wow%20Translator%202.0/bin/Release/Datos/Obj.dat", "Lanthi", "3507042", True, 500)

            End If
        Catch ex As Exception
            'MessageBox.Show("Escritura realizada incorrectamente")
        End Try
        Guardar = True
        Button5.Enabled = True
    End Function
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parar = True
        Button1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub Form1_Leave(sender As Object, e As EventArgs) Handles Me.Leave
        End
    End Sub

  
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If RadioButton1.Checked = True Then
            GroupBox1.Enabled = True
            GroupBox2.Enabled = False

        End If
        TextBox2.Text = DatosItem.ReadToEnd
        DatosItem.Close()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            GroupBox1.Enabled = True
            GroupBox2.Enabled = False
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            GroupBox2.Enabled = True
            GroupBox1.Enabled = False
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Button5.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = True
        Parar = False
        Dim datoshtml As New StreamReader(Obj_temp)
        TextBox6.Text = datoshtml.ReadToEnd
        datoshtml.Close()
        Dim Ignorados As Integer = 0
        Dim Añadidos As Integer = 0
        Dim ID As Integer
        Dim Name As String
        Dim TextoBuscar As String = TextBox6.Text
        Dim BuscarID As String = Chr(34) & "id" & Chr(34)
        Dim BuscarFin As String = ","
        Dim BuscarName As String = Chr(34) & "name" & Chr(34)
        Dim IDini As Integer
        Dim IDfin As Integer
        Dim NameIni As Integer
        Dim NameFin As Integer
        Dim Posicion As Integer = 0
        Dim Contador As Integer = 0
        Me.Refresh()
        Application.DoEvents()
        For I = Posicion To Len(TextBox6.Text)
            If Parar = True Then Exit For
            IDini = TextoBuscar.IndexOf(BuscarID, Posicion)
            If IDini <= 0 Then Exit For
            Posicion = IDini
            IDfin = TextoBuscar.IndexOf(BuscarFin, Posicion)
            ID = Val(Mid(TextoBuscar, IDini + 6, (IDfin - IDini) - 5))
            If ID > Total Then Total = ID
            NameIni = TextoBuscar.IndexOf(BuscarName, Posicion)
            Posicion = NameIni
            NameFin = TextoBuscar.IndexOf(BuscarFin, Posicion)
            Name = Mid(TextoBuscar, NameIni + 9, (NameFin - NameIni) - 9)

            If My.Computer.FileSystem.FileExists(Archivo_temp2) Then My.Computer.FileSystem.DeleteFile(Archivo_temp2)
            ' If Descargar("http://fr.wowhead.com/object=" & ID, Archivo_temp2) = True Then
            My.Computer.Network.DownloadFile("http://fr.wowhead.com/object=" & ID, Archivo_temp2)
            Dim datos As New StreamReader(Archivo_temp2)
            TextBox9.Text = datos.ReadToEnd
            datos.Close()
            Dim SearchWithinThis As String = TextBox9.Text
            Dim SearchForThis As String = "<h1 class="
            Dim SearchForThis2 As String = "</h1>"
            Dim FirstCharacter As Integer = SearchWithinThis.IndexOf(SearchForThis)
            Dim FirstCharacter2 As Integer = SearchWithinThis.IndexOf(SearchForThis2)
            If FirstCharacter > 0 Then
                txtobj02.Text = Mid(SearchWithinThis, FirstCharacter + 28, FirstCharacter2 - FirstCharacter - 27)
            Else
                txtobj02.Text = ""
            End If
            txtobj02.Text = Replace(txtobj02.Text, "&#039;", "´")
            txtobj02.Text = Replace(txtobj02.Text, "'", "´")

            If Microsoft.VisualBasic.Left(txtobj02.Text, 1) <> "[" Then
                'koreano
                If chkLang01.CheckState = CheckState.Checked Then txtobj01.Text = ""
                'Aleman
                If chkLang03.CheckState = CheckState.Checked Then
                    If My.Computer.FileSystem.FileExists(Archivo_temp3) Then My.Computer.FileSystem.DeleteFile(Archivo_temp3)
                    My.Computer.Network.DownloadFile("http://de.wowhead.com/object=" & ID, Archivo_temp3)
                    Dim datos3 As New StreamReader(Archivo_temp3)
                    TextBox9.Text = datos3.ReadToEnd
                    datos3.Close()
                    SearchWithinThis = TextBox9.Text
                    SearchForThis = "<h1 class="
                    SearchForThis2 = "</h1>"
                    FirstCharacter = SearchWithinThis.IndexOf(SearchForThis)
                    FirstCharacter2 = SearchWithinThis.IndexOf(SearchForThis2)
                    If FirstCharacter > 0 Then
                        txtobj03.Text = Mid(SearchWithinThis, FirstCharacter + 28, FirstCharacter2 - FirstCharacter - 27)
                    Else
                        txtobj03.Text = ""
                    End If

                    txtobj03.Text = Replace(txtobj03.Text, "&#039;", "´")
                    txtobj03.Text = Replace(txtobj03.Text, "'", "´")
                End If
                'chino
                If chkLang04.CheckState = CheckState.Checked Then txtobj04.Text = ""
                'Twanes
                If chkLang05.CheckState = CheckState.Checked Then txtobj05.Text = ""
                'Español
                If chkLang06.CheckState = CheckState.Checked Then
                    If My.Computer.FileSystem.FileExists(Archivo_temp6) Then My.Computer.FileSystem.DeleteFile(Archivo_temp6)
                    My.Computer.Network.DownloadFile("http://es.wowhead.com/object=" & ID, Archivo_temp6)
                    Dim datos6 As New StreamReader(Archivo_temp6)
                    TextBox9.Text = datos6.ReadToEnd
                    datos6.Close()
                    SearchWithinThis = TextBox9.Text
                    SearchForThis = "<h1 class="
                    SearchForThis2 = "</h1>"
                    FirstCharacter = SearchWithinThis.IndexOf(SearchForThis)
                    FirstCharacter2 = SearchWithinThis.IndexOf(SearchForThis2)
                    If FirstCharacter > 0 Then
                        txtobj06.Text = Mid(SearchWithinThis, FirstCharacter + 28, FirstCharacter2 - FirstCharacter - 27)
                    Else
                        txtobj06.Text = ""
                    End If
                    txtobj06.Text = Replace(txtobj06.Text, "&#039;", "´")
                    txtobj06.Text = Replace(txtobj06.Text, "'", "´")
                End If
                'Español latino
                If chkLang07.CheckState = CheckState.Checked Then txtobj07.Text = txtobj06.Text
                'rusia
                TextBox1.Text = ""
                If chkLang08.CheckState = CheckState.Checked Then
                    If My.Computer.FileSystem.FileExists(Archivo_temp8) Then My.Computer.FileSystem.DeleteFile(Archivo_temp8)
                    My.Computer.Network.DownloadFile("http://ru.wowhead.com/object=" & ID, Archivo_temp8)
                    Dim datos8 As New StreamReader(Archivo_temp8)
                    TextBox9.Text = datos8.ReadToEnd
                    datos8.Close()
                    SearchWithinThis = TextBox9.Text
                    SearchForThis = "<h1 class="
                    SearchForThis2 = "</h1>"
                    FirstCharacter = SearchWithinThis.IndexOf(SearchForThis)
                    FirstCharacter2 = SearchWithinThis.IndexOf(SearchForThis2)
                    If FirstCharacter > 0 Then
                        txtobj08.Text = Mid(SearchWithinThis, FirstCharacter + 28, FirstCharacter2 - FirstCharacter - 27)
                    Else
                        txtobj08.Text = ""
                    End If
                    txtobj08.Text = Replace(txtobj08.Text, "&#039;", "´")
                    txtobj08.Text = Replace(txtobj08.Text, "'", "´")
                End If
            End If
            If txtobj02.Text <> txtobj05.Text Then
                If Microsoft.VisualBasic.Left(txtobj06.Text, 1) <> "[" Then
                    TextBox7.Text = TextBox7.Text & vbCrLf & "INSERT INTO `locales_gameobject` VALUES (" & ID & ",'" & txtobj01.Text & "','" & txtobj02.Text & "','" & txtobj03.Text & "','" & txtobj04.Text & "','" & txtobj05.Text & "','" & txtobj06.Text & "','" & txtobj07.Text & "','" & txtobj08.Text & "','','','','','','','','','18019');"
                    Añadidos = Añadidos + 1
                    Label4.Text = "Añadidos = " & Añadidos
                    Label6.Text = "Total = " & Total
                Else
                    Ignorados = Ignorados + 1
                    Label6.Text = "Total = " & Total
                    Label5.Text = "Ignorados = " & Ignorados
                End If
                Me.Refresh()
                Application.DoEvents()
            End If
        Next
        Inicio = 1
        If My.Computer.FileSystem.FileExists(Obj_temp) Then My.Computer.FileSystem.DeleteFile(Obj_temp)
        TextBox6.Text = Mid(TextBox6.Text, Posicion)
        Try
            Dim escritor As StreamWriter
            escritor = File.AppendText(Obj_temp)
            escritor.Write(TextBox6.Text)
            escritor.Flush()
            escritor.Close()
            'MessageBox.Show("Lote creado con éxito")
        Catch ex As Exception
            'MessageBox.Show("Escritura realizada incorrectamente")
        End Try
        UltimoObj = Total
        TextBox8.Text = ""
        TextBox8.Text = "--" & vbCrLf & "--      WOW Translator 2.0 generated the day " & Microsoft.VisualBasic.DateAndTime.DateValue(Now) & " - " & Microsoft.VisualBasic.DateAndTime.TimeValue(Now) & vbCrLf & "-- -----------------------------------------------------------------" & vbCrLf & vbCrLf & "--       Full database: " & Total & ", missing : " & Ignorado & ", added :" & Añadido & "      By Lanthi" & vbCrLf & TextBox7.Text
        Call Guardar("Obj")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Parar = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = False


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        End
    End Sub
End Class
