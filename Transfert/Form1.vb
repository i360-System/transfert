Public Class Form1
    Dim oldb1 As New OleDb.OleDbConnection
    Dim oldb2 As New OleDb.OleDbConnection


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Dispose()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
        Class1.da = OpenFileDialog1.FileName
        TextBox1.Text = Class1.da
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click


        Try

            If Not IsNothing(Class1.da) Then
                Select Case LCase(Class1.da.ToString)
                    Case "access 2007 to 2010"
                        Me.doneAcsCns(0)
                    Case "sql"
                        'todo other database
                        'case
                        'instruction
                End Select
            Else
                MsgBox("Specificare un percorso di provenienza.")
                Exit Sub
            End If
            Select Case LCase(a.ToString)
                Case "access 2007 to 2010"
                    Me.doneAcsCns(1)
            End Select
            Todo()
        Catch ex As Exception
            MsgBox(e.ToString)
        End Try
    End Sub

    Private Sub Todo()


        If oldb1.State Then oldb1.Open()

        Try
            Dim dp As New OleDb.OleDbDataAdapter("select * from anagraficastudio", oldb1)
            Dim dp2 As New OleDb.OleDbDataAdapter("select * from anagraficastudio", oldb2)
            Dim es As New DataSet1 : Dim tb As New DataTable : Dim es2 As New DataSet2 : Dim tb2 As New DataTable
            dp.Fill(es, "Anagraficastudio")
            tb = es.Tables("Anagraficastudio")
            ProgressBar1.Step = 1 : ProgressBar1.Value = 0
            ProgressBar1.Maximum = es.Tables("anagraficastudio").Rows.Count
            For Each r As DataRow In es.Tables("anagraficastudio").Rows
                es2.Tables("anagraficastudio").ImportRow(r)
                ProgressBar1.PerformStep() : ProgressBar1.Refresh()
            Next



        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub doneAcsCns(ByVal c As Single)
        If c = 0 Then

            If CheckBox1.CheckState Then
                oldb1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Class1.da.ToString & ";Persist Security Info=False;"
            Else
                LoginForm1.ShowDialog()
                oldb1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Class1.da.ToString & ";Jet OLEDB:Database Password=" & Class1.pass.ToString & ";"
            End If
        Else
            If CheckBox2.CheckState Then
                oldb2.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Class1.da.ToString & ";Persist Security Info=False;"
            Else
                LoginForm1.ShowDialog()
                oldb2.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Class1.da.ToString & ";Jet OLEDB:Database Password=" & Class1.pass.ToString & ";"
            End If
        End If   '

    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        OpenFileDialog2.ShowDialog()
        Class1.a = OpenFileDialog2.FileName
        TextBox2.Text = OpenFileDialog2.FileName
    End Sub

    Public Sub New()

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        Dim p As New List(Of String)
        Dim q As New List(Of String)

        p.Add("Access 2007 to 2010") : p.Add("Sql")
        q.Add("Access 2007 to 2010") : q.Add("Sql")

        For Each e In p
            ComboBox1.Items.Add(e)
        Next
        For Each r In q
            ComboBox2.Items.Add(r)
        Next



    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        If LCase(CheckBox1.Text) = "access 2007 to 2010" Then
            CheckBox1.Enabled = True
        End If
    End Sub

    Private Sub ComboBox2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedValueChanged
        If LCase(CheckBox2.Text) = "access 2007 to 2010" Then
            CheckBox2.Enabled = True
        End If
    End Sub
End Class
