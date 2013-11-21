Imports System.Data.OleDb

Public Class Form1
    Dim oldb1 As New OleDb.OleDbConnection
    Dim oldb2 As New OleDb.OleDbConnection
    Dim mkcstr, mkcstrA As String

    ''' <summary>
    ''' Esci
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Dispose()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.ShowDialog()
        Class1.da = OpenFileDialog1.FileName
        TextBox1.Text = Class1.da
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        If Not (TextBox1.Text = "") And Not (TextBox2.Text = "") And Not (ComboBox1.Text = "") And Not (ComboBox2.Text = "") Then
            Try

                If Not IsNothing(Class1.da) Then
                    Select Case ComboBox1.SelectedIndex
                        Case 0 '"access 2007 to 2010"
                            Me.doneAcsCns(0)
                        Case 1 '"sql"

                            'todo other database
                            'case
                            'instruction
                    End Select
                Else
                    MsgBox("Specificare un percorso di provenienza.")
                    Exit Sub
                End If
                Select Case ComboBox2.SelectedIndex
                    Case 0 '"access 2007 to 2010"
                        Me.doneAcsCns(1)
                    Case 1 '"sql"

                End Select
                Todo()
            Catch ex As Exception
                MsgBox(e.ToString)
            End Try
        Else
            MsgBox("Devi compilare correttamente tutti i campi")
        End If
    End Sub

    ''' <summary>
    ''' Analizza, raccoglie le informazioni ed esegue la copia dei database.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Todo()


        Dim listaNomi, listaNomi2, listaNomi3 As New List(Of String)
        Dim dbCon As OleDb.OleDbConnection = Nothing
        Dim dbAdp As OleDb.OleDbDataAdapter
        Dim Data As New DataSet
        Dim SchemaTable As DataTable
        Dim Tables As New List(Of DataTable)
        Dim count As Long = 0

        ''Primo database
        Try

            makeCon(ComboBox1.SelectedIndex)
            dbCon = New OleDb.OleDbConnection(mkcstr)
            'Apre la connessione
            dbCon.Open()

            'Ottiene tutte le tabelle
            SchemaTable = dbCon.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, Nothing})
            count = SchemaTable.Rows.Count
            ProgressBar1.Value = Nothing : ProgressBar1.Step = 1 : ProgressBar1.Minimum = 0 : ProgressBar1.Maximum = count - 1
            'Le aggiunge alla collezione
            For I As Int16 = 0 To SchemaTable.Rows.Count - 1
                If SchemaTable.Rows(I).Item(3) = "TABLE" Then
                    Dim Name As String
                    Dim Table As DataTable
                    Name = SchemaTable.Rows(I).Item(2)
                    dbAdp = New OleDb.OleDbDataAdapter("SELECT * FROM `" & Name & "`", dbCon)
                    'E tramite questo riempie il dataset
                    dbAdp.Fill(Data)
                    Table = Data.Tables(0)
                    Table.TableName = Name
                    Tables.Add(Table)
                    Data.Dispose()
                    Data = New DataSet
                End If
                ProgressBar1.PerformStep() : ProgressBar1.Refresh()
            Next
            'Chiude la connessione
            dbCon.Close()

            'TextBox3.Text = count & " righe"
            For Each dataTbl As DataTable In Tables
                listaNomi.Add(dataTbl.TableName.ToString())
            Next
            'Rilascia tutto
            Data.Clear()
            Data.Dispose()

        Catch ex As Exception
            TextBox3.Text = count & " righe"
            If dbCon.State = ConnectionState.Open Then dbCon.Close()
            dbCon.Dispose() : If Not IsNothing(Data) Then Data.Dispose()
            MsgBox("Errore durante l'analisi e la raccolta informazioni del database di provenienza." & vbCrLf _
                   & ex.ToString)
            Exit Sub
        End Try


        ''secondo database
        Try

            makeCon(ComboBox2.SelectedIndex)
            dbCon = New OleDb.OleDbConnection(mkcstrA)
            'Apre la connessione
            dbCon.Open()

            'Ottiene tutte le tabelle
            SchemaTable = dbCon.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, Nothing})
            count = SchemaTable.Rows.Count
            ProgressBar2.Value = Nothing : ProgressBar2.Step = 1 : ProgressBar2.Minimum = 0 : ProgressBar2.Maximum = count - 1
            'Le aggiunge alla collezione
            For I As Int16 = 0 To SchemaTable.Rows.Count - 1
                If SchemaTable.Rows(I).Item(3) = "TABLE" Then
                    Dim Name As String
                    Dim Table As DataTable
                    Name = SchemaTable.Rows(I).Item(2)
                    dbAdp = New OleDb.OleDbDataAdapter("SELECT * FROM `" & Name & "`", dbCon)
                    'E tramite questo riempie il dataset
                    dbAdp.Fill(Data)
                    Table = Data.Tables(0)
                    Table.TableName = Name
                    Tables.Add(Table)
                    Data.Dispose()
                    Data = New DataSet
                End If
                ProgressBar2.PerformStep() : ProgressBar2.Refresh()
            Next
            'Chiude la connessione
            dbCon.Close()

            'TextBox3.Text = count & " righe"
            For Each dataTbl As DataTable In Tables
                listaNomi2.Add(dataTbl.TableName.ToString())
            Next
            'Rilascia tutto
            Data.Clear()
            Data.Dispose()

        Catch ex As Exception
            TextBox3.Text = count & " righe"
            If dbCon.State = ConnectionState.Open Then dbCon.Close()
            dbCon.Dispose() : If Not IsNothing(Data) Then Data.Dispose()
            MsgBox("Errore durante l'analisi e la raccolta informazioni del database di destinazione." & vbCrLf _
                   & ex.ToString)
            Exit Sub
        End Try

        ''Controlli sui dati raccolti
        For Each p In listaNomi ' coerenza delle tabelle

            For Each s In listaNomi2

                If s.ToString = p.ToString Then

                    listaNomi3.Add(s.ToString)

                End If

            Next

        Next

        import(listaNomi3)

    End Sub

    ''' <summary>
    ''' Importa da un database ad un'altro.
    ''' </summary>
    ''' <param name="val">
    ''' val = lista di tabelle da elaborare.
    ''' </param>
    ''' <remarks></remarks>
    Private Sub import(ByVal val As List(Of String))

        ''Importazione da un database all'altro
        Dim righe As Integer = 0
        Dim arrCounter As Integer = 0
        Dim righeNome(,) As String
        'If oldb1.State = 0 Then oldb1.Open()
        If oldb2.State = 0 Then oldb2.Open()


        Dim cmd As New OleDbCommand()
        cmd.Connection = oldb2
        cmd.CommandType = CommandType.Text
        ProgressBar3.Value = Nothing : ProgressBar3.Minimum = 0 : ProgressBar3.Maximum = val.Count : ProgressBar3.Step = 1

        For Each item In val
            On Error GoTo prossimo
            cmd.CommandText = Nothing
            cmd.CommandText = "INSERT INTO " & item.ToString & " SELECT * FROM " & item.ToString & " IN '" & My.Settings.Da.ToString & "'"
            righe = cmd.ExecuteNonQuery()
            ReDim Preserve righeNome(2, arrCounter) : righeNome = {{righe.ToString}, {item.ToString}}
            arrCounter += 1
prossimo:
            ProgressBar3.PerformStep() : ProgressBar3.Refresh()
        Next
        With TextBox3
            .AppendText(Messaggi.Sonostate & arrCounter)
            .AppendText(vbCrLf)
        End With
        For i = 0 To righeNome.Length / 2 - 1
            TextBox3.AppendText(Messaggi.sonostati.ToString & righeNome(0, i))
        Next
        oldb2.Dispose()
        oldb1.Close()

        'If righe > 0 Then MsgBox("Importazione avvenuta con successo." & vbCrLf & "Sono state elaborate " & righe & " righe, della tabella AnagraficaStudio")
    End Sub

    ''' <summary>
    ''' Inizializza con la connection string gli oggetti Oledb
    ''' </summary>
    ''' <param name="c">
    ''' 0 = Da
    ''' 1 = A
    ''' </param>
    ''' <remarks></remarks>
    Private Sub doneAcsCns(ByVal c As Single)
        If c = 0 Then

            If CheckBox1.CheckState = 0 Then
                oldb1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & My.Settings.Da.ToString & ";Persist Security Info=False;"
            Else
                LoginForm1.ShowDialog()
                oldb1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & My.Settings.Da.ToString & ";Jet OLEDB:Database Password=" & My.Settings.passDa.ToString & ";"
            End If
        ElseIf c = 1 Then
            If CheckBox2.CheckState = 0 Then
                oldb2.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & My.Settings.A.ToString & ";Persist Security Info=False;"
            Else
                LoginForm1.ShowDialog()
                oldb2.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & My.Settings.A.ToString & ";Jet OLEDB:Database Password=" & My.Settings.passA.ToString & ";"
            End If
        End If   '

    End Sub

    ''' <summary>
    ''' Costruisce e valorizza le variabili stringa di connessione
    ''' </summary>
    ''' <param name="tipoDatabase"></param>
    ''' <remarks></remarks>
    Private Sub makeCon(ByVal tipoDatabase As Byte)

        mkcstr = Nothing
        mkcstrA = Nothing
        Select Case tipoDatabase

            Case 0 'access 97-2003
                If My.Settings.credenzialiDa Then

                    mkcstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & My.Settings.Da.ToString & ";" & "Jet OLEDB:Database Password=" & My.Settings.passDa & ";"
                    mkcstrA = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & My.Settings.A.ToString & ";" & "Jet OLEDB:Database Password=" & My.Settings.passA & ";"
                Else
                    mkcstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & My.Settings.Da.ToString & ";User Id=admin; Password=;"
                    mkcstrA = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & My.Settings.A.ToString & ";User Id=admin; Password=;"
                End If

            Case 1 'access 07-2013
                mkcstr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & My.Settings.Da.ToString & ";Persist Security Info=False;"
                mkcstrA = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & My.Settings.A.ToString & ";Persist Security Info=False;"
            Case 2
                'TODO
        End Select


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        OpenFileDialog2.FileName = Nothing
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
        If LCase(ComboBox1.Text) = "access 2007 to 2010" Then
            CheckBox1.Enabled = True
        End If
    End Sub

    Private Sub ComboBox2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedValueChanged
        If LCase(ComboBox2.Text) = "access 2007 to 2010" Then
            CheckBox2.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' Salva le impostazioni.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        With My.Settings
            .Da = Trim(TextBox1.Text)
            .A = Trim(TextBox2.Text)
            .userDa = Trim(TextBox3.Text)
            .passDa = Trim(TextBox4.Text)
            .userA = Trim(TextBox5.Text)
            .passA = Trim(TextBox6.Text)
            .credenzialiDa = CheckBox1.CheckState
            .credenzialiA = CheckBox2.CheckState
            'aggiungere funzionalita di versione differenti
            .Save()
        End With
        MsgBox("Impostazioni salvate con successo.")
    End Sub

    Private Sub CheckBox1_CheckStateChanged(sender As Object, e As System.EventArgs) Handles CheckBox1.CheckStateChanged
        If CheckBox1.Checked = False Then
            TextBox3.Enabled = False
            TextBox4.Enabled = False
        Else
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox3.Text = My.Settings.userDa
            TextBox4.Text = My.Settings.passDa
        End If
    End Sub

    Private Sub CheckBox2_CheckStateChanged(sender As Object, e As System.EventArgs) Handles CheckBox2.CheckStateChanged
        If CheckBox2.Checked = False Then
            TextBox5.Enabled = False
            TextBox6.Enabled = False
        Else
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox5.Text = My.Settings.userA
            TextBox6.Text = My.Settings.passA
        End If
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        TextBox7.Text = Nothing
    End Sub
    Private Structure Messaggi
        Shared Sonostate = "Sono state elaborate "
        Shared tbl = "della tabella"
        Shared sonostati = "Sono stati elaborati "
        Shared rec = "record"
        Shared p = "."
        Shared br = " "
    End Structure
End Class
