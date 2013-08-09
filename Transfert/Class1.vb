Module Class1

    Public Property pass As String
        Set(ByVal value As String)
            _pass = value
        End Set
        Get
            Return _pass
        End Get
    End Property
    Private _pass As String

    Public Property da As String
        Set(ByVal value As String)
            _da = value
        End Set
        Get
            Return _da
        End Get
    End Property
    Private _da As String

    Public Property a As String
        Set(ByVal value As String)
            _a = value
        End Set
        Get
            Return _a
        End Get
    End Property
    Private _a As String


    Public Sub TransferData(ByVal da As String, ByVal a As String)

        Try

        Catch e As Exception
            MsgBox(e.ToString)
        End Try



    End Sub


End Module