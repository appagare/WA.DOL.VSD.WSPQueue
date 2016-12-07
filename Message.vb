Public Class WSPMessage
    Private _QueueName As String = ""
    Private _OriginatingID As String = ""
    Private _MessageDate As Date = Now
    Private _Auxiliary As String = ""
    Private _Mnemonic As String = ""
    Private _Delimiter As String = "."
    Private _Body As String = ""

    Public Property Auxiliary() As String
        Get
            Return _Auxiliary
        End Get
        Set(ByVal Value As String)
            _Auxiliary = Value
        End Set
    End Property

    Public Property Body() As String
        Get
            Return _Body
        End Get
        Set(ByVal Value As String)
            _Body = Value
        End Set
    End Property

    Public Property Delimiter() As String
        Get
            Return _Delimiter
        End Get
        Set(ByVal Value As String)
            _Delimiter = Value
        End Set
    End Property

    Public Property MessageDate() As Date
        Get
            Return _MessageDate
        End Get
        Set(ByVal Value As Date)
            _MessageDate = Value
        End Set
    End Property

    Public Property Mnemonic() As String
        Get
            Return _Mnemonic
        End Get
        Set(ByVal Value As String)
            _Mnemonic = Value
        End Set
    End Property

    Public Property OriginatingID() As String
        Get
            Return _OriginatingID
        End Get
        Set(ByVal Value As String)
            _OriginatingID = Value
        End Set
    End Property

    Public Property QueueName() As String
        Get
            Return _QueueName
        End Get
        Set(ByVal Value As String)
            _QueueName = Value
        End Set
    End Property


End Class
