Imports System.Messaging
Imports System.Threading

Public Class QueueObject

    'module level variables to maintain the name of the server and queue
    Private _strQueueServer As String = ""
    Private _strQueueName As String = ""

    'instantiating the object requires a server name and queue
    'these can be changed at during the objects lifetime via the QueueServer 
    'and QueueName properties
    Public Sub New(ByVal pstrQueueServer As String, _
           ByVal pstrQueueName As String)

        'trim server and queue name
        QueueServer = Trim(pstrQueueServer)
        QueueName = Trim(pstrQueueName)

        'validate the settings and let the sub throw the error
        _ValidateSettings(pstrQueueServer, pstrQueueName)

        'retain the values
        _strQueueServer = pstrQueueServer
        _strQueueName = pstrQueueName

    End Sub

    'Purpose: Aid in determining whether there are messages in the queue.
    'Note - client can also just read the queue and trap the queue exception that 
    'is thrown when reading an empty queue.
    Public ReadOnly Property CanRead() As Boolean
        Get
            
            Dim Queue As New MessageQueue()
            Queue = _GetQueue(_strQueueServer, _strQueueName)

            'peek to see if there is a message
            On Error Resume Next
            Dim QueueMessage As Message = Queue.Peek(New TimeSpan(0, 0, 1))
            If Err.Number = 0 Then
                'no error when peeking means we can read a record
                QueueMessage = Nothing
                Queue = Nothing
                Return True
            ElseIf Err.Number = 5 Then
                'error 5 when peeking means can't read
                QueueMessage = Nothing
                Queue = Nothing
                Return False
            Else
                'some other error 
                QueueMessage = Nothing
                Queue = Nothing
                On Error GoTo 0
                Throw New Exception(Err.Description)
                Return False
            End If
        End Get
    End Property

    'Purpose: Get or Set the queue name that the object will use.
    Public Property QueueName() As String
        Get
            Return _strQueueName
        End Get
        Set(ByVal Value As String)
            _strQueueName = Trim(Value)
        End Set
    End Property

    'Purpose: Get or Set the queue server that the object will use.
    Public Property QueueServer() As String
        Get
            Return _strQueueServer
        End Get
        Set(ByVal Value As String)
            _strQueueServer = Trim(Value)
        End Set
    End Property

    'Purpose:   Overloaded method to send a WSPMessage object to this object's queue
    'Input:     Message is type DOLQueueObjects.WSPMessage
    Public Overloads Sub SendMessage(ByVal Message As WA.DOL.VSD.WSPQueue.WSPMessage)

        SendMessage(Message.Auxiliary, Message.Mnemonic, Message.Body, Message.OriginatingID, Message.Delimiter)

    End Sub

    'Purpose:   Overloaded method to send the individual elements to this object's queue
    'Input:     Auxiliary = string for the ACCESS switch Auxiliary field
    '           Mnemonic = string containing the ACCESS switch Mnemonic field
    '           Body = string containing the ACCESS switch message body
    '           OriginatingID = string containing the ACCESS switch Originating ID field.
    '           When empty, the vsMSSGateway will insert the appropriate value.
    '           Delimiter = string indicating the ACCESS switch delimiter to use. This
    '           will almost always be "." but supports the ")" for future implementations.
    Public Overloads Sub SendMessage(ByVal pstrAuxiliary As String, ByVal pstrMnemonic As String, _
        ByVal pstrBody As String, Optional ByVal pstrOriginatingID As String = "", _
        Optional ByVal pstrDelimiter As String = ".")

        'verify current settings are still valid
        _ValidateSettings(_strQueueServer, _strQueueName)

        'validate input
        If pstrOriginatingID = "" Then
            pstrOriginatingID = "     "
        End If

        'Body must contain at least one character
        If pstrBody = "" Then
            pstrBody = " "
        End If

        'if delimiter is not a "." or ")", override it to "."
        If pstrDelimiter <> "." AndAlso pstrDelimiter <> ")" Then
            pstrDelimiter = "."
        End If

        'IMPORTANT NOTE:
        'The wrong number of delimiters can bring down the
        'NCIC connection, so strip them out to
        'insure no delimiters are in the fields.
        pstrOriginatingID = Replace(Replace(pstrOriginatingID, ".", " "), ")", " ")
        pstrAuxiliary = Replace(Replace(pstrAuxiliary, ".", " "), ")", " ")
        pstrMnemonic = Replace(Replace(pstrMnemonic, ".", " "), ")", " ")

        'create the recordset w/ values
        Dim rs As New ADODB.Recordset()

        rs.Fields.Append("MsgDate", ADODB.DataTypeEnum.adDate, 8, ADODB.FieldAttributeEnum.adFldIsNullable)
        rs.Fields.Append("QueueLabel", ADODB.DataTypeEnum.adVarChar, 8, ADODB.FieldAttributeEnum.adFldIsNullable)
        rs.Fields.Append("OrigID", ADODB.DataTypeEnum.adChar, 5, ADODB.FieldAttributeEnum.adFldIsNullable)
        rs.Fields.Append("Aux", ADODB.DataTypeEnum.adChar, 4, ADODB.FieldAttributeEnum.adFldIsNullable)
        rs.Fields.Append("Mnem", ADODB.DataTypeEnum.adVarChar, 255, ADODB.FieldAttributeEnum.adFldIsNullable)
        rs.Fields.Append("Delimiter", ADODB.DataTypeEnum.adChar, 1, ADODB.FieldAttributeEnum.adFldIsNullable)
        rs.Fields.Append("Body", ADODB.DataTypeEnum.adVarWChar, Len(pstrBody), ADODB.FieldAttributeEnum.adFldIsNullable)
        rs.Open()

        'add the values
        rs.AddNew()
        rs.Fields("QueueLabel").Value = _RemoveQualifier(_strQueueName)
        rs.Fields("OrigID").Value = Trim(pstrOriginatingID).PadRight(5).Substring(0, 5)
        rs.Fields("Aux").Value = Trim(pstrAuxiliary).PadRight(4).Substring(0, 4)
        rs.Fields("Mnem").Value = Trim(pstrMnemonic).PadRight(5).Substring(0, 5)
        rs.Fields("Delimiter").Value = pstrDelimiter
        rs.Fields("Body").Value = pstrBody
        rs.Fields("MsgDate").Value = Now
        rs.Update()

        'send the message to the queue
        'Dim Queue As MessageQueue = New MessageQueue(_strQueueServer & _strQueueName)
        Dim Queue As New MessageQueue()
        Queue = _GetQueue(_strQueueServer, _strQueueName)

        Queue.Formatter = New ActiveXMessageFormatter()
        Queue.Send(rs)


        'clean up
        Queue = Nothing
        rs = Nothing

    End Sub

    'Purpose:   Read a message from this object's queue.
    'Output:    DOLQueueObjects.WSPMessage object.
    Public Function ReadMessage() As WA.DOL.VSD.WSPQueue.WSPMessage


        Dim blnReturnKey As Boolean = False

        'verify current settings are still valid
        _ValidateSettings(_strQueueServer, _strQueueName)

        'Dim Queue As MessageQueue = New MessageQueue(_strQueueServer & _strQueueName)
        Dim Queue As New MessageQueue()
        Queue = _GetQueue(_strQueueServer, _strQueueName)

        Dim QueueMessage As Message
        Queue.Formatter = New ActiveXMessageFormatter()

        'fetch a message
        Try

            QueueMessage = Queue.Receive(New TimeSpan(0, 0, 1))
            blnReturnKey = True

            'message is a legacy ADODB Recordset
            Dim rs As New ADODB.Recordset()
            Dim Message As New WA.DOL.VSD.WSPQueue.WSPMessage()
            rs = QueueMessage.Body()

            'move recordset fields to WSPMessage elements
            Message.MessageDate = rs.Fields("MsgDate").Value
            Message.Auxiliary = rs.Fields("Aux").Value
            Message.QueueName = rs.Fields("QueueLabel").Value
            Message.Delimiter = rs.Fields("Delimiter").Value
            Message.Mnemonic = rs.Fields("Mnem").Value
            Message.OriginatingID = rs.Fields("OrigID").Value
            Message.Body = rs.Fields("Body").Value

            'clean up
            rs = Nothing
            QueueMessage = Nothing
            Queue = Nothing

            'return the ojbject
            Return Message

        Catch e As MessageQueueException
            'no message in the queue
            Throw New Exception(e.MessageQueueErrorCode)
        Catch e As InvalidOperationException
            'unable to read the queue message
            If blnReturnKey = True Then
                'if we read the message already, return it
                Queue.Send(QueueMessage)
            End If

            'bubble up the exception
            Throw New Exception(e.Message)

        End Try
    End Function

    'Purpose: Common code to open the queue
    Private Function _GetQueue(ByVal pstrQueueServer As String, _
           ByVal pstrQueueName As String) As System.Messaging.MessageQueue

        Dim Queues() As MessageQueue
        Dim Queue As New MessageQueue()
        Queues = Queue.GetPrivateQueuesByMachine(pstrQueueServer)
        Dim i As IEnumerator = Queues.GetEnumerator
        While i.MoveNext
            If LCase(CType(i.Current, MessageQueue).QueueName) = LCase(pstrQueueName) Then
                Return CType(i.Current, MessageQueue)
                Exit Function
            End If
        End While

    End Function

    'Purpose:   Removes the PRIVATE$\ qualifier when writing a message
    Private Function _RemoveQualifier(ByVal pstrQueueName As String) As String
        Return Replace(Replace(pstrQueueName, "PRIVATE$", "", , , CompareMethod.Text), "\", "")
    End Function

    Private Sub _ValidateSettings(ByVal pstrQueueServer As String, ByVal pstrQueueName As String)

        'verify we have a server, queuename, and queue
        If Trim(pstrQueueServer) = "" Then
            Throw New Exception("Queue Server not specified.")
        ElseIf Trim(pstrQueueName) = "" Then
            Throw New Exception("Queue Name not specified.")
        ElseIf _GetQueue(pstrQueueServer, pstrQueueName) Is Nothing Then
            Throw New Exception("Queue '" & pstrQueueName & "' does not exist on server '" & pstrQueueServer & "'")
        End If

    End Sub

End Class
