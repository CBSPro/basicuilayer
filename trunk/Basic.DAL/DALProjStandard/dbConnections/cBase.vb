Public Class cBase
    Protected pId As String
    'used for history trail
    Protected pHistoryTrailUserId As String
    Protected objConnection As Object

    Public Sub New(ByRef objCon As Object)
        objConnection = objCon
    End Sub

    Public Sub New()
    End Sub

    Public Sub SetConnection(ByRef objCon As Object)
        objConnection = objCon
    End Sub

    Public Property Id() As Integer
        Get
            Return pId
        End Get
        Set(ByVal Value As Integer)
            pId = Value
        End Set
    End Property
    Public Property HistoryTrailUserId() As String
        Get
            Return pHistoryTrailUserId
        End Get
        Set(ByVal Value As String)
            pHistoryTrailUserId = Value
        End Set
    End Property

End Class
