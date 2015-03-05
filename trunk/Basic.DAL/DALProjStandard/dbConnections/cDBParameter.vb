Public Class cDBParameter
    Inherits cBase

    Protected pName As String
    Protected pValue As String
    Protected pDataType As String

    Public Sub New(ByRef objCon As Object)
        MyBase.New(objCon)
    End Sub

    Public Sub New()

    End Sub

    Public Property Name() As String
        Get
            Return pName
        End Get
        Set(ByVal Value As String)
            pName = Value
        End Set
    End Property

    Public Property Value() As String
        Get
            Return pValue
        End Get
        Set(ByVal Value As String)
            pValue = Value
        End Set
    End Property

    Public Property DataType() As String
        Get
            Return pDataType
        End Get
        Set(ByVal Value As String)
            pDataType = Value
        End Set
    End Property

End Class
