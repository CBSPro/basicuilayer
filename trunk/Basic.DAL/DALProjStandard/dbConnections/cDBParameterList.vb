
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections

Public Class cDBParameterList


    Dim objData As New Hashtable
    Dim count As Integer

    Public Sub New()
        count = 0
    End Sub

    Public Function AddParameter(ByVal Name As String, ByVal Value As String, ByVal Datatype As String)
        Dim objParameter As New cDBParameter

        objParameter.Name = Name
        objParameter.Value = Value
        objParameter.DataType = Datatype
        objData.Add(count, objParameter)
        count = count + 1
    End Function

    Public Function GetParameter(ByVal Name As String) As cDBParameter
        Dim i As Integer
        For i = 0 To count - 1
            If CType(objData.Item(i), cDBParameter).Name = Name Then
                Return CType(objData.Item(i), cDBParameter)
            End If
        Next
        Return Nothing
    End Function

    Public Function GetParameter(ByVal Index As Integer) As cDBParameter
        If Index < count Then
            Return CType(objData.Item(Index), cDBParameter)
        End If
        Return Nothing
    End Function

    Public ReadOnly Property ParameterCount() As Integer
        Get
            Return count
        End Get
    End Property

    ' Added By Farhana
    Public Function AddParameter(ByVal Parameter As cDBParameter)

        objData.Add(count, Parameter)
        count = count + 1
    End Function
    'Public Function AddParameter(ByVal Name As String, ByVal Value As String, ByVal Datatype As String, ByVal Key As Object)
    '    Dim objParameter As New cDBParameter

    '    objParameter.Name = Name
    '    objParameter.Value = Value
    '    objParameter.DataType = Datatype
    '    objData.Add(Key, objParameter)
    '    count = count + 1

    'End Function

    'Public Sub RemoveParameter(ByVal Key As Object)

    '    objData.Remove(Key)
    '    count = count - 1
    'End Sub


End Class
