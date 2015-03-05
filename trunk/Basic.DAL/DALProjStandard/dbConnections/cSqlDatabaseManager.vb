Imports System.Data
Imports System.Data.SqlClient

Public Class cSqlDatabaseManager
    Implements IDatabaseManager
    Private objConnection As SqlConnection
    Private objTransaction As SqlTransaction
   
    Public Sub SetConnection(ByVal objCon As Object) Implements IDatabaseManager.SetConnection
        Try
            objConnection = objCon
        Catch ex As Exception
            MsgBox(ex.Message)
            Throw New Exception("DB2")
        End Try
    End Sub

#Region " Select Specific Functions"

    Public Function GetData(ByVal StoredProcedureName As String, ByRef objParameters As cDBParameterList) As System.Data.DataSet Implements IDatabaseManager.GetData

        Dim objCommand As SqlCommand
        Dim objDataAdapter As New SqlDataAdapter
        Dim Query As String

        Query = StoredProcedureName

        objCommand = New SqlCommand(Query, objConnection)
        objCommand.CommandType = CommandType.StoredProcedure

        If (Not objTransaction Is Nothing) Then
            objCommand.Transaction = objTransaction
        End If

        Dim i As Integer
        If Not (objParameters Is Nothing) Then
            For i = 0 To objParameters.ParameterCount - 1
                objCommand.Parameters.Add(New SqlParameter(objParameters.GetParameter(i).Name, objParameters.GetParameter(i).DataType))
                If ("" & objParameters.GetParameter(i).Value).ToUpper = "DBNull.Value".ToUpper Then
                    objCommand.Parameters(objParameters.GetParameter(i).Name).Value = DBNull.Value
                Else
                    objCommand.Parameters(objParameters.GetParameter(i).Name).Value = "" & objParameters.GetParameter(i).Value
                End If
            Next
        End If

        objDataAdapter.SelectCommand = objCommand

        Dim dsTemp As New DataSet
        objDataAdapter.Fill(dsTemp, "Table1")

        Return dsTemp
    End Function

    Public Function GetDataFromQuery(ByVal SQLQuery As String) As System.Data.DataSet Implements IDatabaseManager.GetDataFromQuery

        Dim objCommand As SqlCommand
        Dim objDataAdapter As New SqlDataAdapter

        Dim Query As String
        Query = SQLQuery
        objCommand = New SqlCommand(Query, objConnection)
        objCommand.CommandType = CommandType.Text
        If Not (objTransaction Is Nothing) Then
            objCommand.Transaction = objTransaction
        End If
        objDataAdapter.SelectCommand = objCommand

        Dim dsTemp As New DataSet
        objDataAdapter.Fill(dsTemp, "Table1")

        Return dsTemp
    End Function
    Function GetDataFromQuery(ByVal SQLQuery As String, ByVal TableName As String) As DataTable Implements IDatabaseManager.GetDataFromQuery
        Dim objCommand As SqlCommand
        Dim objDataAdapter As New SqlDataAdapter

        Dim Query As String
        Query = SQLQuery
        objCommand = New SqlCommand(Query, objConnection)
        objCommand.CommandType = CommandType.Text
        If Not (objTransaction Is Nothing) Then
            objCommand.Transaction = objTransaction
        End If
        objDataAdapter.SelectCommand = objCommand

        Dim tempTable As New DataTable(TableName)
        objDataAdapter.Fill(tempTable)

        Return tempTable

    End Function

    Public Function GetDataTable(ByVal StoredProcedureName As String, ByRef objParameters As cDBParameterList, Optional ByVal TableName As String = "Table1") As DataTable Implements IDatabaseManager.GetDataTable
        Dim Table As New DataTable(TableName)
        Try
            Dim objCommand As SqlCommand
            Dim objDataAdapter As New SqlDataAdapter
            Dim Query As String

            Query = StoredProcedureName
            objCommand = New SqlCommand(Query, objConnection)
            objCommand.CommandType = CommandType.StoredProcedure

            If (Not objTransaction Is Nothing) Then
                objCommand.Transaction = objTransaction
            End If

            Dim i As Integer
            If Not (objParameters Is Nothing) Then
                For i = 0 To objParameters.ParameterCount - 1
                    objCommand.Parameters.Add(New SqlParameter(objParameters.GetParameter(i).Name, objParameters.GetParameter(i).DataType))
                    If ("" & objParameters.GetParameter(i).Value).ToUpper = "DBNull.Value".ToUpper Then
                        objCommand.Parameters(objParameters.GetParameter(i).Name).Value = DBNull.Value
                    Else
                        objCommand.Parameters(objParameters.GetParameter(i).Name).Value = "" & objParameters.GetParameter(i).Value
                    End If
                Next
            End If

            objDataAdapter.SelectCommand = objCommand
            objDataAdapter.Fill(Table)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return Table
    End Function

#End Region

#Region " Transaction Specific Functions"

    Sub BeginTransaction(Optional ByVal TransactionName As String = "") Implements IDatabaseManager.BeginTransaction
        If (TransactionName <> "") Then
            objTransaction = objConnection.BeginTransaction(TransactionName)

        Else
            objTransaction = objConnection.BeginTransaction()
        End If
    End Sub

    Sub SetTransaction(ByRef objTrans As Object) Implements IDatabaseManager.SetTransaction
        objTransaction = objTrans
    End Sub

    Public Function GetTransaction() As Object Implements IDatabaseManager.GetTransaction
        Return objTransaction
    End Function

    Sub CommitTransction() Implements IDatabaseManager.CommitTransction
        objTransaction.Commit()
    End Sub

    Sub RollBack() Implements IDatabaseManager.RollBack
        objTransaction.Rollback()
    End Sub

#End Region

#Region " Insert/Update/Delete Specific Functions"

    Public Function ExecuteScalar(ByVal StoredProcedureName As String, ByRef objParameters As cDBParameterList) As Object Implements IDatabaseManager.ExecuteScalar
        Dim objCommand As SqlCommand
        Dim objDataAdapter As New SqlDataAdapter
        Dim Query As String = StoredProcedureName

        objCommand = New SqlCommand(Query, objConnection)
        objCommand.CommandType = CommandType.StoredProcedure

        If Not (objTransaction Is Nothing) Then
            objCommand.Transaction = objTransaction
        End If

        Dim i As Integer
        If Not (objParameters Is Nothing) Then
            For i = 0 To objParameters.ParameterCount - 1
                objCommand.Parameters.Add(New SqlParameter(objParameters.GetParameter(i).Name, objParameters.GetParameter(i).DataType))
                If ("" & objParameters.GetParameter(i).Value).ToUpper = "DBNull.Value".ToUpper Then
                    objCommand.Parameters(objParameters.GetParameter(i).Name).Value = DBNull.Value
                Else
                    objCommand.Parameters(objParameters.GetParameter(i).Name).Value = "" & objParameters.GetParameter(i).Value
                End If
            Next
        End If

        Return objCommand.ExecuteScalar()
    End Function

    Public Function ExecuteNonQuery(ByVal StoredProcedureName As String, ByRef objParameters As cDBParameterList) As Integer Implements IDatabaseManager.ExecuteNonQuery
        Dim objCommand As SqlCommand
        Dim objDataAdapter As New SqlDataAdapter
        Dim Query As String = StoredProcedureName
        Try
            objCommand = New SqlCommand(Query, objConnection)
            objCommand.CommandType = CommandType.StoredProcedure

            If Not (objTransaction Is Nothing) Then
                objCommand.Transaction = objTransaction
            End If

            Dim i As Integer
            If Not (objParameters Is Nothing) Then
                For i = 0 To objParameters.ParameterCount - 1
                    Dim str As String
                    Dim val As String
                    str = objParameters.GetParameter(i).Name
                    val = Convert.ToString(objParameters.GetParameter(i).Value)
                    objCommand.Parameters.Add(New SqlParameter(objParameters.GetParameter(i).Name, objParameters.GetParameter(i).DataType))
                    If ("" & objParameters.GetParameter(i).Value).ToUpper = "DBNull.Value".ToUpper Then
                        objCommand.Parameters(objParameters.GetParameter(i).Name).Value = DBNull.Value
                    Else
                        objCommand.Parameters(objParameters.GetParameter(i).Name).Value = "" & objParameters.GetParameter(i).Value
                    End If
                Next
            End If

            'If Not (objTransaction Is Nothing) Then
            '    objCommand.Transaction = objTransaction
            'End If

            'Dim i As Integer
            'If Not (objParameters Is Nothing) Then
            '    For i = 0 To objParameters.ParameterCount - 1
            '        Dim str As String
            '        str = objParameters.GetParameter(i).Name
            '        objCommand.Parameters.Add(New SqlParameter(objParameters.GetParameter(i).Name, objParameters.GetParameter(i).DataType))
            '        If ("" & objParameters.GetParameter(i).Value).ToUpper = "DBNull.Value".ToUpper Then
            '            objCommand.Parameters(objParameters.GetParameter(i).Name).Value = DBNull.Value
            '        Else
            '            objCommand.Parameters(objParameters.GetParameter(i).Name).Value = "" & objParameters.GetParameter(i).Value
            '        End If
            '    Next
            'End If
            ' Return objCommand.ExecuteNonQuery()
            objCommand.ExecuteNonQuery()
        Catch ex As Exception
            'Throw New Exception("DB1")
            MsgBox(ex.Message)
        End Try
    End Function
#End Region

End Class
