Imports System.Data
Imports System.Data.SqlClient




Public Interface IDatabaseManager

    Sub SetConnection(ByVal objCon As Object)
    Sub SetTransaction(ByRef objTrans As Object)
    Sub BeginTransaction(Optional ByVal TransctionName As String = "")
    Sub CommitTransction()
    Sub RollBack()

    Function GetTransaction() As Object

    Function GetData(ByVal StoredProcedureName As String, ByRef pParameters As cDBParameterList) As DataSet
    Function GetDataTable(ByVal StoredProcedureName As String, ByRef pParameters As cDBParameterList, Optional ByVal TableName As String = "Table1") As DataTable
    Function GetDataFromQuery(ByVal SQLQuery As String) As DataSet
    Function GetDataFromQuery(ByVal SQLQuery As String, ByVal TableName As String) As DataTable

    Function ExecuteScalar(ByVal StoredProcedureName As String, ByRef pParameters As cDBParameterList) As Object
    Function ExecuteNonQuery(ByVal StoredProcedureName As String, ByRef pParameters As cDBParameterList) As Integer
 
End Interface
