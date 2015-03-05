Imports System.Data

Public Class cSysParms

    Dim objConnection As Object

    Public FldName As String
    Public FldValue As String
    Public FldDescription As String

    Dim objTransaction As SqlClient.SqlTransaction
    Dim objDatabaseManager As IDatabaseManager
#Region " Transaction Specific Functions"

    Sub New()
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objConnection = cConnectionManager.GetConnection()
        objDatabaseManager.SetConnection(objConnection)
    End Sub

    Sub BeginTransaction(Optional ByVal TransactionName As String = "")
        If (TransactionName <> "") Then
            objDatabaseManager.BeginTransaction(TransactionName)
        Else
            objDatabaseManager.BeginTransaction()
        End If
    End Sub

    Sub CommitTransction()
        objDatabaseManager.CommitTransction()
    End Sub

    Sub RollBack()
        objDatabaseManager.RollBack()
    End Sub

    Public Sub getConnection()
        objConnection = cConnectionManager.GetConnection()
    End Sub

    Public Sub closeConnection()
        cConnectionManager.CloseConnection(objConnection)
    End Sub

#End Region
    Public Function LoadValues() As DataTable
        Dim dt As New DataTable

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        dt = objDatabaseManager.GetDataTable("GL_SysParms_Get", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)
        Return dt
    End Function

    Public Sub SaveValues()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@FldName", FldName, "nvarchar")
        objDBParameters.AddParameter("@FldValue", FldValue, "nvarchar")
        objDBParameters.AddParameter("@FldDescription", FldDescription, "nvarchar")

        objDatabaseManager.ExecuteNonQuery("GL_SysParms_Save", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub

    Public Sub EditSysParms()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@FldName", FldName, "nvarchar")
        objDBParameters.AddParameter("@FldValue", FldValue, "nvarchar")
        objDBParameters.AddParameter("@FldDescription", FldDescription, "nvarchar")

        objDatabaseManager.ExecuteNonQuery("GL_SysParms_Edit", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub

    Public Sub DelSysParms()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDatabaseManager.ExecuteNonQuery("GL_SysParms_Del", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub
End Class
