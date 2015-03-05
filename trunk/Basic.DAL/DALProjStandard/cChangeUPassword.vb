Imports System.Data
Imports System.Data.SqlClient
Imports Basic.DAL
Imports Basic.Constants.ProjConst
Imports Basic.DAL.Utils
Public Class cChangeUPassword

    Dim objConnection As Object

    Public OldPassword As String
    Public NewPassword As String
    Public UserName As String

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
    Public Sub SaveCPValues()
        Try
            objConnection = cConnectionManager.GetConnection()
            Dim objDatabaseManager As IDatabaseManager
            Dim objDBParameters As New cDBParameterList
            objDatabaseManager = cDataBaseManager.GetDatabaseManager()
            objDatabaseManager.SetConnection(objConnection)

            objDBParameters.AddParameter("@OldPassword", OldPassword, "nvarchar")
            objDBParameters.AddParameter("@NewPassword", NewPassword, "nvarchar")
            objDBParameters.AddParameter("@UserName", UserName, "nvarchar")


            objDatabaseManager.ExecuteNonQuery("GL_ChangeUPassword_Save", objDBParameters)
            cConnectionManager.CloseConnection(objConnection)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
        End Try

    End Sub
    Public Sub SaveValues()
        Try
            objConnection = cConnectionManager.GetConnection()
            Dim objDatabaseManager As IDatabaseManager
            Dim objDBParameters As New cDBParameterList
            objDatabaseManager = cDataBaseManager.GetDatabaseManager()
            objDatabaseManager.SetConnection(objConnection)

            objDBParameters.AddParameter("@NewPassword", NewPassword, "nvarchar")
            objDBParameters.AddParameter("@UserName", UserName, "nvarchar")


            objDatabaseManager.ExecuteNonQuery("GL_ChangeUPassword_Save", objDBParameters)
            cConnectionManager.CloseConnection(objConnection)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
        End Try
    End Sub
End Class
