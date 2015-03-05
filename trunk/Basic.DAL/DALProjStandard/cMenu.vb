Imports System.Data

Public Class cMenu

    Dim objConnection As Object

    Public MenuName As String
    Public MenuCaption As String
    Public MenuSeqNo As Double

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
    Public Sub DelMenu()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDatabaseManager.ExecuteNonQuery("Sp_DelMenu", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub

    Public Function SelMenu() As DataTable
        Dim dt As New DataTable

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        dt = objDatabaseManager.GetDataTable("GL_Menus_Get", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)
        Return dt
    End Function

    Public Sub SaveMenu()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@MenuName", MenuName, "nvarchar")
        objDBParameters.AddParameter("@MenuCaption", MenuCaption, "nvarchar")
        objDBParameters.AddParameter("@MenuSeqNo", MenuSeqNo, "nvarchar")

        objDatabaseManager.ExecuteNonQuery("GL_Menus_Save", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub
End Class
