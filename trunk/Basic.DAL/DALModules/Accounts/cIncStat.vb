
Imports System.Data
Imports System.Data.SqlClient

Public Class cIncStat

    Dim objConnection As Object

    Public PageNo As String
    Public mLineNo As String
    Public NoteRef As String
    Public Description As String
    Public Code As String
    Public CDesc As String
    Public RepType As String
    Public LBalance As Double
    Public CBalance As Double
    Public ABalance As Double
    Public LBudget As Double
    Public CBudget As Double
    Public ABudget As Double
    Public StPageNo As String
    Public EndPageNo As String
    Public BrField As Double
    Public TBal As Double
    Public mSQL As String

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
    Public Function SelDetail() As DataTable
        Dim dt As New DataTable

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        dt = objDatabaseManager.GetDataTable("GL_SelDetail", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

        Return dt
    End Function

    Public Sub DelDetail()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDatabaseManager.ExecuteNonQuery("GL_Detail_Del", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub

    Public Function SelPftLoss() As DataTable
        Dim dt As New DataTable

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@StPageNo", StPageNo, "nvarchar")
        objDBParameters.AddParameter("@EndPageNo", EndPageNo, "nvarchar")

        dt = objDatabaseManager.GetDataTable("GL_SelPftLoss_Get", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

        Return dt
    End Function

    Public Sub SaveDetail()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@PageNo", PageNo, "nvarchar")
        objDBParameters.AddParameter("@mLineNo", mLineNo, "nvarchar")
        objDBParameters.AddParameter("@NoteRef", NoteRef, "nvarchar")
        objDBParameters.AddParameter("@Description", Description, "nvarchar")
        objDBParameters.AddParameter("@LBalance", LBalance, "double")
        objDBParameters.AddParameter("@CBalance", CBalance, "double")
        objDBParameters.AddParameter("@ABalance", ABalance, "double")
        objDBParameters.AddParameter("@LBudget", LBudget, "double")
        objDBParameters.AddParameter("@CBudget", CBudget, "double")
        objDBParameters.AddParameter("@ABudget", ABudget, "double")

        objDatabaseManager.ExecuteNonQuery("GL_PageDetail_Save", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub

    Public Sub AppendDetail()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@PageNo", PageNo, "nvarchar")
        objDBParameters.AddParameter("@mLineNo", mLineNo, "nvarchar")
        objDBParameters.AddParameter("@NoteRef", NoteRef, "nvarchar")
        objDBParameters.AddParameter("@Description", Description, "nvarchar")
        objDBParameters.AddParameter("@Code", Code, "nvarchar")
        objDBParameters.AddParameter("@CDesc", CDesc, "nvarchar")
        objDBParameters.AddParameter("@LBalance", LBalance, "double")
        objDBParameters.AddParameter("@CBalance", CBalance, "double")
        objDBParameters.AddParameter("@ABalance", ABalance, "double")
        objDBParameters.AddParameter("@LBudget", LBudget, "double")
        objDBParameters.AddParameter("@CBudget", CBudget, "double")
        objDBParameters.AddParameter("@ABudget", ABudget, "double")

        objDatabaseManager.ExecuteNonQuery("GL_AppendDetail_Save", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub

    Public Sub AddDetail()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@PageNo", PageNo, "nvarchar")
        objDBParameters.AddParameter("@mLineNo", mLineNo, "nvarchar")
        objDBParameters.AddParameter("@NoteRef", NoteRef, "nvarchar")
        objDBParameters.AddParameter("@Description", Description, "nvarchar")
        objDBParameters.AddParameter("@Code", Code, "nvarchar")
        objDBParameters.AddParameter("@CDesc", CDesc, "nvarchar")
        objDBParameters.AddParameter("@BrField", ABudget, "double")

        objDatabaseManager.ExecuteNonQuery("GL_Detail_Add", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub

    Public Sub UpdatePftLoss()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@PageNo", PageNo, "nvarchar")
        objDBParameters.AddParameter("@mLineNo", mLineNo, "nvarchar")
        objDBParameters.AddParameter("@LBalance", LBalance, "double")
        objDBParameters.AddParameter("@CBalance", CBalance, "double")
        objDBParameters.AddParameter("@ABalance", ABalance, "double")
        objDBParameters.AddParameter("@LBudget", LBudget, "double")
        objDBParameters.AddParameter("@CBudget", CBudget, "double")
        objDBParameters.AddParameter("@ABudget", ABudget, "double")

        objDatabaseManager.ExecuteNonQuery("GL_PftLoss_Update", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub

    Public Sub AppendPftLoss()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@PageNo", PageNo, "nvarchar")
        objDBParameters.AddParameter("@mLineNo", mLineNo, "nvarchar")
        objDBParameters.AddParameter("@BrField", BrField, "double")

        objDatabaseManager.ExecuteNonQuery("GL_AppendPftLoss_Update", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub
End Class
