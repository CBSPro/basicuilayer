Imports System.Data
Imports System.Data.SqlClient
Imports Basic.DAL.Utils
Imports Basic.Constants.ProjConst
Imports Basic.DAL.Security
Imports Basic.DAL.cUsers
Imports Basic.DAL.cSysParms



Public Class cUsers

    Dim objConnection As Object

    Public MenuName As String
    Public MenuCaption As String
    Public MenuSeqNo As String
    Public UserID As String
    Public UserName As String
    Public Password As String
    Public SpUser As Boolean
    Public SeqNo As String
    Public ShowMenu As String
    Public CanAdd As String
    Public CanEdit As String
    Public CanDelete As String
    Public CanPost As String
    Public CanPrint As String
    Public CanPreview As String
    Public AddOn As String
    Public AddBy As String
    Public EditOn As String
    Public EditBy As String
    Public PostOn As String
    Public PostBy As String
    Public PrintOn As String
    Public PrintBy As String

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
    Public Function LoadMenu() As DataTable
        Dim dt As New DataTable

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        dt = objDatabaseManager.GetDataTable("GL_Menu_Get", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)
        Return dt
    End Function

    Public Function LoadSecurity() As DataTable
        Dim dt As New DataTable

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@UserID", UserID, "nvarchar")
        'objDBParameters.AddParameter("@MenuName", MenuName, "nvarchar")
        dt = objDatabaseManager.GetDataTable("GL_UserSecurity_Get", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)
        Return dt
    End Function
    Public Function LoadfrmSecurity() As DataTable
        Dim dt As New DataTable

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@UserID", UserID, "nvarchar")
        objDBParameters.AddParameter("@MenuName", MenuName, "nvarchar")

        dt = objDatabaseManager.GetDataTable("Sp_GetfrmSecurity", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)
        Return dt
    End Function
    Public Function LoadUsers() As DataTable
        Dim dt As New DataTable

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        dt = objDatabaseManager.GetDataTable("GL_AllUsers_Get", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)
        Return dt
    End Function

    Public Function GetSecurity() As DataTable
        Dim dt As New DataTable

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@UserID", UserID, "nvarchar")

        dt = objDatabaseManager.GetDataTable("GL_UserSecurity_Get", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)
        Return dt
    End Function

    Public Sub DelUser()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@UserID", UserID, "nvarchar")

        objDatabaseManager.ExecuteNonQuery("GL_User_Del", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub

    Public Sub DelDBUser()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@UserName", UserName, "nvarchar")

        objDatabaseManager.ExecuteNonQuery("GL_DBUser_Del", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub

    Public Sub SaveMaster()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@UserID", UserID, "nvarchar")
        objDBParameters.AddParameter("@UserName", UserName, "nvarchar")
        objDBParameters.AddParameter("@SpUser", SpUsers, "nvarchar")
        objDBParameters.AddParameter("@AddOn", AddOn, "nvarchar")
        objDBParameters.AddParameter("@AddBy", AddBy, "nvarchar")

        objDatabaseManager.ExecuteNonQuery("GL_MasterUser_Save", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub

    Public Sub SaveDetail()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@UserID", UserID, "nvarchar")
        objDBParameters.AddParameter("@MenuName", MenuName, "nvarchar")
        objDBParameters.AddParameter("@SeqNo", SeqNo, "nvarchar")
        objDBParameters.AddParameter("@ShowMenu", ShowMenu, "nvarchar")
        objDBParameters.AddParameter("@CanAdd", CanAdd, "nvarchar")
        objDBParameters.AddParameter("@CanEdit", CanEdit, "nvarchar")
        objDBParameters.AddParameter("@CanDelete", CanDelete, "nvarchar")
        objDBParameters.AddParameter("@CanPost", CanPost, "nvarchar")
        objDBParameters.AddParameter("@CanPrint", CanPrint, "nvarchar")
        objDBParameters.AddParameter("@CanPreview", CanPreview, "nvarchar")
        objDBParameters.AddParameter("@AddOn", AddOn, "nvarchar")
        objDBParameters.AddParameter("@AddBy", AddBy, "nvarchar")

        objDatabaseManager.ExecuteNonQuery("GL_DetailUser_Save", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub

    Public Sub SaveDBUser()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@UserName", UserName, "nvarchar")

        objDatabaseManager.ExecuteNonQuery("GL_DBUser_Save", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub

    Public Sub EditMaster()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@UserID", UserID, "nvarchar")
        objDBParameters.AddParameter("@UserName", UserName, "nvarchar")
        objDBParameters.AddParameter("@EditOn", EditOn, "nvarchar")
        objDBParameters.AddParameter("@EditBy", EditBy, "nvarchar")

        objDatabaseManager.ExecuteNonQuery("GL_MasterUser_Edit", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub

    Public Sub EditDetail()

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@UserID", UserID, "nvarchar")
        objDBParameters.AddParameter("@MenuName", MenuName, "nvarchar")
        objDBParameters.AddParameter("@SeqNo", SeqNo, "nvarchar")
        objDBParameters.AddParameter("@ShowMenu", ShowMenu, "nvarchar")
        objDBParameters.AddParameter("@CanAdd", CanAdd, "nvarchar")
        objDBParameters.AddParameter("@CanEdit", CanEdit, "nvarchar")
        objDBParameters.AddParameter("@CanDelete", CanDelete, "nvarchar")
        objDBParameters.AddParameter("@CanPost", CanPost, "nvarchar")
        objDBParameters.AddParameter("@CanPrint", CanPrint, "nvarchar")
        objDBParameters.AddParameter("@CanPreview", CanPreview, "nvarchar")
        objDBParameters.AddParameter("@EditOn", EditOn, "nvarchar")
        objDBParameters.AddParameter("@EditBy", EditBy, "nvarchar")

        objDatabaseManager.ExecuteNonQuery("GL_DetailUser_Edit", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)

    End Sub

    Public Function FindUser() As DataTable
        Dim dt As New DataTable

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@UserID", UserID, "nvarchar")

        dt = objDatabaseManager.GetDataTable("GL_FindUser_Get", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)
        Return dt
    End Function

    Public Function FindUserByName() As DataTable
        Dim dt As New DataTable

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@UserName", UserName, "nvarchar")

        dt = objDatabaseManager.GetDataTable("GL_FindUserByName_Get", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)
        Return dt
    End Function

    Public Function GetSrvDate() As DataTable
        Dim dt As New DataTable

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        dt = objDatabaseManager.GetDataTable("Sp_GetSrvDate", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)
        Return dt
    End Function

    Public Function GetSrvTime() As DataTable
        Dim dt As New DataTable

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        dt = objDatabaseManager.GetDataTable("GL_SrvTime_Get", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)
        Return dt
    End Function
End Class
