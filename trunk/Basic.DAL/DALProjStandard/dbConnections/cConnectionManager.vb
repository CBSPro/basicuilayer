Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Web
Imports Basic.DAL
Imports Basic.Constants.ProjConst
Imports Basic.DAL.Utils




Public Class cConnectionManager
    Private Shared objSqlConn As SqlConnection

    'changed to shared variables bcoz to be used by desktop application
    Public Shared DatabaseUser As String 'ConfigurationSettings.AppSettings.Get("user id") 
    Public Shared DatabasePassword As String 'ConfigurationSettings.AppSettings.Get("password") 
    Public Shared Database As String 'ConfigurationSettings.AppSettings.Get("initial catalog") 
    Public Shared DatabaseServer As String '"H-76B83038D5584" '"mk"  'ConfigurationSettings.AppSettings.Get("data Source") 
    Public Shared connstring As String

    Public Shared Function GetConnection() As Object
        'Dim DBType As String
        'Dim mode As String
        Try
            objSqlConn = New SqlConnection

            'creating connection string based on values from session
            Dim cString As String
            DatabaseServer = SysServer
            'DatabaseServer = "CBS-Server"
            Database = SysDataBase
            DatabaseUser = SysUser
            DatabasePassword = SysPassword
            'cString = "data source=" & DatabaseServer & ";initial catalog=" & Database & ";user id=" & DatabaseUser & "; password=" & DatabasePassword
            cString = "Server=" & DatabaseServer & ";Database=" & Database & ";user id=" & DatabaseUser & ";password=" & DatabasePassword
            'cString = connstring
            'passing created connection string 
            objSqlConn.ConnectionString = cString

            If objSqlConn.State = ConnectionState.Closed Then
                objSqlConn.Open()


                'frmMdi.mnuTransactions.Enabled = True
                'frmMdi.mnuBasicData.Enabled = True
                'frmMdi.mnuReports.Enabled = True
                'frmMdi.mnuAdmin.Enabled = True

            End If
            Return objSqlConn
        Catch ex As Exception
            'Throw New Exception("DB1")
            'MsgBox(ex.Message)
        End Try
    End Function

    Public Shared Function CloseConnection(ByVal objCon As Object)
        Try
            'Dim DBType As String
            'DBType = 'System.Configuration.ConfigurationSettings.AppSettings("DatabaseType").ToString.ToUpper
            CType(objCon, SqlConnection).Close()
            Return 0
        Catch ex As Exception
            'Throw New Exception("DB1")
            MsgBox(ex.Message)
        End Try
    End Function

    Public Shared Function GetSQLConnection() As SqlConnection
        Return objSqlConn
    End Function

    Public Shared ReadOnly Property ConnectionString() As String
        Get
            'Return SetUp.manageConnString
            Return ""
        End Get
    End Property

End Class
