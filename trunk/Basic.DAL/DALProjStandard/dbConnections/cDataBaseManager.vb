Imports System.Data
Imports System.Data.SqlClient

Public Class cDataBaseManager

    Public Shared Function GetDatabaseManager() As IDatabaseManager

        Dim DBType As String
        DBType = "SQLSERVER" 'System.Configuration.ConfigurationSettings.AppSettings("DatabaseType").ToString.ToUpper

        Select Case DBType
            Case "SQLSERVER"
                Dim objSqlDatabaseManager As New cSqlDatabaseManager
                Return objSqlDatabaseManager
            Case Else
                Throw New Exception("DB1")
        End Select
    End Function

End Class
