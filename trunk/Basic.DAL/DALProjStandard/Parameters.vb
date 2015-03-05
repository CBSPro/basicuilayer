Imports System.Data
Imports System.Data.SqlClient
Imports Basic.Constants.ProjConst
Imports Basic.DAL.Utils

Public Module Parameters

    Dim objConnection As Object

    Public TotPAmt As Double
    Public TotSC As Double
    Public pCash As String
    Public pBank As String
    Public pReceivable As String
    Public pPayable As String
    Public pIncome As String
    Public pExpense As String
    Public pRetEarn As String
    Public pBranchCode As String
    Public pDeptCode As String
    Public FldName As String
    Public FldValue As String

    Public Sub SetParms()
        ''
        pCash = Left(Trim(GetParms("CASH")), 2)
        pBank = Left(Trim(GetParms("BANK")), 2)
        pReceivable = Left(Trim(GetParms("DEBTORS")), 2)
        pPayable = Left(Trim(GetParms("CREDITORS")), 2)
        pIncome = Trim(GetParms("INCOME"))
        pExpense = Trim(GetParms("EXPENSE"))
        pRetEarn = Trim(GetParms("RETAINED EARNING"))
        pBranchCode = Trim(GetParms("BRANCH"))
        SysBrCode = pBranchCode
        pDeptCode = Trim(GetParms("DEPARTMENT"))
        ''
    End Sub

    Public Function GetParms(ByVal pFieldName As String) As String

        objConnection = cConnectionManager.GetConnection()
        Dim objDatabaseManager As IDatabaseManager
        Dim objDBParameters As New cDBParameterList
        objDatabaseManager = cDataBaseManager.GetDatabaseManager()
        objDatabaseManager.SetConnection(objConnection)

        objDBParameters.AddParameter("@FldName", pFieldName, "nvarchar")
        objDatabaseManager.ExecuteNonQuery("Sp_GetParms", objDBParameters)
        cConnectionManager.CloseConnection(objConnection)
        mSQL = "Select FldValue From Parms Where FldName = '" & pFieldName & "'"
        FldValue = GetFldValue(mSQL, "FldValue")
        Return FldValue
    End Function
End Module
