
Imports System
Imports System.Data.SqlClient

Public Class ProjConst


    'Public Shared SqlCon As SqlClient.SqlConnection
    Public Shared sqlconn As New SqlConnection
    Public Shared DAdp As New SqlDataAdapter
    Public Shared strConec As String
    Public Shared mStr As String
    Public Shared mTemp As String
    Public Shared mcboBranch As String

    Public Shared mfrmBP As Boolean
    Public Shared mfrmBR As Boolean
    Public Shared mfrmCP As Boolean
    Public Shared mfrmCR As Boolean
    Public Shared mfrmJV As Boolean

    'Admin Menu
    Public Shared mUMaintenence As Boolean
    Public Shared mCPass As Boolean
    Public Shared mUPass As Boolean
    Public Shared mBckDB As Boolean
    Public Shared mEPostVoch As Boolean
    Public Shared mSysParm As Boolean

    'Reports Menu
    'On Screen View Menu
    Public Shared mGLScr As Boolean
    Public Shared mCStatus As Boolean
    Public Shared mBBookScr As Boolean
    Public Shared mCBookScr As Boolean
    Public Shared mAcCStatus As Boolean
    Public Shared mBBook As Boolean
    Public Shared mCBook As Boolean
    Public Shared mTBal As Boolean
    Public Shared mTBalCross As Boolean
    Public Shared mATrial As Boolean
    Public Shared mSubLedg As Boolean
    Public Shared mAAnalysis As Boolean
    Public Shared mPDefRep As Boolean
    Public Shared mPDefRefCross As Boolean
    Public Shared mPPstVouchers As Boolean

    Public Shared mfrmEPinfo As Boolean
    Public Shared mfrmECinfo As Boolean
    Public Shared mfrmEQinfo As Boolean
    Public Shared mfrmEduinfo As Boolean
    Public Shared mfrmEmpExp As Boolean
    Public Shared mfrmEmpRel As Boolean
    Public Shared mfrmEOinfo As Boolean
    Public Shared mfrmDAManual As Boolean
    Public Shared mfrmOTDedCalc As Boolean
    Public Shared mfrmEmpAllowDed As Boolean
    Public Shared mfrmPromInfo As Boolean
    Public Shared mfrmTranInfo As Boolean
    Public Shared mfrmLeaveInfo As Boolean
    Public Shared mfrmLWOPay As Boolean
    Public Shared mfrmTermRet As Boolean
    Public Shared mfrmSInfo As Boolean
    Public Shared mfrmAdvInfo As Boolean
    Public Shared mfrmCSalSheet As Boolean
    Public Shared mfrmBranch As Boolean
    Public Shared mfrmDept As Boolean
    Public Shared mfrmSection As Boolean
    Public Shared mfrmDesig As Boolean
    Public Shared mfrmLeaveType As Boolean
    Public Shared mfrmReligion As Boolean
    Public Shared mfrmLEntitle As Boolean
    Public Shared mfrmShift As Boolean
    Public Shared mfrmHolidays As Boolean
    Public Shared mfrmFACodes As Boolean
    Public Shared mfrmManufacturer As Boolean
    Public Shared mfrmSupplier As Boolean
    Public Shared mfrmSupplierType As Boolean
    Public Shared mfrmFAAddition As Boolean
    Public Shared mfrmFADisposal As Boolean
    Public Shared mfrmFAAdjustment As Boolean
    Public Shared mfrmFATransfer As Boolean
    Public Shared mfrmFADepriciation As Boolean
    Public Shared mFilter As Boolean
    Public Shared mfrmSales As Boolean
    Public Shared mfrmSalesReturn As Boolean
    Public Shared mfrmCustomer As Boolean

    'Basic Data
    Public Shared mBranch As Boolean
    Public Shared mAccMaintanance As Boolean
    Public Shared mBudget As Boolean
    Public Shared mDefReport As Boolean


    Public Shared mfrmAcc As Boolean
    Public Shared mfrmBudget As Boolean
    Public Shared mfrmRepDef As Boolean
    Public Shared mfrmBrowseGL As Boolean
    Public Shared mfrmACStatus As Boolean
    Public Shared mfrmBnkBook As Boolean
    Public Shared mfrmCashBook As Boolean

    Public Shared mfrmStore As Boolean
    Public Shared mfrmManufacture As Boolean
    Public Shared mfrmCategory As Boolean
    Public Shared mfrmItems As Boolean
    Public Shared mfrmPayTerm As Boolean
    Public Shared mfrmDelTerm As Boolean

    Public Shared mfrmPO As Boolean
    Public Shared mfrmPurchase As Boolean
    Public Shared mfrmPurchaseReturn As Boolean
    Public Shared mfrmIssue As Boolean
    Public Shared mfrmIssueReturn As Boolean
    Public Shared mfrmTransferIn As Boolean
    Public Shared mfrmTransferOut As Boolean
    Public Shared mfrmAdjustment As Boolean
    Public Shared mfrmPReqsition As Boolean
    Public Shared mFrmIssueToFA As Boolean
    Public Shared mfrmIssueFAReturn As Boolean





    'Public Shared StrConec As String
    Public Shared sqlCon As New SqlConnection
    Public Shared SqlTrans As SqlTransaction
    Public Shared SysDataBase As String
    Public Shared BuiltInUser As String
    Public Shared SysUserID As String
    Public Shared SysUser As String
    Public Shared SysBrCode As String
    Public Shared SySDate As String
    Public Shared SysTime As String
    Public Shared SysPassword As String
    Public Shared SysStatus As String
    Public Shared SysCompany As String
    Public Shared SysDsn As String
    Public Shared SysServer As String
    Public Shared SysSystem As String
    Public Shared SysAddress1 As String
    Public Shared SysAddress2 As String
    Public Shared SpUser As Boolean
    Public Shared SpUsers As Integer
    Public Shared MenuName As String
    Public Shared pPara As Integer
    Public Shared pParaJV As Integer
    Public Shared SortSign As String
    Public Shared Fld1 As String
    Public Shared Fld2 As String
    Public Shared FlagUser As Boolean
    Public Shared RepPath As String
    Public Shared RepTitle As String
    Public Shared AVIPath As String
    Public Shared RepConn As String
    Public Shared StValue As String
    Public Shared EndValue As String
    Public Shared mOrder As Integer
    Public Shared mReport As Integer
    Public Shared mTotRecs As String
    Public Shared TRepType As String
    Public Shared RepType As String
    Public Shared mType As String
    Public Shared mVno As String
    Public Shared FlagDup As Boolean
    Public Shared Logged As Boolean
    Public Shared mRow As Integer
    Public Shared RptBalBF As String
    Public Shared RptBalCF As String
    Public Shared strHeader As String
    Public Shared strHeader1 As String
    Public Shared strHeader2 As String
    Public Shared strHeader3 As String
    Public Const SysPostedPass = "helloworld"

    ''
    Public Shared SDate As Date
    Public Shared EDate As Date
    Public Shared Branch As String
    Public Shared FltFirst As Boolean
    Public Shared FltSecond As Boolean
    Public Shared FltThird As Boolean
    Public Shared FltAll As Boolean
    Public Shared Fltpost As Boolean
    Public Shared FltUnpost As Boolean
    Public Shared mShowMenu As Boolean
    Public Shared FileAVI As Integer
    Public Shared EmpIDParam As String
    ''Accounts Information
    Public Shared StrCode As String
    Public Shared StrTitle As String
    Public Shared AccCode As String
    Public Shared Opt As MsgBoxResult
    Public Shared mSQL As String
    Public Shared repPathArr
    ''Error Messages
    Public Const ErrPK = -2147217873
    Public Const ErrLock = -2147217871
    Public Const ErrLockEdit = -2147217864
    Public Const ErrRfIntegrity = -2147217873
    Public Const ErrEfStstus = -2147467259


    '''''Accounts


    Public Shared aP1P As Integer
    Public Shared aP1L As Integer
    Public Shared aP2P As Integer
    Public Shared aP2L As Integer
    Public Shared aP3P As Integer
    Public Shared aP3L As Integer
    Public Shared aP4P As Integer
    Public Shared aP4L As Integer
    Public Shared aCodeL As Integer
    Public Shared aICodeL As Integer

    '' Variables
    Dim TLBal As Double
    Dim TCBal As Double
    Dim TABal As Double
    Dim TLBudg As Double
    Dim TCBudg As Double
    Dim TABudg As Double
    Dim D_TLBal As Double
    Dim D_TCBal As Double
    Dim D_TABal As Double
    Dim mi As Double
    Dim OPERATORS As String

    ' 4 Level GL System Parameters
    'Public Const aEmptyCode = "  -  -  -     "
    'Public Const aMaskCode = "##-##-##-#####"
    'Public Const aEmpMaskCode = "  -  -  -"
    'Public Const aMaskFGrid = "99-99-99-99999"
    'Public Const aMaskLkp = "@@-@@-@@-@@@@@"
    'Public Const mEmpty = "99-99-99-99999"
    Public Const NumFmt = "###,###,###,##0.00"
    Public Const NumFmt1 = "(###,###,###,##0.00)"
    ' 3 Level GL System Parameters
    Public Const aEmptyCode = "  -  -    "
    Public Const aMaskCode = "##-##-####"
    Public Const aMaskTDed = "999"
    Public Const aEmpMaskCode = "  -  -"
    Public Const aMaskFGrid = "99-99-9999"
    Public Const aMaskLkp = "@@-@@-@@@@"
    Public Const mEmpty = "99-99-9999"



    Public Shared Sub Main()
        BuiltInUser = "SA"
        SysSystem = "ERP System"
        SysCompany = "Madni Poly Tex Pvt Ltd"
        SysAddress1 = ""
        SysAddress2 = ""
        SortSign = " *"
        RepTitle = "CBS"
        'SysServer = "CBS-02\SQLExpress"
        'SysServer = "CBS-99"
        'SySDate = GetDate(Now.Date, "dd-MMM-yyyy")
        'SySDate = Format(Now.Date, "dd-MMM-yyyy")
        'SysTime = System.DateTime.Now
        'SySCon.CommandTimeout = 10

        'SysDataBase = "JKFS_sz"
        'SysPassword = "147"
        'RepPath = "E:\Projects.Net\Accounts\Accounts\Reports\"
        'RepPath = "G:\Vb Projects\Professional Projects\Projects.Net\Accounts\Accounts\Reports\"
        RepPath = System.AppDomain.CurrentDomain.BaseDirectory()
        repPathArr = RepPath.Split("bin")
        RepPath = repPathArr(0) & "Reports\"
        'AVIPath = "E:\Projects.Net\Accounts\AviFiles\"
        AVIPath = "G:\Vb Projects\Professional Projects\Projects.Net\Accounts\AviFiles\"
        'SysDsn = "AccSys"
        'SysDsn = "AccSysJKFS_GL"
        SysDsn = "ERP"
        'frmLogin.ShowDialog()
    End Sub


End Class
