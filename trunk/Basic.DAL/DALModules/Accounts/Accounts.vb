Imports System.Data
Imports System.Data.SqlClient
Imports Basic.DAL
Imports Basic.Constants.ProjConst
Imports Basic.DAL.Utils

Public Module Accounts
    Public aP1P As Integer
    Public aP1L As Integer
    Public aP2P As Integer
    Public aP2L As Integer
    Public aP3P As Integer
    Public aP3L As Integer
    Public aP4P As Integer
    Public aP4L As Integer
    Public aCodeL As Integer
    Public aICodeL As Integer

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


    Public Sub SetAccParam()
        aP1P = 1
        aP1L = 2
        aP2P = 3
        aP2L = 2
        aP3P = 5
        aP3L = 4
        aP4P = 7
        'aP4L = 5

        'aCodeL = aP1L + aP2L + aP3L + aP4L
        aCodeL = aP1L + aP2L + aP3L
    End Sub
    Public Function aDeptCtrlExist(ByVal mCode As String) As Boolean
        Dim mCtrl As String
        Dim mSQL As String

        mCtrl = Mid(mCode, aP1P, aP1L)

        If Mid(mCode, aP1P, aP1L) <> Fillit("", "0", aP1L) Then   ' Then 'Compare With Zero
            mCtrl = Mid(mCode, aP2P, aP2L)
            If Mid(mCode, aP2P, aP2L) <> Fillit("", "0", aP2L) Then   ' Then 'Compare With Zero
                If Len(Trim(mCtrl)) <> 0 Then
                    mSQL = "SELECT deptCode FROM Depts " & _
                           "WHERE " & _
                           "DeptCode = '" & Mid(mCode, 1, aP1L) & Fillit("", "0", aP2L) & "'"
                    If Not Found(mSQL) Then
                        aDeptCtrlExist = False
                        Exit Function
                    End If
                Else
                    aDeptCtrlExist = False
                    Exit Function
                End If
            Else

                aDeptCtrlExist = True
                Exit Function
            End If
        Else
            aDeptCtrlExist = False
            Exit Function
        End If

        aDeptCtrlExist = True
    End Function

    Public Function aCtrlExist(ByVal mCode As String) As Boolean
        Dim mCtrl As String
        Dim mSQL As String

        mCtrl = Mid(mCode, aP1P, aP1L)

        If Mid(mCode, aP1P, aP1L) <> Fillit("", "0", aP1L) Then   ' Then 'Compare With Zero
            mCtrl = Mid(mCode, aP2P, aP2L)
            If Mid(mCode, aP2P, aP2L) <> Fillit("", "0", aP2L) Then   ' Then 'Compare With Zero
                mCtrl = Mid(mCode, aP3P, aP3L)
                If Mid(mCode, aP3P, aP3L) <> Fillit("", "0", aP3L) Then   ' Then 'Compare With Zero
                    'mCtrl = Mid(mCode, aP4P, aP4L)
                    'If Mid(mCode, aP4P, aP4L) <> Fillit("", "0", aP4L) Then   ' Then 'Compare With Zero
                    If Len(Trim(mCtrl)) <> 0 Then
                        mSQL = "SELECT Code FROM Codes " & _
                               "WHERE " & _
                               "Code = '" & Mid(mCode, 1, aP1L) & Mid(mCode, aP2P, aP2L) & Fillit("", "0", aP3L) & "'"
                        '"Code = '" & Mid(mCode, 1, aP1L) & Mid(mCode, aP2P, aP2L) & Mid(mCode, aP3P, aP3L) & Fillit("", "0", aP4L) & "'"

                        If Not Found(mSQL) Then
                            aCtrlExist = False
                            Exit Function
                        End If
                    Else
                        aCtrlExist = False
                        Exit Function
                    End If
                Else
                    '            If Mid(mCode, aP3P, aP3L) <> Fillit("", "0", aP3L) Then   ' Then 'Compare With Zero
                    If Len(Trim(mCtrl)) <> 0 Then
                        mSQL = "SELECT Code FROM Codes " & _
                               "WHERE " & _
                               "Code = '" & Mid(mCode, 1, aP1L) & Fillit("", "0", aP2L) & Fillit("", "0", aP3L) & "'"

                        If Not Found(mSQL) Then
                            aCtrlExist = False
                            Exit Function
                        End If
                    Else
                        aCtrlExist = False
                        Exit Function
                    End If
                End If
                '            End If
            Else
                mCtrl = Mid(mCode, aP3P, aP3L)
                If Mid(mCode, aP3P, aP3L) <> Fillit("", "0", aP3L) Then   ' Then 'Compare With Zero
                    If Len(Trim(mCtrl)) <> 0 Then
                        mSQL = "SELECT Code FROM Codes " & _
                               "WHERE " & _
                               "Code = '" & Mid(mCode, 1, aP1L) & Mid(mCode, aP2P, aP2L) & Mid(mCode, aP3P, aP3L) & "'"
                        '                          "FaCode = '" & Mid(mCode, 1, aP1L) & Fillit("", "0", aP2L) & Fillit("", "0", aP3L) & "'"

                        If Not Found(mSQL) Then
                            aCtrlExist = False
                            Exit Function
                        End If
                    Else
                        aCtrlExist = False
                        Exit Function
                    End If
                End If
            End If
        End If
        aCtrlExist = True
    End Function
    Public Function fChkCode(ByVal mCode As String) As Boolean
        Dim mSQL As String
        If Len(Trim(mCode)) < aCodeL And Not ((Len(Trim(mCode)) = aP1L Or Len(Trim(mCode)) = aP1L + aP2L)) Then
            MsgBox("Invalid Code Length", vbInformation, SysCompany)
            fChkCode = False
            Exit Function
        ElseIf Len(Trim(mCode)) < aCodeL And (Len(Trim(mCode)) = aP1L Or Len(Trim(mCode)) = aP1L + aP2L) Then
            MsgBox("Control Code Not Allowed Here", vbInformation, SysCompany)
            fChkCode = False
            Exit Function
        Else
            fChkCode = True
        End If
        ''
        mSQL = "SELECT FaCODE,FaName FROM FaCodes " & _
                    "WHERE LEN(RTRIM(FaCODE)) = " & aCodeL & " AND " & _
                    "FaCODE = '" & mCode & "'"
        If Not Found(mSQL) Then
            MsgBox("This Fixed Asset Does Not Exist, No Transactions are Allowed", vbCritical, SysCompany)
            fChkCode = False
            Exit Function
        End If
        fChkCode = True
    End Function


    Public Function GetJVno(ByVal mBrCode As String, ByVal mType As String, ByVal mDate As Date) As String
        'Dim mDa As SqlDataAdapter
        'Dim mDt As New DataTable
        Dim mSQL As String
        Dim mNewNo As String
        Dim mMonth As String
        Dim mYear As String

        mYear = Format(Year(mDate), "0000")
        mMonth = Format(Month(mDate), "00")
        ''
        mSQL = "SELECT MAX(VNO) VNo FROM GLHead WHERE " & _
                "LEFT(VNO,4) = '" & mYear & "' AND " & _
                "SUBSTRING(VNO,5,2) = '" & mMonth & "' AND " & _
                "TYPE = '" & mType & "' And BrCode = '" & mBrCode & "'"

        'mDa = New SqlDataAdapter(mSQL, sqlCon)
        'mDa.Fill(mDt)
        mSQL = GetFldValue(mSQL, "VNo")
        If mSQL <> "" Then
            'If Not IsDBNull(mDt.Rows(0).Item(0)) Then
            ''mNewNo = Format(Val(mDt.Rows(0).Item(0)) + 1, "0000000000")
            mNewNo = Format(Val(mSQL) + 1, "0000000000")
        Else
            mNewNo = mYear & mMonth & Format(1, "0000")
        End If
        GetJVno = mNewNo
    End Function

    Public Function ChkCode(ByVal mCode As String) As Boolean
        Dim mSQL As String
        If (mCode <> "") Then
            If Len(Trim(mCode)) < aCodeL And Not ((Len(Trim(mCode)) = aP1L Or Len(Trim(mCode)) = aP1L + aP2L)) Then
                MsgBox("Invalid Code Length", vbInformation, SysCompany)
                ChkCode = False
                Exit Function
            ElseIf Len(Trim(mCode)) < aCodeL And (Len(Trim(mCode)) = aP1L Or Len(Trim(mCode)) = aP1L + aP2L) Then
                MsgBox("Control Code Not Allowed Here", vbInformation, SysCompany)
                ChkCode = False
                Exit Function
            Else
                ChkCode = True
            End If
            ''
            mSQL = "SELECT CODE,Description FROM Codes " & _
                        "WHERE LEN(RTRIM(CODE)) = " & aCodeL & " AND " & _
                        "CODE = '" & mCode & "'"
            If Not Found(mSQL) Then
                MsgBox("This Account Does Not Exist, No Transactions are Allowed", vbCritical, SysCompany)
                ChkCode = False
                Exit Function
            End If
            ChkCode = True
        End If
    End Function

    ''*****************************************************************************
    '' ************ PROCESS OF INCOME STATEMENT ***********************************
    ''*****************************************************************************

    Public Sub Process_IncStat(ByVal StrPgNo As String, ByVal EndPgNo As String, ByVal TDate As Date, ByVal PrtNotes As Boolean)

        Dim objIncStat As New cIncStat

        Dim dtPftLoss As New DataTable
        Dim dtDetail As New DataTable
        Dim dtCodesBal As New DataTable
        Dim rowNum As Integer
        Dim rowNum1 As Integer
        Dim rowNum2 As Integer
        Dim mTStr
        Dim DtCode As String
        Dim DtDesc As String
        Dim mCodeLen As Integer
        Dim TLimit As Integer
        Dim II As Integer
        Dim mBudg As String
        ''**
        Dim TPageNo As String
        Dim PrDate As Date
        Dim TLineNo As String
        Dim TValue As String
        Dim Tdesc As String
        Dim TNoteRef As String
        Dim TCode As String
        Dim TpNo As String
        Dim TLNo As String
        ''**
        Dim mAmt As Double
        Dim FAmt As Double
        Dim BAmt As Double
        Dim TLLBal As Double
        Dim TCLBal As Double
        Dim TALBaL As Double
        Dim TLLBudg As Double
        Dim TCLBudg As Double
        Dim TALBudg As Double
        ''**
        Dim TCdlBal As Double
        Dim TCdcBal As Double
        Dim TCdaBal As Double
        Dim TCdlBudg As Double
        Dim TCdcBudg As Double
        Dim TCdaBudg As Double
        ''**
        Dim DtCdlBal As Double
        Dim DtCdcBal As Double
        Dim DtCdaBal As Double
        Dim DtCdlBudg As Double
        Dim DtCdcBudg As Double
        Dim DtCdaBudg As Double
        ''**
        Dim getBudget As Boolean

        ''** TO DELETE ALL PREVIUOS ENTRIES IN DETAIL
        objIncStat.DelDetail()
        ''** END OF DELETE PROCESS IN DETAIL
        dtDetail = objIncStat.SelDetail()
        ''*** GET PREVIOUS DATE
        PrDate = PMonth(TDate)
        ''*** PFTLOSS TABLE QUERY
        dtPftLoss = objIncStat.SelPftLoss()
        rowNum = dtPftLoss.Rows.Count - 1
        ''*****************
        ''*** MAIN LOOP
        ''*****************
        If rowNum >= 0 Then
            rowNum = 0
            Do While rowNum < dtPftLoss.Rows.Count
                TPageNo = dtPftLoss.Rows(rowNum).Item("PageNo")
                Do While dtPftLoss.Rows(rowNum).Item("PageNo") = TPageNo
                    TLineNo = dtPftLoss.Rows(rowNum).Item("mLineNo")
                    TValue = dtPftLoss.Rows(rowNum).Item("mValue")
                    Tdesc = dtPftLoss.Rows(rowNum).Item("Description")
                    TNoteRef = dtPftLoss.Rows(rowNum).Item("NoteRef")
                    TRepType = dtPftLoss.Rows(rowNum).Item("RepType")
                    mi = 1
                    TCode = ""
                    TpNo = TPageNo              '' PAGE NUMBER
                    TLNo = ""                   '' LINE NUMBER
                    mAmt = 0                    '' NUMERIC VALUE TO BE ADDED IN THE MONTH
                    FAmt = 0                    '' NUMERIC VALUE TO BE ADDED IN FIRST COLUMN
                    BAmt = 0                    '' NUMERIC VALUE TO BE ADDED IN BOTH COLUMN
                    TLBal = 0                   '' UP TO LAST MONTH BALANCE
                    TCBal = 0                   '' CURRENT MONTH BALANCE
                    TABal = 0                   '' TO DATE BALANCE
                    TLBudg = 0                  '' UP TO LAST MONTH BUDGET
                    TCBudg = 0                  '' CURRENT MONTH BUDGET
                    TABudg = 0                  '' TO DATE BUDGET
                    TLLBal = 0
                    TCLBal = 0
                    TALBaL = 0
                    TLLBudg = 0
                    TCLBudg = 0
                    TALBudg = 0
                    Do
                        Do While mi < Len(TValue)
                            If Mid(TValue, mi, 1) = Space(1) Then
                                mi = mi + 1
                                Exit Do
                            End If
                            '' START OF SELECT OF  mI CHAR
                            Select Case Mid(TValue, mi, 1)
                                Case "C"
                                    mi = mi + 1
                                    TCode = ""
                                    RepType = "C"
                                    Do While Mid(TValue, mi, 1) <> "" And Mid(TValue, mi, 1) <> " "
                                        TCode = TCode + Mid(TValue, mi, 1)
                                        mi = mi + 1
                                    Loop
                                Case "A"
                                    mi = mi + 1
                                    TCode = ""
                                    RepType = "A"
                                    Do While Mid(TValue, mi, 1) <> "" And Mid(TValue, mi, 1) <> " "
                                        TCode = TCode + Mid(TValue, mi, 1)
                                        mi = mi + 1
                                    Loop
                                Case "O"
                                    mi = mi + 1
                                    TCode = ""
                                    RepType = "O"
                                    Do While Mid(TValue, mi, 1) <> "" And Mid(TValue, mi, 1) <> " "
                                        TCode = TCode + Mid(TValue, mi, 1)
                                        mi = mi + 1
                                    Loop
                                Case "D"
                                    mi = mi + 1
                                    TCode = ""
                                    RepType = "D"
                                    Do While Mid(TValue, mi, 1) <> "" And Mid(TValue, mi, 1) <> " "
                                        TCode = TCode + Mid(TValue, mi, 1)
                                        mi = mi + 1
                                    Loop
                                Case "R"
                                    mi = mi + 1
                                    TCode = ""
                                    RepType = "R"
                                    Do While Mid(TValue, mi, 1) <> "" And Mid(TValue, mi, 1) <> " "
                                        TCode = TCode + Mid(TValue, mi, 1)
                                        mi = mi + 1
                                    Loop
                                Case "+", "-", "/", "*", "^"
                                    OPERATORS = Mid(TValue, mi, 1)
                                    mi = mi + 1
                                    Exit Do
                                Case "L"   ''** ADDITION OF LINE IN DETAIL
                                    mi = mi + 1
                                    TLNo = ""
                                    Do While Mid(TValue, mi, 1) <> "" And Mid(TValue, mi, 1) <> " "
                                        TLNo = TLNo + Mid(TValue, mi, 1)
                                        mi = mi + 1
                                    Loop
                                    TLLBal = 0
                                    TCLBal = 0
                                    TALBaL = 0
                                    TLLBudg = 0
                                    TCLBudg = 0
                                    TALBudg = 0
                                    If Trim(TpNo) <> "" And Trim(TLNo) <> "" Then
                                        mSQL = "SELECT LBalance,CBalance,ABalance," & _
                                               "LBudget,CBudget,ABudget FROM PFTLOSS " & _
                                               "WHERE PAGENO = '" & TpNo & "' AND " & _
                                               "mLINENO = '" & TLNo & "'"
                                    ElseIf Trim(TLNo) <> "" And Trim(TpNo) = "" Then
                                        mSQL = "SELECT LBalance,CBalance,ABalance," & _
                                               "LBudget,CBudget,ABudget FROM PFTLOSS " & _
                                               "Where PageNo = '" & TPageNo & "' AND " & _
                                               "mLINENO = '" & TLNo & "'"
                                    End If
                                    TLLBal = Val(GetFldValue(mSQL, "LBalance"))
                                    TCLBal = Val(GetFldValue(mSQL, "CBalance"))
                                    TALBaL = Val(GetFldValue(mSQL, "ABalance"))
                                    TLLBudg = Val(GetFldValue(mSQL, "LBudget"))
                                    TCLBudg = Val(GetFldValue(mSQL, "CBudget"))
                                    TALBudg = Val(GetFldValue(mSQL, "ABudget"))
                                    If Len(Trim(TNoteRef)) <> 0 Then
                                        objIncStat.PageNo = TPageNo
                                        objIncStat.mLineNo = TLineNo
                                        objIncStat.NoteRef = TNoteRef
                                        objIncStat.Description = Tdesc
                                        objIncStat.LBalance = TLLBal
                                        objIncStat.CBalance = TCLBal
                                        objIncStat.ABalance = TALBaL
                                        objIncStat.LBudget = TLLBudg
                                        objIncStat.CBudget = TCLBudg
                                        objIncStat.ABudget = TALBudg
                                        objIncStat.SaveDetail()
                                    End If
                                Case "P"
                                    mi = mi + 1
                                    TpNo = ""
                                    Do While Mid(TValue, mi, 1) <> "" And Mid(TValue, mi, 1) <> " "
                                        TpNo = TpNo + Mid(TValue, mi, 1)
                                        mi = mi + 1
                                    Loop
                                    Exit Do
                                Case "M"
                                    mi = mi + 1
                                    mTStr = ""
                                    Do While Mid(TValue, mi, 1) <> "" And Mid(TValue, mi, 1) <> " "
                                        mTStr = mTStr + Mid(TValue, mi, 1)
                                        mi = mi + 1
                                    Loop
                                    mAmt = Val(mTStr)
                                Case "F"

                                    mi = mi + 1
                                    mTStr = ""
                                    Do While Mid(TValue, mi, 1) <> "" And Mid(TValue, mi, 1) <> " "
                                        mTStr = mTStr + Mid(TValue, mi, 1)
                                        mi = mi + 1
                                    Loop
                                    FAmt = Val(mTStr)
                                Case "B"
                                    mi = mi + 1
                                    mTStr = ""
                                    Do While Mid(TValue, mi, 1) <> "" And Mid(TValue, mi, 1) <> " "
                                        mTStr = mTStr + Mid(TValue, mi, 1)
                                        mi = mi + 1
                                    Loop
                                    BAmt = Val(mTStr)
                                Case Else
                                    mi = mi + 1
                            End Select
                            '' END OF SELECT OF mI CHAR
                            ''
                            ''** START OF IF STATEMENT
                            If Len(Trim(TCode)) <> 0 Then
                                TCdlBal = 0     '' BALANCE OF CODE UPTO LAST MONTH
                                TCdcBal = 0     '' BALANCE OF CODE FOR CURRENT MONTH
                                TCdaBal = 0     '' BALANCE OF CODE UPTO CURRENT MONTH
                                TCdlBudg = 0    '' BUDGET OF CODE UPTO LAST MONTH
                                TCdcBudg = 0    '' BUDGET OF CODE FOR CURRENT MONTH
                                TCdaBudg = 0    '' BUDGET OF CODE UPTO CURRENT MONTH
                                mCodeLen = Len(Trim(TCode))
                                TCode = Trim(TCode)
                                If PrtNotes Then
                                    If TRepType = "C" Then
                                        If RepType = "C" Then
                                            mSQL = GetSqlCF(PrDate, TDate, TCode, mCodeLen, True)
                                        ElseIf RepType = "A" Then
                                            mSQL = GetSqlAF(PrDate, TDate, TCode, mCodeLen, True)
                                        ElseIf RepType = "O" Then
                                            mSQL = GetSqlOF(PrDate, TDate, TCode, mCodeLen, True)
                                        ElseIf RepType = "D" Then
                                            mSQL = GetSqlDF(PrDate, TDate, TCode, mCodeLen, True)
                                        ElseIf RepType = "R" Then
                                            mSQL = GetSqlRF(PrDate, TDate, TCode, mCodeLen, True)
                                        End If
                                    Else
                                        If RepType = "C" Then
                                            mSQL = GetSqlC(PrDate, TDate, TCode, mCodeLen, True)
                                        ElseIf RepType = "A" Then
                                            mSQL = GetSqlA(PrDate, TDate, TCode, mCodeLen, True)
                                        ElseIf RepType = "O" Then
                                            mSQL = GetSqlO(PrDate, TDate, TCode, mCodeLen, True)
                                        ElseIf RepType = "D" Then
                                            mSQL = GetSqlD(PrDate, TDate, TCode, mCodeLen, True)
                                        ElseIf RepType = "R" Then
                                            mSQL = GetSqlR(PrDate, TDate, TCode, mCodeLen, True)
                                        End If
                                    End If
                                Else
                                    If TRepType = "C" Then
                                        If RepType = "C" Then
                                            mSQL = GetSqlCF(PrDate, TDate, TCode, mCodeLen, False)
                                        ElseIf RepType = "A" Then
                                            mSQL = GetSqlAF(PrDate, TDate, TCode, mCodeLen, False)
                                        ElseIf RepType = "O" Then
                                            mSQL = GetSqlOF(PrDate, TDate, TCode, mCodeLen, False)
                                        ElseIf RepType = "D" Then
                                            mSQL = GetSqlDF(PrDate, TDate, TCode, mCodeLen, False)
                                        ElseIf RepType = "R" Then
                                            mSQL = GetSqlRF(PrDate, TDate, TCode, mCodeLen, False)
                                        End If
                                    Else
                                        If RepType = "C" Then
                                            mSQL = GetSqlC(PrDate, TDate, TCode, mCodeLen, False)
                                        ElseIf RepType = "A" Then
                                            mSQL = GetSqlA(PrDate, TDate, TCode, mCodeLen, False)
                                        ElseIf RepType = "O" Then
                                            mSQL = GetSqlO(PrDate, TDate, TCode, mCodeLen, False)
                                        ElseIf RepType = "D" Then
                                            mSQL = GetSqlD(PrDate, TDate, TCode, mCodeLen, False)
                                        ElseIf RepType = "R" Then
                                            mSQL = GetSqlR(PrDate, TDate, TCode, mCodeLen, False)
                                        End If
                                    End If
                                End If
                                dtCodesBal = GetDataTable(mSQL)
                                rowNum1 = dtCodesBal.Rows.Count - 1
                                getBudget = True
                                If rowNum1 >= 0 Then
                                    rowNum1 = 0
                                    Do While rowNum1 < dtCodesBal.Rows.Count
                                        DtCode = dtCodesBal.Rows(rowNum1).Item("Code")
                                        DtCdlBal = 0
                                        DtCdcBal = 0
                                        DtCdaBal = 0
                                        DtCdlBudg = 0
                                        DtCdcBudg = 0
                                        DtCdaBudg = 0
                                        If Not IsDBNull(dtCodesBal.Rows(rowNum1).Item("DtCdlBal")) Then
                                            DtCdlBal = dtCodesBal.Rows(rowNum1).Item("DtCdlBal")
                                        End If
                                        If Not IsDBNull(dtCodesBal.Rows(rowNum1).Item("DtCdaBal")) Then
                                            DtCdaBal = dtCodesBal.Rows(rowNum1).Item("DtCdaBal")
                                        End If
                                        DtCdcBal = DtCdaBal - DtCdlBal
                                        ''** GET BUDGET VALUE FROM CODES
                                        If getBudget Then
                                            mSQL = BudgetSql(TCode, mCodeLen)
                                            dtDetail = GetDataTable(mSQL)
                                            rowNum2 = dtDetail.Rows.Count - 1
                                            If rowNum2 >= 0 Then
                                                II = 1
                                                TLimit = Month(TDate)
                                                Do While II < TLimit
                                                    mBudg = "BUDGET" & Trim(Str(II))
                                                    If Not IsDBNull(dtDetail.Rows(rowNum2).Item("mBudg")) Then
                                                        DtCdlBudg = DtCdlBudg + dtDetail.Rows(rowNum2).Item("mBudg").ToString
                                                    End If
                                                    II = II + 1
                                                Loop
                                                mBudg = "BUDGET" & Trim(Str(II))
                                                If Not IsDBNull(dtDetail.Rows(rowNum2).Item("mBudg")) Then
                                                    DtCdcBudg = DtCdcBudg + dtDetail.Rows(rowNum2).Item("mBudg").ToString
                                                End If
                                            End If
                                            DtCdaBudg = DtCdlBudg + DtCdcBudg
                                            getBudget = False
                                        End If
                                        DtDesc = dtCodesBal.Rows(rowNum1).Item("Description") & ""
                                        objIncStat.PageNo = TPageNo
                                        objIncStat.mLineNo = TLineNo
                                        objIncStat.NoteRef = TNoteRef
                                        objIncStat.Description = Trim(Tdesc)
                                        objIncStat.RepType = RepType
                                        objIncStat.Code = DtCode
                                        objIncStat.CDesc = Trim(DtDesc)
                                        objIncStat.LBudget = DtCdlBudg
                                        objIncStat.CBudget = DtCdcBudg
                                        objIncStat.ABudget = DtCdaBudg

                                        If OPERATORS = "-" Then
                                            objIncStat.LBalance = 0 - DtCdlBal
                                            objIncStat.CBalance = 0 - DtCdcBal
                                            objIncStat.ABalance = 0 - DtCdaBal
                                            DtCdlBudg = 0 - DtCdlBudg
                                            DtCdcBudg = 0 - DtCdcBudg
                                            DtCdaBudg = 0 - DtCdaBudg
                                        Else
                                            objIncStat.LBalance = DtCdlBal
                                            objIncStat.CBalance = DtCdcBal
                                            objIncStat.ABalance = DtCdaBal
                                        End If
                                        objIncStat.AppendDetail()
                                        TCdlBal = TCdlBal + DtCdlBal
                                        TCdcBal = TCdcBal + DtCdcBal
                                        TCdaBal = TCdaBal + DtCdaBal
                                        TCdlBudg = TCdlBudg + DtCdlBudg
                                        TCdcBudg = TCdcBudg + DtCdcBudg
                                        TCdaBudg = TCdaBudg + DtCdaBudg
                                        rowNum1 = rowNum + 1
                                    Loop
                                    Call SumBal(TCdlBal, TCdcBal, TCdaBal)
                                    Call SumBudg(TCdlBudg, TCdcBudg, TCdaBudg)
                                    TCode = ""
                                End If
                            ElseIf mAmt <> 0 Then
                                Call SumBal(0, mAmt, mAmt)
                                mAmt = 0
                            ElseIf FAmt <> 0 Then
                                Call SumBal(FAmt, 0, FAmt)
                                FAmt = 0
                            ElseIf BAmt <> 0 Then
                                Call SumBal(BAmt, BAmt, BAmt)
                                BAmt = 0
                            ElseIf TLLBal <> 0 Or TCLBal <> 0 Or TALBaL <> 0 Then
                                Call SumBal(TLLBal, TCLBal, TALBaL)
                                Call SumBudg(TLLBudg, TCLBudg, TALBudg)
                                TLLBal = 0
                                TCLBal = 0
                                TALBaL = 0
                                TLLBudg = 0
                                TCLBudg = 0
                                TALBudg = 0
                            End If
                        Loop
                    Loop Until Not (mi < Len(TValue))
                    ''** ENTRIES IN PFTLOSS
                    objIncStat.PageNo = TPageNo
                    objIncStat.mLineNo = TLineNo
                    objIncStat.UpdatePftLoss()
                    rowNum = rowNum + 1
                Loop
            Loop
        End If
    End Sub

    Private Function GetSqlCF(ByVal PrDate As Date, ByVal TDate As Date, ByVal TCode As String, ByVal CodeLen As Integer, ByVal PrtNotes As Boolean) As String
        Dim mSQL As String
        Dim PrStr As String
        Dim TmpStr As String

        PrStr = GetCmpDate(PrDate)
        TmpStr = Format(TDate, "dd-MMM-yyyy")
        If PrtNotes Then

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description, SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
               "SUM(DEBIT-CREDIT) DtCdaBal, " & _
               "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN DEBIT-CREDIT ELSE 0 END + " & _
               "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
               "THEN 0 ELSE DEBIT-CREDIT END) DtCdlbal " & _
             "From Gledg, Codes, " & _
                   "(SELECT DISTINCT SUBSTRING(CODE,1," & CodeLen & ") SUBCODE FROM Gledg " & _
                   " WHERE SUBSTRING(CODE,1," & CodeLen & ")='" & TCode & "') G " & _
               "WHERE Type In (Select CFType From CFTypes) And VDATE <= '" & Format(TmpStr, "dd-MMM-yyyy") & "' " & _
               "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
               "AND GLEDG.POSTED = 'Y' " & _
               "AND SUBSTRING(Gledg.CODE,1," & CodeLen & " ) = G.SUBCODE " & _
               "AND Codes.CODE = G.SUBCODE+'" & Fillit("", "0", aCodeL - CodeLen) & "'" & _
               " GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ") "

        Else

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description,SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
                   "SUM(DEBIT-CREDIT) DtCdaBal, " & _
                   "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN DEBIT-CREDIT ELSE 0 END + " & _
                   "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
                   "THEN 0 ELSE DEBIT-CREDIT END) DtCdlbal " & _
               "FROM Gledg, Codes " & _
                   "WHERE Type In (Select CFType From CFTypes) And VDATE <= '" & Format(TmpStr, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & ") = '" & TCode & "' " & _
                   "AND Codes.CODE = '" & TCode & Fillit("", "0", aCodeL - CodeLen) & "' " & _
                   "GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ")"

        End If
        GetSqlCF = mSQL

    End Function

    Private Function GetSqlAF(ByVal PrDate As Date, ByVal TDate As Date, ByVal TCode As String, ByVal CodeLen As Integer, ByVal PrtNotes As Boolean) As String
        Dim mSQL As String
        Dim PrStr As String
        Dim TmpStr As String

        PrStr = GetCmpDate(PrDate)
        TmpStr = Format(TDate, "dd-MMM-yyyy")
        If PrtNotes Then

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description, SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
               "SUM(DEBIT-CREDIT) DtCdaBal, " & _
               "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN DEBIT-CREDIT ELSE 0 END + " & _
               "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
               "THEN 0 ELSE DEBIT-CREDIT END) DtCdlbal " & _
             "From Gledg, Codes, " & _
                   "(SELECT DISTINCT SUBSTRING(CODE,1," & CodeLen & ") SUBCODE FROM Gledg " & _
                   " WHERE SUBSTRING(CODE,1," & CodeLen & ")='" & TCode & "') G " & _
                   "WHERE Type In (Select CFType From CFTypes) And VDATE Between " & _
                   "'" & Format(SDate, "dd-MMM-yyyy") & "' And " & _
                   "'" & Format(TmpStr, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & " ) = G.SUBCODE " & _
                   "AND Codes.CODE = G.SUBCODE+'" & Fillit("", "0", aCodeL - CodeLen) & "'" & _
                   " GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ") "

        Else

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description,SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
                   "SUM(DEBIT-CREDIT) DtCdaBal, " & _
                   "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN DEBIT-CREDIT ELSE 0 END + " & _
                   "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
                   "THEN 0 ELSE DEBIT-CREDIT END) DtCdlbal " & _
               "FROM Gledg, Codes " & _
                   "WHERE Type In (Select CFType From CFTypes) And VDATE Between " & _
                   "'" & Format(SDate, "dd-MMM-yyyy") & "' And " & _
                   "'" & Format(TmpStr, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & ") = '" & TCode & "' " & _
                   "AND Codes.CODE = '" & TCode & Fillit("", "0", aCodeL - CodeLen) & "' " & _
                   "GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ")"

        End If
        GetSqlAF = mSQL

    End Function

    Private Function GetSqlOF(ByVal PrDate As Date, ByVal TDate As Date, ByVal TCode As String, ByVal CodeLen As Integer, ByVal PrtNotes As Boolean) As String
        Dim mSQL As String
        Dim PrStr As String
        Dim TmpStr As String

        PrStr = GetCmpDate(PrDate)
        TmpStr = Format(TDate, "dd-MMM-yyyy")
        If PrtNotes Then

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description, SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
               "SUM(DEBIT-CREDIT) DtCdaBal, " & _
               "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN DEBIT-CREDIT ELSE 0 END + " & _
               "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
               "THEN 0 ELSE DEBIT-CREDIT END) DtCdlbal " & _
             "From Gledg, Codes, " & _
                   "(SELECT DISTINCT SUBSTRING(CODE,1," & CodeLen & ") SUBCODE FROM Gledg " & _
                   " WHERE SUBSTRING(CODE,1," & CodeLen & ")='" & TCode & "') G " & _
                   "WHERE Type In (Select CFType From CFTypes) And " & _
                   "VDATE < '" & Format(SDate, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & " ) = G.SUBCODE " & _
                   "AND Codes.CODE = G.SUBCODE+'" & Fillit("", "0", aCodeL - CodeLen) & "'" & _
                   " GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ") "

        Else

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description,SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
                   "SUM(DEBIT-CREDIT) DtCdaBal, " & _
                   "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN DEBIT-CREDIT ELSE 0 END + " & _
                   "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
                   "THEN 0 ELSE DEBIT-CREDIT END) DtCdlbal " & _
               "FROM Gledg, Codes " & _
                   "WHERE Type In (Select CFType From CFTypes) And " & _
                   "VDATE < '" & Format(SDate, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & ") = '" & TCode & "' " & _
                   "AND Codes.CODE = '" & TCode & Fillit("", "0", aCodeL - CodeLen) & "' " & _
                   "GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ")"

        End If
        GetSqlOF = mSQL

    End Function

    Private Function GetSqlDF(ByVal PrDate As Date, ByVal TDate As Date, ByVal TCode As String, ByVal CodeLen As Integer, ByVal PrtNotes As Boolean) As String
        Dim mSQL As String
        Dim PrStr As String
        Dim TmpStr As String

        PrStr = GetCmpDate(PrDate)
        TmpStr = Format(TDate, "dd-MMM-yyyy")
        If PrtNotes Then

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description, SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
               "SUM(DEBIT) DtCdaBal, " & _
               "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN DEBIT ELSE 0 END + " & _
               "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
               "THEN 0 ELSE DEBIT END) DtCdlbal " & _
             "From Gledg, Codes, " & _
                   "(SELECT DISTINCT SUBSTRING(CODE,1," & CodeLen & ") SUBCODE FROM Gledg " & _
                   " WHERE SUBSTRING(CODE,1," & CodeLen & ")='" & TCode & "') G " & _
                   "WHERE Type In (Select CFType From CFTypes) And VDATE Between " & _
                   "'" & Format(SDate, "dd-MMM-yyyy") & "' And " & _
                   "'" & Format(TmpStr, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & " ) = G.SUBCODE " & _
                   "AND Codes.CODE = G.SUBCODE+'" & Fillit("", "0", aCodeL - CodeLen) & "'" & _
                   " GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ") "

        Else

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description,SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
                   "SUM(DEBIT) DtCdaBal, " & _
                   "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN DEBIT ELSE 0 END + " & _
                   "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
                   "THEN 0 ELSE DEBIT END) DtCdlbal " & _
               "FROM Gledg, Codes " & _
                   "WHERE Type In (Select CFType From CFTypes) And VDATE Between " & _
                   "'" & Format(SDate, "dd-MMM-yyyy") & "' And " & _
                   "'" & Format(TmpStr, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & ") = '" & TCode & "' " & _
                   "AND Codes.CODE = '" & TCode & Fillit("", "0", aCodeL - CodeLen) & "' " & _
                   "GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ")"

        End If
        GetSqlDF = mSQL

    End Function

    Private Function GetSqlRF(ByVal PrDate As Date, ByVal TDate As Date, ByVal TCode As String, ByVal CodeLen As Integer, ByVal PrtNotes As Boolean) As String
        Dim mSQL As String
        Dim PrStr As String
        Dim TmpStr As String

        PrStr = GetCmpDate(PrDate)
        TmpStr = Format(TDate, "dd-MMM-yyyy")
        If PrtNotes Then

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description, SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
               "SUM(CREDIT) DtCdaBal, " & _
               "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN CREDIT ELSE 0 END + " & _
               "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
               "THEN 0 ELSE CREDIT END) DtCdlbal " & _
             "From Gledg, Codes, " & _
                   "(SELECT DISTINCT SUBSTRING(CODE,1," & CodeLen & ") SUBCODE FROM Gledg " & _
                   " WHERE SUBSTRING(CODE,1," & CodeLen & ")='" & TCode & "') G " & _
                   "WHERE Type In (Select CFType From CFTypes) And VDATE Between " & _
                   "'" & Format(SDate, "dd-MMM-yyyy") & "' And " & _
                   "'" & Format(TmpStr, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & " ) = G.SUBCODE " & _
                   "AND Codes.CODE = G.SUBCODE+'" & Fillit("", "0", aCodeL - CodeLen) & "'" & _
                   " GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ") "

        Else

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description,SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
                   "SUM(CREDIT) DtCdaBal, " & _
                   "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN CREDIT ELSE 0 END + " & _
                   "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
                   "THEN 0 ELSE CREDIT END) DtCdlbal " & _
               "FROM Gledg, Codes " & _
                   "WHERE Type In (Select CFType From CFTypes) And VDATE Between " & _
                   "'" & Format(SDate, "dd-MMM-yyyy") & "' And " & _
                   "'" & Format(TmpStr, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & ") = '" & TCode & "' " & _
                   "AND Codes.CODE = '" & TCode & Fillit("", "0", aCodeL - CodeLen) & "' " & _
                   "GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ")"

        End If
        GetSqlRF = mSQL

    End Function

    Private Function GetSqlC(ByVal PrDate As Date, ByVal TDate As Date, ByVal TCode As String, ByVal CodeLen As Integer, ByVal PrtNotes As Boolean) As String
        Dim mSQL As String
        Dim PrStr As String
        Dim TmpStr As String

        PrStr = GetCmpDate(PrDate)
        TmpStr = Format(TDate, "dd-MMM-yyyy")
        If PrtNotes Then

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description, SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
               "SUM(DEBIT-CREDIT) DtCdaBal, " & _
               "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN DEBIT-CREDIT ELSE 0 END + " & _
               "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
               "THEN 0 ELSE DEBIT-CREDIT END) DtCdlbal " & _
             "From Gledg, Codes, " & _
                   "(SELECT DISTINCT SUBSTRING(CODE,1," & CodeLen & ") SUBCODE FROM Gledg " & _
                   " WHERE SUBSTRING(CODE,1," & CodeLen & ")='" & TCode & "') G " & _
               "WHERE VDATE <= '" & Format(TmpStr, "dd-MMM-yyyy") & "' " & _
               "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
               "AND GLEDG.POSTED = 'Y' " & _
               "AND SUBSTRING(Gledg.CODE,1," & CodeLen & " ) = G.SUBCODE " & _
               "AND Codes.CODE = G.SUBCODE+'" & Fillit("", "0", aCodeL - CodeLen) & "'" & _
               " GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ") "

        Else

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description,SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
                   "SUM(DEBIT-CREDIT) DtCdaBal, " & _
                   "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN DEBIT-CREDIT ELSE 0 END + " & _
                   "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
                   "THEN 0 ELSE DEBIT-CREDIT END) DtCdlbal " & _
               "FROM Gledg, Codes " & _
                   "WHERE VDATE <= '" & Format(TmpStr, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & ") = '" & TCode & "' " & _
                   "AND Codes.CODE = '" & TCode & Fillit("", "0", aCodeL - CodeLen) & "' " & _
                   "GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ")"

        End If
        GetSqlC = mSQL

    End Function

    Private Function GetSqlA(ByVal PrDate As Date, ByVal TDate As Date, ByVal TCode As String, ByVal CodeLen As Integer, ByVal PrtNotes As Boolean) As String
        Dim mSQL As String
        Dim PrStr As String
        Dim TmpStr As String

        PrStr = GetCmpDate(PrDate)
        TmpStr = Format(TDate, "dd-MMM-yyyy")
        If PrtNotes Then

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description, SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
               "SUM(DEBIT-CREDIT) DtCdaBal, " & _
               "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN DEBIT-CREDIT ELSE 0 END + " & _
               "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
               "THEN 0 ELSE DEBIT-CREDIT END) DtCdlbal " & _
             "From Gledg, Codes, " & _
                   "(SELECT DISTINCT SUBSTRING(CODE,1," & CodeLen & ") SUBCODE FROM Gledg " & _
                   " WHERE SUBSTRING(CODE,1," & CodeLen & ")='" & TCode & "') G " & _
                   "WHERE VDATE Between '" & Format(SDate, "dd-MMM-yyyy") & "' And " & _
                   "'" & Format(TmpStr, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & " ) = G.SUBCODE " & _
                   "AND Codes.CODE = G.SUBCODE+'" & Fillit("", "0", aCodeL - CodeLen) & "'" & _
                   " GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ") "

        Else

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description,SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
                   "SUM(DEBIT-CREDIT) DtCdaBal, " & _
                   "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN DEBIT-CREDIT ELSE 0 END + " & _
                   "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
                   "THEN 0 ELSE DEBIT-CREDIT END) DtCdlbal " & _
               "FROM Gledg, Codes " & _
                   "WHERE VDATE Between '" & Format(SDate, "dd-MMM-yyyy") & "' And " & _
                   "'" & Format(TmpStr, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & ") = '" & TCode & "' " & _
                   "AND Codes.CODE = '" & TCode & Fillit("", "0", aCodeL - CodeLen) & "' " & _
                   "GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ")"

        End If
        GetSqlA = mSQL

    End Function

    Private Function GetSqlO(ByVal PrDate As Date, ByVal TDate As Date, ByVal TCode As String, ByVal CodeLen As Integer, ByVal PrtNotes As Boolean) As String
        Dim mSQL As String
        Dim PrStr As String
        Dim TmpStr As String

        PrStr = GetCmpDate(PrDate)
        TmpStr = Format(TDate, "dd-MMM-yyyy")
        If PrtNotes Then

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description, SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
               "SUM(DEBIT-CREDIT) DtCdaBal, " & _
               "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN DEBIT-CREDIT ELSE 0 END + " & _
               "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
               "THEN 0 ELSE DEBIT-CREDIT END) DtCdlbal " & _
             "From Gledg, Codes, " & _
                   "(SELECT DISTINCT SUBSTRING(CODE,1," & CodeLen & ") SUBCODE FROM Gledg " & _
                   " WHERE SUBSTRING(CODE,1," & CodeLen & ")='" & TCode & "') G " & _
                   "WHERE VDATE < '" & Format(SDate, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & " ) = G.SUBCODE " & _
                   "AND Codes.CODE = G.SUBCODE+'" & Fillit("", "0", aCodeL - CodeLen) & "'" & _
                   " GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ") "

        Else

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description,SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
                   "SUM(DEBIT-CREDIT) DtCdaBal, " & _
                   "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN DEBIT-CREDIT ELSE 0 END + " & _
                   "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
                   "THEN 0 ELSE DEBIT-CREDIT END) DtCdlbal " & _
               "FROM Gledg, Codes " & _
                   "WHERE VDATE < '" & Format(SDate, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & ") = '" & TCode & "' " & _
                   "AND Codes.CODE = '" & TCode & Fillit("", "0", aCodeL - CodeLen) & "' " & _
                   "GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ")"

        End If
        GetSqlO = mSQL

    End Function

    Private Function GetSqlD(ByVal PrDate As Date, ByVal TDate As Date, ByVal TCode As String, ByVal CodeLen As Integer, ByVal PrtNotes As Boolean) As String
        Dim mSQL As String
        Dim PrStr As String
        Dim TmpStr As String

        PrStr = GetCmpDate(PrDate)
        TmpStr = Format(TDate, "dd-MMM-yyyy")
        If PrtNotes Then

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description, SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
               "SUM(DEBIT) DtCdaBal, " & _
               "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN DEBIT ELSE 0 END + " & _
               "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
               "THEN 0 ELSE DEBIT END) DtCdlbal " & _
             "From Gledg, Codes, " & _
                   "(SELECT DISTINCT SUBSTRING(CODE,1," & CodeLen & ") SUBCODE FROM Gledg " & _
                   " WHERE SUBSTRING(CODE,1," & CodeLen & ")='" & TCode & "') G " & _
                   "WHERE VDATE Between '" & Format(SDate, "dd-MMM-yyyy") & "' And " & _
                   "'" & Format(TmpStr, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & " ) = G.SUBCODE " & _
                   "AND Codes.CODE = G.SUBCODE+'" & Fillit("", "0", aCodeL - CodeLen) & "'" & _
                   " GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ") "

        Else

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description,SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
                   "SUM(DEBIT) DtCdaBal, " & _
                   "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN DEBIT ELSE 0 END + " & _
                   "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
                   "THEN 0 ELSE DEBIT END) DtCdlbal " & _
               "FROM Gledg, Codes " & _
                   "WHERE VDATE Between '" & Format(SDate, "dd-MMM-yyyy") & "' And " & _
                   "'" & Format(TmpStr, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & ") = '" & TCode & "' " & _
                   "AND Codes.CODE = '" & TCode & Fillit("", "0", aCodeL - CodeLen) & "' " & _
                   "GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ")"

        End If
        GetSqlD = mSQL

    End Function

    Private Function GetSqlR(ByVal PrDate As Date, ByVal TDate As Date, ByVal TCode As String, ByVal CodeLen As Integer, ByVal PrtNotes As Boolean) As String
        Dim mSQL As String
        Dim PrStr As String
        Dim TmpStr As String

        PrStr = GetCmpDate(PrDate)
        TmpStr = Format(TDate, "dd-MMM-yyyy")
        If PrtNotes Then

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description, SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
               "SUM(CREDIT) DtCdaBal, " & _
               "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN CREDIT ELSE 0 END + " & _
               "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
               "THEN 0 ELSE CREDIT END) DtCdlbal " & _
             "From Gledg, Codes, " & _
                   "(SELECT DISTINCT SUBSTRING(CODE,1," & CodeLen & ") SUBCODE FROM Gledg " & _
                   " WHERE SUBSTRING(CODE,1," & CodeLen & ")='" & TCode & "') G " & _
                   "WHERE VDATE Between '" & Format(SDate, "dd-MMM-yyyy") & "' And " & _
                   "'" & Format(TmpStr, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & " ) = G.SUBCODE " & _
                   "AND Codes.CODE = G.SUBCODE+'" & Fillit("", "0", aCodeL - CodeLen) & "'" & _
                   " GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ") "

        Else

            mSQL = "SELECT MAX(Codes.DESCRIPTION) AS Description,SUBSTRING(Gledg.CODE,1," & CodeLen & ") AS Code, " & _
                   "SUM(CREDIT) DtCdaBal, " & _
                   "SUM(CASE WHEN DATEDIFF(DD,Vdate,'" & PrStr & "') = 0 THEN CREDIT ELSE 0 END + " & _
                   "    CASE WHEN ABS(DATEDIFF(DD,Vdate,'" & PrStr & "')) = DATEDIFF(DD,Vdate,'" & PrStr & "') " & _
                   "THEN 0 ELSE CREDIT END) DtCdlbal " & _
               "FROM Gledg, Codes " & _
                   "WHERE VDATE Between '" & Format(SDate, "dd-MMM-yyyy") & "' And " & _
                   "'" & Format(TmpStr, "dd-MMM-yyyy") & "' " & _
                   "AND GLEDG.BRCODE = '" & SysBrCode & "' " & _
                   "AND GLEDG.POSTED = 'Y' " & _
                   "AND SUBSTRING(Gledg.CODE,1," & CodeLen & ") = '" & TCode & "' " & _
                   "AND Codes.CODE = '" & TCode & Fillit("", "0", aCodeL - CodeLen) & "' " & _
                   "GROUP BY SUBSTRING(Gledg.CODE,1," & CodeLen & ")"

        End If
        GetSqlR = mSQL

    End Function

    Private Function BudgetSql(ByVal TCode As String, ByVal CodeLen As Integer) As String
        Dim mSQL As String
        'Dim PrStr As String
        'Dim TmpStr As String

        mSQL = "SELECT SUBSTRING(CODES.CODE,1," & CodeLen & ")" & " CODE, " & _
                "MAX(CODES.DESCRIPTION) DESCRIPTION, " & _
                "SUM(BUDGET.BUDGET1) BUDGET1, SUM(BUDGET.BUDGET2) BUDGET2, " & _
                "SUM(BUDGET.BUDGET3) BUDGET3, SUM(BUDGET.BUDGET4) BUDGET4,  " & _
                "SUM(BUDGET.BUDGET5) BUDGET5, SUM(BUDGET.BUDGET6) BUDGET6,  " & _
                "SUM(BUDGET.BUDGET7) BUDGET7, SUM(BUDGET.BUDGET8) BUDGET8,  " & _
                "SUM(BUDGET.BUDGET9) BUDGET9, SUM(BUDGET.BUDGET10) BUDGET10,  " & _
                "SUM(BUDGET.BUDGET11) BUDGET11, SUM(BUDGET.BUDGET12) BUDGET12  " & _
                "FROM BUDGET INNER JOIN CODES ON BUDGET.CODE = CODES.CODE " & _
                "WHERE CODES.CODE = '" & TCode & Fillit("", "0", aCodeL - CodeLen) & "' " & _
                " AND LEN(RTRIM(LTRIM(CODES.CODE))) = " & aCodeL & _
                " AND BUDGET.BRCODE= '" & SysBrCode & "' " & _
                "GROUP BY SUBSTRING(CODES.CODE,1," & CodeLen & ")"

        BudgetSql = mSQL

    End Function
    ''**************** PROCEDURES USED IN INCOME STATEMENT *****************
    ''**
    Private Sub SumBal(ByVal mpLSTBAL, ByVal mpCURBAL, ByVal mpACCBAL)
        Select Case OPERATORS
            Case "+"
                TLBal = TLBal + mpLSTBAL
                TCBal = TCBal + mpCURBAL
                TABal = TLBal + TCBal
            Case "-"
                TLBal = TLBal - mpLSTBAL
                TCBal = TCBal - mpCURBAL
                TABal = TLBal + TCBal
            Case "*"
                If mpLSTBAL <> 0 Then
                    TLBal = TLBal * mpLSTBAL
                End If
                If mpCURBAL <> 0 Then
                    TCBal = TCBal * mpCURBAL
                End If
                TABal = TLBal + TCBal
            Case "/"
                If mpLSTBAL <> 0 Then
                    TLBal = TLBal / mpLSTBAL
                End If
                If mpCURBAL <> 0 Then
                    TCBal = TCBal / mpCURBAL
                End If
                TABal = TLBal + TCBal
            Case "^"
                If mpLSTBAL <> 0 Then
                    TLBal = TLBal ^ mpLSTBAL
                End If
                If mpCURBAL <> 0 Then
                    TCBal = TCBal ^ mpCURBAL
                End If
                TABal = TLBal + TCBal
        End Select
    End Sub
    Private Sub SumBudg(ByVal mpLSTBUDG, ByVal mpCURBUDG, ByVal mpACCBUDG)

        Select Case OPERATORS
            Case "+"
                TLBudg = TLBudg + mpLSTBUDG
                TCBudg = TCBudg + mpCURBUDG
                TABudg = TABudg + mpACCBUDG
            Case "-"
                TLBudg = TLBudg - mpLSTBUDG
                TCBudg = TCBudg - mpCURBUDG
                TABudg = TABudg - mpACCBUDG
            Case "*"
                If mpLSTBUDG <> 0 Then
                    TLBudg = TLBudg * mpLSTBUDG
                End If
                If mpCURBUDG <> 0 Then
                    TCBudg = TCBudg * mpCURBUDG
                End If
                If mpACCBUDG <> 0 Then
                    TABudg = TABudg * mpACCBUDG
                End If
            Case "/"
                If mpLSTBUDG <> 0 Then
                    TLBudg = TLBudg / mpLSTBUDG
                End If
                If mpCURBUDG <> 0 Then
                    TCBudg = TCBudg / mpCURBUDG
                End If
                If mpACCBUDG <> 0 Then
                    TABudg = TABudg / mpACCBUDG
                End If
            Case "^"
                If mpLSTBUDG <> 0 Then
                    TLBudg = TLBudg ^ mpLSTBUDG
                End If
                If mpCURBUDG <> 0 Then
                    TCBudg = TCBudg ^ mpCURBUDG
                End If
                If mpACCBUDG Then
                    TABudg = TABudg ^ mpACCBUDG
                End If
        End Select
    End Sub

End Module
