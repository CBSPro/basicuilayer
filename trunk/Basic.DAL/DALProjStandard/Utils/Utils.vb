Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Basic.Constants.ProjConst
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles

Public Class Utils



    Public Shared Function GetFldValue(ByVal strSQl As String, ByVal Field As String) As Object

        Dim Objcmd As SqlClient.SqlCommand
        Dim obj As Object
        Erase obj
        Try
            strConec = "Data Source=" & SysServer & ";Initial Catalog=" & SysDataBase & ";User ID= " & SysUser & " ;password=" & SysPassword & ";"
            sqlCon = New SqlConnection(strConec)
            If sqlCon.State = 1 Then
                sqlCon.Close()
            End If
            sqlCon.Open()
            Objcmd = New SqlClient.SqlCommand(strSQl, sqlCon)
            Objcmd.CommandType = CommandType.Text
            obj = Objcmd.ExecuteScalar
            If obj Is Nothing Then
                obj = ""
            ElseIf obj.Equals(System.DBNull.Value) Then
                obj = ""
            Else : Return obj
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information)
        Finally
            sqlCon.Close()

        End Try
        Return obj
    End Function

    Public Shared Function Found(ByVal strSQl As String) As Boolean
        Dim Objcmd As SqlClient.SqlCommand
        Dim obj As Object
        Erase obj
        Try
            strConec = "Data Source=" & SysServer & ";Initial Catalog=" & SysDataBase & ";User ID= " & SysUser & " ;password=" & SysPassword & ";"
            sqlCon = New SqlConnection(strConec)
            If sqlCon.State = 1 Then
                sqlCon.Close()
            End If
            sqlCon.Open()
            Objcmd = New SqlClient.SqlCommand(strSQl, sqlCon)
            Objcmd.CommandType = CommandType.Text
            obj = Objcmd.ExecuteScalar
            If obj Is Nothing Then
                Found = False
            ElseIf obj.Equals(System.DBNull.Value) Then
                Found = False
            Else
                'Return obj
                Found = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information)
        Finally
            sqlCon.Close()
        End Try
        'Return obj
    End Function

    Public Shared Function Lookup(ByVal StrSql As String) As DataSet
        Dim Ds As New DataSet

        Try
            strConec = "Data Source=" & SysServer & ";Initial Catalog=" & SysDataBase & ";User ID=" & SysUser & ";password=" & SysPassword & ";"
            sqlCon = New SqlConnection(strConec)
            If sqlCon.State = ConnectionState.Closed Then
                sqlCon.Open()
            End If
            DAdp = New SqlDataAdapter(StrSql, sqlCon)
            DAdp.Fill(Ds)


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information)
        Finally
            If sqlCon.State = ConnectionState.Open Then
                sqlCon.Close()
            End If


        End Try
        Return Ds
    End Function
    'FUNCTION TO REMOVE HYPHEN FROM Department CODE
    Public Function aFarmCode2Str(ByVal mTmpCode As Object) As String
        Dim mCode As String
        mCode = mTmpCode
        aFarmCode2Str = Mid(mCode, aP1P, aP1L) & Mid(mCode, aP2P + 1, aP2L)
    End Function
    'FUNCTION TO ADD HYPHEN FROM Department CODE
    Public Function aStr2FarmCode(ByVal mTmpCode As String) As String
        'Call SetAccParam()
        Dim mStr As String
        mStr = mTmpCode & Space(aCodeL - Len(Trim(mTmpCode)))
        aStr2FarmCode = Mid(mStr, aP1P, aP1L) & "-" & Mid(mStr, aP2P, aP2L)
    End Function
    'FUNCTION TO REMOVE HYPHEN FROM ACCOUNT CODE
    Public Function aCode2Str(ByVal mTmpCode As Object) As String
        'Call SetAccParam()
        Dim mCode As String

        mCode = mTmpCode
        aCode2Str = Mid(mCode, aP1P, aP1L) & Mid(mCode, aP2P + 1, aP2L) & Mid(mCode, aP3P + 2, aP3L) & Mid(mCode, aP4P + 3, aP4L)
    End Function
    'FUNCTION TO ADD HYPHEN FROM ACCOUNT CODE
    Public Function aStr2Code(ByVal mTmpCode As String) As String
        'Call SetAccParam()
        Dim mStr As String

        mStr = mTmpCode & Space(aCodeL - Len(Trim(mTmpCode)))
        aStr2Code = Mid(mStr, aP1P, aP1L) & "-" & Mid(mStr, aP2P, aP2L) & "-" & Mid(mStr, aP3P, aP3L)

    End Function
    'FUNCTION TO REMOVE HYPHEN FROM ACCOUNT CODE
    Public Function aLndCode2Str(ByVal mTmpCode As Object) As String
        'Call SetAccParam()
        Dim mCode As String

        mCode = Trim(mTmpCode)
        aLndCode2Str = Mid(mCode, aP1P, aP1L) & Mid(mCode, aP2P + 1, aP2L) & Mid(mCode, aP4P, aP1L) & Mid(mCode, aP4P + 3, aP1L)
    End Function
    'FUNCTION TO ADD HYPHEN FROM ACCOUNT CODE
    Public Function aStr2LndCode(ByVal mTmpCode As String) As String
        'Call SetAccParam()
        Dim mStr As String
        'Dim alndCodeL As String
        mStr = Trim(mTmpCode)
        'alndCodeL = 11
        'mStr = mTmpCode & Space(alndCodeL - Len(Trim(mTmpCode)))
        'aStr2LndCode = Mid(mStr, aP1P, aP1L) & "-" & Mid(mStr, aP2P, aP2L) & "-" & Mid(mStr, aP3P, aP3L)
        aStr2LndCode = Mid(mStr, aP1P, aP1L) & "-" & Mid(mStr, aP2P, aP2L) & "-" & Mid(mStr, aP3P, aP1L) & "-" & Mid(mStr, aP4P, aP1L)
    End Function

    Public Shared Function Fillit(ByVal Fillstr As String, ByVal Fillchar As String, ByVal Fillnos As Integer) As String
        Dim i As Integer
        For i = 1 To Fillnos
            Fillstr = Fillstr & Fillchar
        Next
        Fillit = Fillstr

    End Function

    Public Function CmpDate(ByVal mDate As Date) As String
        CmpDate = Format(mDate, "yyyy-MM-dd")
    End Function

    Public Function GetDate(ByVal DisplayDate As String, ByVal DateFormat As String) As String

        If (DisplayDate <> "" And Not DisplayDate Is Nothing) Then
            'Dim Dtemp As Integer
            Dim Day, month, year, dateArr

            'Dim ShrtMonth As String
            Dim ShortMonths As String = "Jan,Feb,Mar,Apr,May,Jun,July,Aug,Sept,Oct,Nov,Dec"
            Dim LongMonths As String = "January,February,March,April,May,June,July,August,September,October,November,December"
            Dim x As Array = ShortMonths.Split(",")
            Dim y As Array = LongMonths.Split(",")
            If (DateFormat = "MM/dd/yyyy") Then
                dateArr = DisplayDate.Split("/")
                month = dateArr(0)
                Day = dateArr(1)
                year = dateArr(2)
            ElseIf (DateFormat = "dd/MM/yyyy") Then
                dateArr = DisplayDate.Split("/")
                month = dateArr(1)
                Day = dateArr(0)
                year = dateArr(2)
            ElseIf (DateFormat = "MM-dd-yyyy") Then
                dateArr = DisplayDate.Split("-")
                month = dateArr(0)
                Day = dateArr(1)
                year = dateArr(2)
            ElseIf (DateFormat = "dd MM yyyy") Then
                dateArr = DisplayDate.Split(" ")
                month = dateArr(1)
                Day = dateArr(0)
                year = dateArr(2)
            ElseIf (DateFormat = "yyyy-MM-dd") Then
                'ElseIf (DateFormat = "dd-MMM-yyyy") Then
                dateArr = DisplayDate.Split("-")
                month = dateArr(1)
                Day = dateArr(0)
                year = dateArr(2)
            ElseIf (DateFormat = "MM dd yyyy") Then
                dateArr = DisplayDate.Split(" ")
                month = dateArr(0)
                Day = dateArr(1)
                year = dateArr(2)
            ElseIf Mid(DisplayDate, 3, 1) = "/" Then
                'ElseIf (DateFormat = "dd/MM/yyyy") Then
                dateArr = DisplayDate.Split("/")
                'month = GetMonth(dateArr(1))
                Day = dateArr(0)
                year = dateArr(2)
                month = dateArr(1)
                'month = Mid(dateArr, 1, 2)
                'Day = Mid(dateArr, 4, 2)
                'year = Mid(dateArr, 7, 4)
            ElseIf (DateFormat = "dd-MMM-yyyy" Or DateFormat = "dd-MMMM-yyyy" Or DateFormat = "MMM-dd-yyyy" Or DateFormat = "MMMM-dd-yyyy") Then
                dateArr = DisplayDate.Split("-")
                If (IsNumeric(dateArr(0))) Then
                    month = GetMonth(dateArr(1))
                    Day = dateArr(0)
                    'month = dateArr(1)
                Else
                    month = GetMonth(dateArr(0))
                    Day = dateArr(1)
                End If
                year = dateArr(2)

            ElseIf Mid(DisplayDate, 3, 1) = "-" Then
                'ElseIf (DateFormat = "dd-MMM-yyyy") Then
                dateArr = DisplayDate.Split("-")
                month = dateArr(1)
                Day = dateArr(0)
                year = dateArr(2)

            ElseIf (DateFormat = "dd/MMM/yyyy" Or DateFormat = "dd/MMMM/yyyy" Or DateFormat = "MMM/dd/yyyy" Or DateFormat = "MMMM/dd/yyyy") Then
                dateArr = DisplayDate.Split("/")
                If (IsNumeric(dateArr(0))) Then
                    month = GetMonth(dateArr(1))
                    Day = dateArr(0)
                    'month = dateArr(1)
                Else
                    month = GetMonth(dateArr(0))
                    Day = dateArr(1)
                End If
                year = dateArr(2)
            ElseIf (DateFormat = "dd MMM yyyy" Or DateFormat = "dd MMMM yyyy" Or DateFormat = "MMM dd yyyy" Or DateFormat = "MMMM dd yyyy") Then
                dateArr = DisplayDate.Split(" ")
                If (IsNumeric(dateArr(0))) Then
                    month = GetMonth(dateArr(1))
                    Day = dateArr(0)
                Else
                    month = GetMonth(dateArr(0))
                    Day = dateArr(1)
                End If
                year = dateArr(2)
            Else
                Return CType(DisplayDate, Date).ToString("MM/dd/yyyy")
            End If
            If ((Not month Is Nothing) And (Not Day Is Nothing) And (Not year Is Nothing)) Then
                'If (month <= 12 And Day <= 31 And IsNumeric(year) And year.length < 5) Then
                'Return CType((Day & "-" & month + "-" & year), String)
                Return CType((Day & "-" & month & "-" & year), String)
                'End If
            End If
        End If
        Return ""
    End Function
    Public Function GetMDate(ByVal DisplayDate As String, ByVal DateFormat As String) As String
        Dim Day, month, year, dateArr
        Dim ShortMonths As String = "Jan,Feb,Mar,Apr,May,Jun,July,Aug,Sept,Oct,Nov,Dec"
        Dim LongMonths As String = "January,February,March,April,May,June,July,August,September,October,November,December"
        Dim x As Array = ShortMonths.Split(",")
        Dim y As Array = LongMonths.Split(",")
        If (DateFormat = "dd/MMM/yyyy" Or DateFormat = "dd/MMMM/yyyy" Or DateFormat = "MMM/dd/yyyy" Or DateFormat = "MMMM/dd/yyyy") Then
            dateArr = DisplayDate.Split("/")
            If (IsNumeric(dateArr(0))) Then
                month = GetMonth(dateArr(1))
                Day = dateArr(0)
                'month = dateArr(1)
            Else
                month = GetMonth(dateArr(0))
                Day = dateArr(1)
            End If
            year = dateArr(2)
        ElseIf (DateFormat = "dd MMM yyyy" Or DateFormat = "dd MMMM yyyy" Or DateFormat = "MMM dd yyyy" Or DateFormat = "MMMM dd yyyy") Then
            dateArr = DisplayDate.Split(" ")
            If (IsNumeric(dateArr(0))) Then
                month = GetMonth(dateArr(1))
                Day = dateArr(0)
            Else
                month = GetMonth(dateArr(0))
                Day = dateArr(1)
            End If
            year = dateArr(2)
        Else
            Return CType(DisplayDate, Date).ToString("MM/dd/yyyy")
        End If
        'If ((Not Month() Is Nothing) And (Not Day Is Nothing) And (Not Year() Is Nothing)) Then
        '    'If (month <= 12 And Day <= 31 And IsNumeric(year) And year.length < 5) Then
        '    'Return CType((Day & "-" & month + "-" & year), String)
        '    Return CType((Year() & "-" & Month() & "-" & Day), String)
        '    'End If
        'End If
        Return CType((CType(year(), String) & "-" & CType(month(), String) & "-" & CType(Day, String)), String)
    End Function
    Public Function GetAlphaMonth(ByVal DisplayDate As String) As String
        Dim MonthArr() As String = {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}
        Dim count = 0
        Dim dateArr, Month, Day, Year
        If Mid(DisplayDate, 3, 1) = "/" Then
            'ElseIf (DateFormat = "dd/MM/yyyy") Then
            dateArr = DisplayDate.Split("/")
            Month = GetMonth(dateArr(1))
            Day = dateArr(0)
            Year = dateArr(2)
            Month = MonthArr(dateArr(1))
            'month = Mid(dateArr, 1, 2)
            'Day = Mid(dateArr, 4, 2)
            'year = Mid(dateArr, 7, 4)
            Return CType((Month & "-" & Day & "-" & Year), String)
        ElseIf Mid(DisplayDate, 3, 1) = "-" Then
            'ElseIf (DateFormat = "dd-MMM-yyyy") Then
            dateArr = DisplayDate.Split("-")
            'Month = dateArr(1)

            Day = dateArr(0)
            Year = dateArr(2)
            'Return CType((Month & "-" & Day & "-" & Year), String)
        End If


        'For count = 0 To MonthArr.Length - 1
        '    If (UCase(MonthArr(count)).IndexOf(UCase(Month)) >= 0) Then
        '        Return (count + 1)
        '        'Return MonthArr(Month)
        '    End If
        'Next

    End Function

    Public Function GetMonth(ByVal Month) As String
        Dim MonthArr() As String = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
        Dim count = 0

        'For count = 0 To MonthArr.Length - 1
        'If (UCase(MonthArr(count)).IndexOf(UCase(Month)) >= 0) Then
        '    Return (count + 1)
        Return MonthArr(Month - 1)
        'End If
        'Next

    End Function

    Public Shared Function PMonth(ByVal mDate As Date) As Date
        PMonth = CDate("01/" + Str(Month(mDate)) + "/" + Str(Year(mDate)))
    End Function

    Public Shared Function GetCmpDate(ByVal mDate As Date) As String
        GetCmpDate = UCase(Format(mDate, "dd-MMM-yyyy"))
    End Function

    Public Shared Function GetDataTable(ByVal strQuery As String) As DataTable
        Dim DTable As New DataTable
        Dim DA As SqlDataAdapter

        Try
            sqlCon = New SqlConnection(strConec)
            sqlCon.Open()
            DA = New SqlDataAdapter(strQuery, sqlCon)
            DA.Fill(DTable)
            Return DTable
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information)
        Finally
            sqlCon.Close()
        End Try
        Return DTable
    End Function
    'Public Function NumInput(ByVal KeyAscii As Integer) As Object
    '    'Select Case KeyAscii

    '    If KeyAscii = 13 Then
    '        SendKeys.Send("{Tab}")
    '    ElseIf KeyAscii = 8 Or KeyAscii = 46 Or (KeyAscii > 48 And KeyAscii < 57) Then
    '    Else
    '        KeyAscii = 0
    '        'Case 8, 46, 48 To 57
    '        'Case Else
    '        'KeyAscii = 0
    '        'End Select
    '    End If
    '    Return KeyAscii
    'End Function


    Public Sub NumInput(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
        '        e.Handled = True
        '    End If
        'End If
        Select Case Asc(e.KeyChar)
            Case 13
                SendKeys.Send("{Tab}")
            Case 8, 46, 48 To 57
            Case Else
                e.Handled = True
        End Select
    End Sub

    Public Sub AlphaInput(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        'If Asc(e.KeyChar) <> 8 Then
        '    If (Asc(e.KeyChar) < 65 Or Asc(e.KeyChar) > 90) Or (Asc(e.KeyChar) < 97 Or Asc(e.KeyChar) > 122) Or Asc(e.KeyChar) = 32 Then
        '        e.Handled = True
        '    End If
        'End If
        Select Case Asc(e.KeyChar)
            Case 13
                SendKeys.Send("{Tab}")
            Case 8, 40, 41, 46, 32, 65 To 91, 97 To 123
            Case Else
                e.Handled = True
        End Select
    End Sub

    Public Function GetFADocNo(ByVal mType As String, ByVal mDate As Date) As String
        'Dim mDa As SqlDataAdapter
        'Dim mDt As New DataTable
        Dim mSQL As String
        Dim mNewNo As String
        Dim mMonth As String
        Dim mYear As String

        mYear = Format(Year(mDate), "0000")
        mMonth = Format(Month(mDate), "00")
        ''
        mSQL = "SELECT MAX(FarNO) FarNo FROM FarHead WHERE " & _
                "LEFT(FarNO,4) = '" & mYear & "' AND " & _
                "SUBSTRING(FarNO,5,2) = '" & mMonth & "' AND " & _
                "NoteType = '" & mType & "'"

        'mDa = New SqlDataAdapter(mSQL, sqlCon)
        'mDa.Fill(mDt)
        mSQL = GetFldValue(mSQL, "FarNO")
        If mSQL <> "" Then
            'If Not IsDBNull(mDt.Rows(0).Item(0)) Then
            ''mNewNo = Format(Val(mDt.Rows(0).Item(0)) + 1, "0000000000")
            mNewNo = Format(Val(mSQL) + 1, "0000000000")
        Else
            mNewNo = mYear & mMonth & Format(1, "0000")
        End If
        GetFADocNo = mNewNo
    End Function
    Public Function GetNewCode(ByVal FldName As String, ByVal FldLen As Integer, ByVal FldTable As String) As String
        mSQL = "SELECT Max(" & FldName & ") As NewCode FROM " & FldTable
        mSQL = Right(GetFldValue(mSQL, "NewCode"), 4)
        If mSQL = "" Then
            mSQL = Format(1, Fillit("", "0", FldLen))
        Else
            mSQL = Format(Val(mSQL) + 1, Fillit("", "0", FldLen))
        End If
        GetNewCode = mSQL
    End Function


    Public Function GetVno(ByVal Brcode As String, ByVal mDate As String, ByVal mType As String) As String
        Dim Objcmd As SqlClient.SqlCommand
        Dim obj As Object
        Dim mYear, mMonth As String
        Dim mSql As String
        Dim Found As Boolean

        mYear = Mid(Trim(Str(Year(mDate))), 1, 4)
        mMonth = String.Format(Trim(Str(Month(mDate))), "00")
        mSql = "Select Top 1 Vno from Glhead where BrCode = '" + Brcode + "' and Type = '" + mType + "' and Year(Vdate) ='" + mYear + "' and Month(Vdate) ='" + mMonth + "'"
        Try
            strConec = "Data Source=" & SysServer & ";Initial Catalog=" & SysDataBase & ";User ID=" & SysUser & ";password=" & SysPassword & ";"
            sqlCon = New SqlConnection(strConec)
            If sqlCon.State = 1 Then
                sqlCon.Close()
            End If
            sqlCon.Open()
            Objcmd = New SqlClient.SqlCommand(mSql, sqlCon)
            Objcmd.CommandType = CommandType.Text
            obj = Objcmd.ExecuteScalar
            If obj Is Nothing Then
                Found = False
                GetVno = mYear + mMonth + String.Format(1, "00000")
            ElseIf obj.Equals(System.DBNull.Value) Then
                Found = False
                GetVno = mYear + mMonth + String.Format(1, "00000")
            Else
                'Return obj
                GetVno = Convert.ToString(obj) + 1
                Found = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information)
        Finally
            sqlCon.Close()
        End Try
    End Function

    Public Shared Function ExecQuery(ByVal strSQl As String)
        Dim Objcmd As SqlClient.SqlCommand
        Dim obj As Object
        Erase obj
        Try
            strConec = "Data Source=" & SysServer & ";Initial Catalog=" & SysDataBase & ";User ID=" & SysUser & ";password=" & SysPassword & ";"
            sqlCon = New SqlConnection(strConec)
            If sqlCon.State = 1 Then
                sqlCon.Close()
            End If
            sqlCon.Open()
            Objcmd = New SqlClient.SqlCommand(strSQl, sqlCon)
            Objcmd.CommandType = CommandType.Text
            Objcmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information)
        Finally
            sqlCon.Close()
        End Try
        'Return obj
    End Function

    Public Shared Sub ShowMenus(ByVal status As Boolean)
        mfrmBP = status
        mfrmBR = status
        mfrmCP = status
        mfrmCR = status
        mfrmJV = status

        'Admin Menu
        mUMaintenence = status
        mCPass = status
        mUPass = status
        mBckDB = status
        mEPostVoch = status
        mSysParm = status

        'Reports Menu
        'On Screen View Menu
        mGLScr = status
        mCStatus = status
        mBBookScr = status
        mCBookScr = status
        mAcCStatus = status
        mBBook = status
        mCBook = status
        mTBal = status
        mTBalCross = status
        mATrial = status
        mSubLedg = status
        mAAnalysis = status
        mPDefRep = status
        mPDefRefCross = status
        mPPstVouchers = status

        'Basic Data
        mBranch = status
        mAccMaintanance = status
        mBudget = status
        mDefReport = status

        mfrmBranch = status
        mfrmAcc = status
        mfrmBudget = status
        mfrmRepDef = status
        mfrmBrowseGL = status
        mfrmACStatus = status
        mfrmBnkBook = status
        mfrmCashBook = status
    End Sub









End Class
