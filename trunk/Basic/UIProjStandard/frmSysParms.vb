Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports Basic.DAL
Imports Basic.Constants.ProjConst
Imports Basic.DAL.Utils


Public Class frmSysParms
    Inherits System.Windows.Forms.Form

    Dim objSysParms As New cSysParms

    Dim PODS As New DataSet
    Dim DtDet As New DataTable
    Dim DaDet As SqlDataAdapter
    Dim objRow As Data.DataRow
    Dim ObjFind As Grid_Help
    Dim tableName As String = "Parms"
    Dim strFind As String
    Dim Flag As Boolean
    Dim AddMode As Boolean
    Dim EditMode As Boolean
    Dim KeyReturn As Boolean
    Dim ColBack As Boolean
    Dim SelRow As Boolean
    Dim dtSysParms As New DataTable
    Dim rowNum As Integer
    Dim mi As Integer

    Private Sub frmSysParms_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
    End Sub

    Private Sub frmSysParms_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        mSysParm = False
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub frmSysParms_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.MdiParent = New frmMdi
        Me.Top = 0
        Me.Left = 0
        'Me.WindowState = FormWindowState.Maximized
        Flag = True
        Call SetEntryMode()
        Call LockGrid(True)
        btnSave.Enabled = False
        btnCancel.Enabled = False
        dtSysParms = objSysParms.LoadValues()
        rowNum = dtSysParms.Rows.Count - 1
        Call LoadValues()
        btnExit.Focus()

    End Sub

    Private Sub LoadValues()
        Dim mi As Integer

        rowNum = 0
        For mi = 1 To dtSysParms.Rows.Count
            mStr = dtSysParms.Rows(rowNum).Item("FldName").ToString()
            vspGrid.SetText(1, mi, mStr)
            mStr = dtSysParms.Rows(rowNum).Item("FldValue").ToString()
            vspGrid.SetText(2, mi, mStr)
            mStr = dtSysParms.Rows(rowNum).Item("FldDescription").ToString()
            vspGrid.SetText(3, mi, mStr)
            If mi < dtSysParms.Rows.Count Then
                vspGrid.MaxRows = vspGrid.MaxRows + 1
                rowNum = rowNum + 1
            End If
        Next
    End Sub

    Sub SetEntryMode()
        btnAdd.Enabled = Flag
        btnEdit.Enabled = Flag
        btnView.Enabled = Flag
        btnDelete.Enabled = Flag
        btnPost.Enabled = Flag
        btnPrint.Enabled = Flag
        btnSave.Enabled = Flag
        btnCancel.Enabled = Flag
        btnFind.Enabled = Flag
        btnRefresh.Enabled = Flag
        btnTop.Enabled = Flag
        btnPrevious.Enabled = Flag
        btnNext.Enabled = Flag
        btnBottom.Enabled = Flag
        btnExit.Enabled = Flag
    End Sub

    Private Sub btnAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Dim StrTemp As Object
        AddMode = True
        EditMode = False
        Flag = False
        Call SetEntryMode()
        Call LockGrid(False)
        'If vspGrid.MaxRows > 1 Then
        Erase StrTemp
        vspGrid.GetText(1, vspGrid.MaxRows, StrTemp)
        If Not mTemp = " " Then
            vspGrid.MaxRows = vspGrid.MaxRows + 1
            'SetCol(vspGrid.MaxRows, 1)
        End If
        Call SetCol(vspGrid.MaxRows, 1)
        'End If
        btnSave.Enabled = True
        btnCancel.Enabled = True
        lblCompany.Text = "Recorded On " & SySDate
        lblToolTip.Text = "Add New Record"
        lblBy.Text = "Recorded By : " & SysUserID
        'lblBy.Text = "Recorded By : " & SysUserID
    End Sub

    Private Sub SetCol(ByVal mRow As Integer, ByVal mCol As Integer)
        vspGrid.Row = mRow
        vspGrid.Col = mCol
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        lblToolTip.Text = "Edit Current Record"
        EditMode = True
        AddMode = False
        Flag = False
        Call SetEntryMode()
        Call LockGrid(False)
        btnEdit.Enabled = False
        btnSave.Enabled = True
        btnCancel.Enabled = True
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        lblToolTip.Text = "Close Form"
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        lblToolTip.Text = "Cancel Last Action"
        Opt = MsgBox("Do you wish to Abort ?", MsgBoxStyle.YesNo)
        If Opt = MsgBoxResult.No Then
            Exit Sub
        End If
        Call ClearGrid()
        Call LoadValues()
        Flag = True
        Call SetEntryMode()
        Call LockGrid(True)
        btnSave.Enabled = False
        btnCancel.Enabled = False
    End Sub

    Sub ClearGrid()
        Call vspGrid.ClearRange(1, 1, vspGrid.MaxCols, vspGrid.MaxRows, True)
        vspGrid.MaxRows = 1
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        'Dim m, i As Integer
        'Dim mTemp As String
        lblToolTip.Text = "Save Current Record"
        If CheckValidation() Then
            'SetData()
            objSysParms.getConnection()
            objSysParms.BeginTransaction()
            If AddMode Then
                Try
                    objSysParms.DelSysParms()
                    SetData()
                    Call ClearGrid()
                    'objSysParms.SaveValues()
                    dtSysParms = objSysParms.LoadValues()
                    rowNum = dtSysParms.Rows.Count - 1
                    Call LoadValues()
                    Call btnAdd_Click(Nothing, Nothing)
                Catch ex As Exception
                    MsgBox(ex.Message)
                    objSysParms.RollBack()
                End Try
            ElseIf EditMode Then
                Try
                    objSysParms.DelSysParms()
                    SetData()
                    'objSysParms.SaveValues()
                    Flag = True
                    Call ClearGrid()
                    dtSysParms = objSysParms.LoadValues()
                    rowNum = dtSysParms.Rows.Count - 1
                    Call LoadValues()
                    Call SetEntryMode()
                    Call LockGrid(True)
                    btnSave.Enabled = False
                    btnCancel.Enabled = False
                Catch ex As Exception
                    MsgBox(ex.Message)
                    objSysParms.RollBack()
                End Try
            End If
        End If
        objSysParms.RollBack()
    End Sub

    Private Sub SetData()
        Dim mi As Integer
        Dim mTemp As Object

        For mi = 1 To vspGrid.MaxRows
            Erase mTemp
            vspGrid.GetText(1, mi, mTemp)
            objSysParms.FldName = mTemp
            Erase mTemp
            vspGrid.GetText(2, mi, mTemp)
            objSysParms.FldValue = mTemp
            Erase mTemp
            vspGrid.GetText(3, mi, mTemp)
            objSysParms.FldDescription = mTemp
            objSysParms.SaveValues()
        Next
    End Sub

    Private Function CheckValidation() As Boolean
        Dim mi As Integer
        Dim mTemp As Object

        For mi = 1 To vspGrid.MaxRows
            Erase mTemp
            vspGrid.GetText(1, mi, mTemp)
            If mTemp = "" Then
                MsgBox("Please Enter Field Name", MsgBoxStyle.Information, SysCompany)
                Call SetCol(vspGrid.ActiveRow, 1)
                Return False
            End If
        Next
        Return True
    End Function

    Private Sub vspGrid_Change(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles vspGrid.Change
        If AddMode Or EditMode Then

            Dim mTmpStr As Object
            Select Case vspGrid.ActiveCol
                Case 1  'Field Name
                    Erase mTmpStr
                    vspGrid.GetText(1, vspGrid.ActiveRow, mTmpStr)
                    If mTmpStr <> "" Then
                        Call mChkDup()
                        If FlagDup = True Then
                            vspGrid.SetText(1, vspGrid.ActiveRow, "")
                            Call SetCol(vspGrid.ActiveRow, 0)
                            Exit Sub
                        End If
                    End If
                Case 2  'Field Value
                    Erase mTmpStr
                    vspGrid.GetText(1, vspGrid.ActiveRow, mTmpStr)
                    If Trim(mTmpStr) = "" Then
                        MsgBox("You Can't give Type Without giving Name", vbInformation, SysCompany)
                        vspGrid.SetText(2, vspGrid.ActiveRow, "")
                        Call SetCol(vspGrid.ActiveRow, 2)
                        ColBack = True
                        Exit Sub
                    End If
                Case 3  'Field Description
                    Erase mTmpStr
                    vspGrid.GetText(1, vspGrid.ActiveRow, mTmpStr)
                    If Trim(mTmpStr) = "" Then
                        MsgBox("You Can't give Type Without giving Name", vbInformation, SysCompany)
                        vspGrid.SetText(3, vspGrid.ActiveRow, "")
                        Call SetCol(vspGrid.ActiveRow, 3)
                        ColBack = True
                        Exit Sub
                    End If
            End Select
        End If
    End Sub

    Private Function mChkDup() As Boolean
        Dim mDup As Integer
        Dim mDupCode1 As Object
        Dim mDupCode2 As Object

        Erase mDupCode2
        vspGrid.GetText(1, vspGrid.ActiveRow, mDupCode2)
        For mDup = 1 To vspGrid.MaxRows
            Erase mDupCode1
            vspGrid.GetText(1, mDup, mDupCode1)
            If mDupCode1 = mDupCode2 Then
                If mDup = vspGrid.ActiveRow Then
                    mDup = vspGrid.MaxRows
                Else
                    MsgBox("Duplicate Code")
                    FlagDup = True
                    Exit Function
                End If
            End If
            FlagDup = False
        Next
    End Function

    Private Sub vspGrid_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles vspGrid.KeyDownEvent
        If AddMode Or EditMode Then
            If e.keyCode = Keys.Down And vspGrid.ActiveRow = vspGrid.MaxRows Then
                vspGrid.MaxRows = vspGrid.MaxRows + 1
            ElseIf e.keyCode = Keys.Delete Then
                If SelRow Or e.shift = 2 Then
                    vspGrid.Row = vspGrid.ActiveRow
                    vspGrid.Action = 5
                    If vspGrid.MaxRows > 1 Then
                        vspGrid.MaxRows = vspGrid.MaxRows - 1
                    End If
                End If
            ElseIf e.keyCode = Keys.Insert Then
                If SelRow Then
                    vspGrid.MaxRows = vspGrid.MaxRows + 1
                    vspGrid.Row = vspGrid.ActiveRow
                    vspGrid.Action = 7
                End If
            ElseIf e.keyCode = Keys.Return And vspGrid.ActiveCol = 3 Then
                If vspGrid.ActiveRow = vspGrid.MaxRows Then
                    vspGrid.MaxRows = vspGrid.MaxRows + 1
                    vspGrid.Col = 1
                End If
                'ElseIf e.keyCode = Keys.Delete Then
                
            End If
            ' ''If e.keyCode = Keys.Delete Then
            ' ''    If SelRow Then
            ' ''        vspGrid.MaxRows = vspGrid.MaxRows - 1
            ' ''        vspGrid.Row = vspGrid.ActiveRow
            ' ''        vspGrid.Action = 7
            ' ''    End If
            ' ''End If
            '** TO Check Enter Key In Leave Cell
            'If e.keyCode = Keys.Return Or e.keyCode = Keys.Tab Then
            '    KeyReturn = True
            'Else
            '    KeyReturn = False
            'End If
            'If e.keyCode = Keys.Tab Then
            '    KeyReturn = True
            'Else
            '    KeyReturn = False
            'End If
            ''**
        End If
    End Sub

    Private Sub LockGrid(ByVal mFlag As Boolean)
        vspGrid.Row = 1
        vspGrid.Col = 1
        vspGrid.Row2 = vspGrid.MaxRows
        vspGrid.Col2 = vspGrid.MaxCols
        vspGrid.BlockMode = True
        'set the cells lock
        vspGrid.Lock = mFlag
        'set the block mode off
        vspGrid.BlockMode = False
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        lblToolTip.Text = "Refresh Records"
        Call ClearGrid()
        dtSysParms = objSysParms.LoadValues()
        rowNum = dtSysParms.Rows.Count - 1
        Call LoadValues()
    End Sub

    Private Sub vspGrid_Advance(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_AdvanceEvent) Handles vspGrid.Advance

    End Sub

    Private Sub vspGrid_ClickEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles vspGrid.ClickEvent
        If AddMode Or EditMode Then
            If e.col = 0 Then
                SelRow = True
            Else
                SelRow = False
            End If
        End If
    End Sub
End Class