Imports System.Windows.Forms

Imports System.Data
Imports System.Data.SqlClient
Imports Basic.DAL
Imports Basic.Constants.ProjConst
Imports Basic.DAL.Utils



Public Class Grid_Help
    Inherits System.Windows.Forms.Form
    Private Message As String
    Private bok As Boolean
    Private strfindrow As String
    Private strPKValue As String
    Private sqlquery As String
    Private HelpDs As DataSet
    Public dtLookup As DataTable
    Dim vColumn As Integer

    Public Property sqlqueryFun() As String
        Get
            Return Me.sqlquery
        End Get
        Set(ByVal value As String)
            Me.sqlquery = value
        End Set
    End Property
    Public Property strFindfun() As String
        Get
            Return Me.strfindrow
        End Get
        Set(ByVal value As String)
            Me.strfindrow = value
        End Set
    End Property
    Public Property strPKfun() As String
        Get
            Return Me.strPKValue
        End Get
        Set(ByVal value As String)
            Me.strPKValue = value
        End Set
    End Property

    Public Property PMessage() As String
        Get
            Return Me.Message
        End Get
        Set(ByVal Value As String)
            Me.Message = Value
        End Set
    End Property

    Public Property PbOk() As Boolean
        Get
            Return bok
        End Get
        Set(ByVal Value As Boolean)
            Me.bok = Value
        End Set
    End Property

    Private Sub Help_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub

    Private Sub Help_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblNoOfRecords.Visible = False
        Try
            Me.Text = Me.PMessage
            dtLookup = Lookup(Me.sqlqueryFun).Tables(0)
            Me.GVHelp.DataSource = dtLookup.DefaultView
            Me.GVHelp.Columns(1).Width = 250
            lblNoOfRecords.Text = Me.GVHelp.RowCount & " Records Found"
        Catch ex As Exception
            ex.Message.ToString()
        End Try


    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        If Me.GVHelp.Rows.Count > 0 Then
            Me.PbOk = True
            Me.strfindrow = GVHelp.CurrentRow.Index
            Me.strPKValue = Me.GVHelp.Item(0, GVHelp.CurrentRow.Index).Value
            Me.sqlquery = Me.GVHelp.Item(1, GVHelp.CurrentRow.Index).Value
            Me.Close()
        Else : MsgBox("Record Does Not Exist", MsgBoxStyle.Information, SysCompany)
            Me.txtSearch.Text = String.Empty
            Me.PbOk = False
            Me.strPKValue = 0
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.PbOk = False
        Me.strPKValue = 0
        Me.Close()
    End Sub

    Private Sub txtSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSearch.KeyDown
        If e.KeyCode = Keys.Enter And Me.GVHelp.Rows.Count > 0 Then
            Me.PbOk = True
            If mOrder = 0 Then
                Me.strPKValue = Me.GVHelp.Item(0, GVHelp.CurrentRow.Index).Value
                Me.sqlquery = Me.GVHelp.Item(1, GVHelp.CurrentRow.Index).Value
            ElseIf mOrder = 1 Then
                Me.strPKValue = Me.GVHelp.Item(1, GVHelp.CurrentRow.Index).Value
                Me.sqlquery = Me.GVHelp.Item(0, GVHelp.CurrentRow.Index).Value
            End If
            Me.Close()
        ElseIf e.KeyCode = Keys.Enter And Me.GVHelp.Rows.Count <= 0 Then
            'MsgBox("Record Does Not Exist", MsgBoxStyle.Information, SysCompany)
            Me.txtSearch.Text = String.Empty
            Me.PbOk = False
            Me.strPKValue = 0
        End If
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{Tab}")
        End If
    End Sub

    Private Sub txtSearch_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearch.KeyPress
        If (Char.IsNumber(e.KeyChar)) <> True And (Char.IsLetter(e.KeyChar)) <> True And (Char.IsWhiteSpace(e.KeyChar) <> True) And (Char.IsControl(e.KeyChar)) <> True Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged
        vColumn = 0
        If txtSearch.Text = "" Then
            'vColumn = 0
            dtLookup.DefaultView.RowFilter = Nothing
            'lblNoOfRecords.Text = Me.GVHelp.RowCount & " Records Found"
        Else
            With dtLookup
                If .Columns(vColumn).DataType.ToString = "System.Int32" Or .Columns(vColumn).DataType.ToString = "System.Double" Then
                    .DefaultView.RowFilter = .Columns(vColumn).Caption & " = " & Val(txtSearch.Text)
                ElseIf .Columns(vColumn).DataType.ToString = "System.String" Or .Columns(vColumn).DataType.ToString = "System.char" Then
                    .DefaultView.RowFilter = .Columns(vColumn).Caption & " like '" & txtSearch.Text & "%'"
                End If
                lblNoOfRecords.Text = Me.GVHelp.RowCount & " Records Found"
            End With
        End If
    End Sub

    Private Sub GVHelp_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GVHelp.CellClick
        vColumn = e.ColumnIndex
    End Sub

    Private Sub GVHelp_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles GVHelp.DoubleClick
        If GVHelp.RowCount <> 0 Then
            Me.PbOk = True
            Me.strPKfun = Me.GVHelp.Item(0, GVHelp.CurrentRow.Index).Value
            Me.sqlquery = Me.GVHelp.Item(1, GVHelp.CurrentRow.Index).Value
            btnOk_Click(sender, e)
            Me.Close()
        End If
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter And Me.GVHelp.Rows.Count > 0 Then
            Me.PbOk = True
            If mOrder = 0 Then
                Me.strPKValue = Me.GVHelp.Item(0, GVHelp.CurrentRow.Index).Value
                Me.sqlquery = Me.GVHelp.Item(1, GVHelp.CurrentRow.Index).Value
            ElseIf mOrder = 1 Then
                Me.strPKValue = Me.GVHelp.Item(1, GVHelp.CurrentRow.Index).Value
                Me.sqlquery = Me.GVHelp.Item(0, GVHelp.CurrentRow.Index).Value
            End If
            Me.Close()
        ElseIf e.KeyCode = Keys.Enter And Me.GVHelp.Rows.Count <= 0 Then
            'MsgBox("Record Does Not Exist", MsgBoxStyle.Information, SysCompany)
            Me.TextBox1.Text = String.Empty
            Me.PbOk = False
            Me.strPKValue = 0
        End If
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{Tab}")
        End If
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If (Char.IsNumber(e.KeyChar)) <> True And (Char.IsLetter(e.KeyChar)) <> True And (Char.IsWhiteSpace(e.KeyChar) <> True) And (Char.IsControl(e.KeyChar)) <> True And (Char.IsPunctuation(e.KeyChar)) <> True And (Char.IsSymbol(e.KeyChar)) <> True Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        vColumn = 1
        If TextBox1.Text = "" Then
            'vColumn = 0
            dtLookup.DefaultView.RowFilter = Nothing
            'lblNoOfRecords.Text = Me.GVHelp.RowCount & " Records Found"
        Else
            With dtLookup
                If .Columns(vColumn).DataType.ToString = "System.Int32" Or .Columns(vColumn).DataType.ToString = "System.Double" Then
                    .DefaultView.RowFilter = .Columns(vColumn).Caption & " = " & Val(TextBox1.Text)
                ElseIf .Columns(vColumn).DataType.ToString = "System.String" Or .Columns(vColumn).DataType.ToString = "System.char" Then
                    .DefaultView.RowFilter = .Columns(vColumn).Caption & " like '" & TextBox1.Text & "%'"
                End If
                'lblNoOfRecords.Text = Me.GVHelp.RowCount & " Records Found"
            End With
        End If
    End Sub

    Private Sub GVHelp_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GVHelp.CellContentClick

    End Sub

    Private Sub GVHelp_CellErrorTextChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GVHelp.CellErrorTextChanged

    End Sub
End Class