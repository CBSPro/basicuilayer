
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.BorderType
Imports Basic.Constants.ProjConst
Imports Basic.DAL
Imports Basic.DAL.Utils
Imports Basic.UILayer.frmMdi



Public Class frmChangeUPass
    Inherits System.Windows.Forms.Form

    Dim objChangeUPass As New cChangeUPassword

    Dim PODS As New DataSet
    Dim objRow As Data.DataRow
    Dim ObjFind As Grid_Help
    Dim tableName As String = "Users"
    Dim strFind As String
    Dim dtAccount As New DataTable
    Dim rowNum As Integer

    Private Sub txtUserName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUserName.KeyDown
        ObjFind = New Grid_Help
        If e.KeyCode = Keys.F1 Then
            ObjFind.PMessage = "Users"
            strFind = "SELECT UserName,UserID FROM USERS Order By UserName"
            ObjFind.sqlqueryFun = strFind
            ObjFind.ShowDialog()

            If ObjFind.PbOk = True Then
                Me.txtUserName.Text = ObjFind.strPKfun & ""
            End If
        ElseIf e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub txtUserName_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtUserName.Validating
        If Trim(txtUserName.Text) <> "" Then
            mSQL = "SELECT UserName FROM USERS WHERE UserName = '" & _
                    Trim(txtUserName.Text) & "'"
            mSQL = GetFldValue(mSQL, "UserName")
            If mSQL = "" Then
                MsgBox("No Record Found", vbInformation, SysCompany)
                txtUserName.Text = ""
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub btnOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOk.Click
        If Not SpUser Then
            MsgBox("You are not Allowed to Change Other Users Password", vbInformation, SysCompany)
            Exit Sub
        End If

        If Trim(txtNewPass.Text) = "" And Trim(txtConfirmPass.Text) = "" Then
            MsgBox("Please Enter New Password", vbInformation, SysCompany)
            txtNewPass.Focus()
            Exit Sub
        ElseIf Trim(txtConfirmPass.Text) = "" Then
            MsgBox("Please Enter Password Confirmation", vbInformation, SysCompany)
            txtConfirmPass.Focus()
            Exit Sub
        ElseIf txtNewPass.Text <> txtConfirmPass.Text Then
            MsgBox("Password Confirmation Failed", vbInformation, SysCompany)
            txtConfirmPass.Focus()
            Exit Sub
        End If
        objChangeUPass.getConnection()
        objChangeUPass.BeginTransaction()
        Try

            mSQL = "USE[master] " & _
                    "CREATE LOGIN [" + txtUserName.Text + "] WITH PASSWORD=N'" + txtNewPass.Text + "', DEFAULT_DATABASE=[master], CHECK_EXPIRATION=OFF, CHECK_POLICY=OFF " & _
                    "EXEC master..sp_addsrvrolemember @loginame = N'" + txtUserName.Text + "', @rolename = N'sysadmin' " & _
                    "USE [" + SysDataBase + "] " & _
                    "CREATE USER [" + txtUserName.Text + "] FOR LOGIN [" + txtUserName.Text + "]  " & _
                    "USE [" + SysDataBase + "]  " & _
                    "EXEC sp_addrolemember N'db_owner', N'" + txtUserName.Text + "'  "
            ExecQuery(mSQL)

            'objChangeUPass.NewPassword = Trim(txtNewPass.Text)
            'objChangeUPass.UserName = Trim(txtUserName.Text)
            'mSQL = "EXEC sp_password NULL,'" & txtNewPass.Text & "', '" & txtUserName.Text & "'"
            'objChangeUPass.SaveValues()
            MsgBox("Your Password Successfully Changed", vbInformation, SysCompany)
            SysPassword = Trim(txtNewPass.Text)
            txtUserName.Text = ""
            txtNewPass.Text = ""
            txtConfirmPass.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message)
            objChangeUPass.RollBack()
        End Try
        objChangeUPass.CommitTransction()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub frmChangeUPass_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        mUPass = False
    End Sub

    Private Sub frmChangeUPass_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.MdiParent = frmMdi
        txtUserName.Text = ""
        txtNewPass.Text = ""
        txtConfirmPass.Text = ""
    End Sub

    Private Sub txtConfirmPass_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtConfirmPass.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{Tab}")
        End If
    End Sub

    Private Sub txtNewPass_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNewPass.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{Tab}")
        End If
    End Sub

    Private Sub frmChangeUPass_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        Me.WindowState = FormWindowState.Normal
    End Sub
End Class