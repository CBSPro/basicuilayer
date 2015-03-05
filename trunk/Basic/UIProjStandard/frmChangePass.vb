Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports Basic.DAL
Imports Basic.Constants.ProjConst
Imports Basic.DAL.Utils
Imports Basic.DAL.Security
Imports Basic.DAL.cSysParms



Public Class frmChangePass
    Inherits System.Windows.Forms.Form

    Dim objChangePass As New cChangePassword

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOk.Click
        If Trim(txtOldPass.Text) = "" Then
            MsgBox("Please Enter Your Current Password", vbInformation, SysCompany)
            txtOldPass.Focus()
            Exit Sub
        ElseIf Trim(txtNewPass.Text) = "" And Trim(txtConfirmPass.Text) = "" Then
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
        ElseIf Trim(txtOldPass.Text) <> SysPassword Then
            MsgBox("Invalid Current Password", vbInformation, SysCompany)
            txtConfirmPass.Focus()
            Exit Sub
        End If
        objChangePass.getConnection()
        objChangePass.BeginTransaction()
        Try
            objChangePass.OldPassword = Trim(txtOldPass.Text)
            objChangePass.NewPassword = Trim(txtNewPass.Text)
            'mSQL = "EXEC sp_password NULL,'" & txtOldPass.Text & "', '" & txtNewPass.Text & "'"
            mSQL = "ALTER LOGIN [" & SysUser & "] WITH PASSWORD=N'" & txtNewPass.Text & "' OLD_PASSWORD=N'" & txtOldPass.Text & "'"
            ExecQuery(mSQL)
            'objChangePass.SaveValues()
            MsgBox("User " + SysUser + " Password Successfully Changed", vbInformation, SysCompany)
            SysPassword = Trim(txtNewPass.Text)
            txtOldPass.Text = ""
            txtNewPass.Text = ""
            txtConfirmPass.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message)
            objChangePass.RollBack()
        End Try
        objChangePass.CommitTransction()
    End Sub

    Private Sub frmChangePass_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        txtOldPass.Text = ""
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

    Private Sub txtOldPass_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOldPass.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{Tab}")
        End If
    End Sub
End Class