Imports System.Text
Imports Microsoft.Win32
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles
Imports System.Data
Imports System.Data.SqlClient
Imports Basic.DAL
Imports Basic.Constants.ProjConst
Imports Basic.DAL.Utils
Imports Basic.DAL.cSysParms
Imports Basic.DAL.Security
Imports Basic.UILayer


Public Class frmLogin

    'Inherits System.Windows.Forms.Form


    Dim objConnection As Object
    Dim objLogin As New cLogin
    Dim objUsers As New cUsers
    Dim objMenu As New cMenu

    Dim dtMenu As DataTable
    Dim rowNum As Integer

    Dim dtLogin As New DataTable
    Dim rowNum1 As Integer
    Dim dtSecurity As DataTable
    Dim rowNum2 As Integer
    Dim mSeqNo As Integer

    Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
       (ByVal hwndParent As Long, ByVal fRequest As Long, _
       ByVal lpszDriver As String, ByVal lpszAttributes As String) _
       As Long

    Private Const ODBC_ADD_SYS_DSN = 4            ' Add data source
    Private Const ODBC_REMOVE_DSN = 3             ' Remove data source
    Private Const vbAPINull As Long = 0           ' NULL Pointer

    Private Sub frmLogin_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ShowMenus(False)
        'picLogin.Image.Dispose()
        'picLogin.Image = System.Drawing.Image.FromFile(Environment.CurrentDirectory & "\PNG-Logout_png-256x256.png")
        'picLogin.Text = "PNG-Logout_png-256x256.png"

        'txtServer.Text = "CERP"
        txtServer.Text = "RIO"
        txtDatabase.Text = "ERP"
        txtUserName.Text = "SA"
        txtPassword.Text = "sql123"
        'txtServer.Text = "HP-PROBOOK\HPpro"
        'txtDatabase.Text = "JKFS"
        'txtUserName.Text = "SA"
        'txtPassword.Text = "147"

    End Sub

    Private Sub btnConnect_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnConnect.Click
        Dim mdiMain As New frmMdi
        mdiMain.Text = SysSystem
        Call MenuSet(mdiMain, False)

        SysServer = Trim(txtServer.Text)
        SysDataBase = Trim(txtDatabase.Text)
        SysUser = Trim(txtUserName.Text)
        SysPassword = Trim(txtPassword.Text)

        Try
            objConnection = cConnectionManager.GetConnection()
        Catch ex As SqlException
        Finally
            dtLogin = objUsers.LoadUsers
            rowNum = dtLogin.Rows.Count - 1
            If rowNum = -1 Then
                Call Add1stUser()
                objLogin.AddFirstUser()
            End If
            objUsers.UserName = SysUser
            dtLogin = objUsers.FindUserByName
            rowNum = dtLogin.Rows.Count - 1
            If rowNum >= 0 Then
                If Not objConnection Is Nothing Then
                    Dim objDatabaseManager As IDatabaseManager
                    Dim objDBParameters As New cDBParameterList
                    objDatabaseManager = cDataBaseManager.GetDatabaseManager()
                    objDatabaseManager.SetConnection(objConnection)
                    'Call CreateDSN(SysDataBase, SysDataBase, "DSN Accounts", "SQL Server", "C:\Winnt\System32\sqlsrv32.dll", SysUser, SysServer, "")
                    dtLogin.Clear()
                    'dtLogin = objUsers.LoadUsers
                    'rowNum = dtLogin.Rows.Count - 1
                    objUsers.UserName = SysUser
                    'dtLogin.Clear()
                    objUsers.UserName = SysUser
                    dtLogin = objUsers.FindUserByName
                    rowNum = dtLogin.Rows.Count - 1
                    If rowNum >= 0 Then
                        rowNum = 0
                        'check it
                        If dtLogin.Rows(rowNum).Item("SpUser").ToString = False Then
                            SpUser = True
                        Else
                            SpUser = False
                        End If
                        SysUserID = dtLogin.Rows(rowNum).Item("UserID").ToString
                        If Len(SysUserID) = 1 Then
                            SysUserID = "000" & SysUserID
                        ElseIf Len(SysUserID) = 2 Then
                            SysUserID = "00" & SysUserID
                        ElseIf Len(SysUserID) = 3 Then
                            SysUserID = "0" & SysUserID
                        End If
                        Call CreateDSN(SysDataBase, SysDsn, "DSN Accounts", "SQL Server", "C:\Winnt\System32\sqlsrv32.dll", SysUser, SysServer, "")
                        ''Report Connection String
                        RepConn = "DSN = " & SysDsn & ";UID = " & SysUser & ";PWD =" & SysPassword & ";DSQ = " & SysDataBase & ""
                        FlagUser = True
                        'picLogin.Image.Dispose()
                        'picLogin.Image = System.Drawing.Image.FromFile(Environment.CurrentDirectory & "\PNG-Logout_png-256x256.png")
                        'picLogin.Text = "PNG-Logout_png-256x256"
                        'Timer1.Interval = 1
                        'While Timer1.Interval < 5000
                        '    Timer1.Interval = Timer1.Interval + 1
                        '    Application.DoEvents()
                        'End While
                        Call SetParms()
                        Call SetAccParam()
                        'dtLogin.Clear()
                        'dtLogin = objUsers.GetSrvDate
                        'rowNum = dtLogin.Rows.Count - 1
                        'If rowNum >= 0 Then
                        '    SySDate = dtLogin.Rows(rowNum).Item("Date")
                        '    SySDate = SySDate
                        '    'SySDate = GetDate(dtLogin.Rows(rowNum).Item("Date").ToString, "dd-MMM-yyyy")
                        'End If
                        dtLogin.Clear()
                        dtLogin = objUsers.GetSrvTime
                        rowNum = dtLogin.Rows.Count - 1
                        If rowNum >= 0 Then
                            SysTime = dtLogin.Rows(rowNum).Item("Time").ToString
                        End If
                        'frmMdi.Text = SysSystem & " , " & SysUserID & " - " & SysUser
                        Logged = True
                        Call InitMenu()
                        If SysUser = BuiltInUser Then
                            Call formEnabled()
                            Call MenuSet(mdiMain, True)
                        Else

                            mSQL = "SELECT UserID FROM USERS WHERE UserName = '" & _
                                Trim(SysUser) & "'"
                            mSQL = GetFldValue(mSQL, "UserID")
                            objUsers.UserID = mSQL
                            Call formEnabled()
                            Call Security()
                        End If
                        Me.Close()
                    Else
                        FlagUser = False
                        txtUserName.Text = ""
                        txtPassword.Text = ""
                        txtDatabase.Text = ""
                        txtServer.Focus()
                    End If
                End If
            Else
                MsgBox("User ID or Password is not correct.", MsgBoxStyle.Information)

            End If
        End Try

    End Sub

    Private Sub btnAbort_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAbort.Click
        Me.Close()
    End Sub

    '************** CODE FOR SYSTEM DSN CREATION  ******************
    '***************************************************************
    Private Function CreateDSN(ByVal DB_Name As String, _
                            ByVal DSN As String, _
                            ByVal Description As String, _
                            ByVal Driver_Name As String, _
                            ByVal Driver_Path As String, _
                            ByVal Last_User As String, _
                            ByVal Server_Name As String, _
                            ByRef Status As String _
                            ) As Boolean

        Dim regHandle As RegistryKey            ' Stores the Handle to Registry in which values need to be set

        Dim reg As RegistryKey = Registry.LocalMachine
        Dim conRegKey1 As String = "SOFTWARE\ODBC\ODBC.INI\" & DSN
        Dim conRegKey2 As String = "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"

        Try
            regHandle = reg.CreateSubKey(conRegKey1)
            regHandle.SetValue("Database", DB_Name)
            regHandle.SetValue("Description", Description)
            regHandle.SetValue("Driver", Driver_Path)
            regHandle.SetValue("LastUser", Last_User)
            regHandle.SetValue("Server", Server_Name)
            regHandle.Close()
            reg.Close()

            regHandle = reg.CreateSubKey(conRegKey2)
            regHandle.SetValue(DSN, Driver_Name)
            regHandle.Close()
            reg.Close()
            'MsgBox("Successfully created the System DSN.", MsgBoxStyle.Information, "Create System DSN")
        Catch err As Exception
            MsgBox(err.Message)
        End Try
    End Function

    Private Sub txtDatabase_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDatabase.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{Tab}")
        End If
    End Sub

    Private Sub txtPassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPassword.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{Tab}")
        End If
    End Sub

    Private Sub txtUserName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUserName.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{Tab}")
        End If
    End Sub

    Sub Add1stUser()
        objLogin.UserID = "0001"
        objLogin.UserName = BuiltInUser
        objLogin.SpUser = True
        objLogin.AddOn = SySDate
        objLogin.AddBy = "0001"
    End Sub

    Public Sub InitMenu()
        Dim TmpStr As String
        Dim msCtrl As System.Windows.Forms.ToolStripMenuItem
        mSeqNo = 0

        objMenu.DelMenu()

        dtMenu = objMenu.SelMenu
        rowNum1 = dtMenu.Rows.Count - 1
        If rowNum1 < 0 Then

            'I DID IT FOR GIVING ACEES' 
            Dim frmMdi As New frmMdi

            For Each msCtrl In frmMdi.MenuStrip.Items
                If TypeOf msCtrl Is ToolStripItem Then
                    Select Case msCtrl.Text
                        Case "&File"
                        Case Else
                            mSeqNo = mSeqNo + 1
                            TmpStr = GetString(msCtrl.Text, "&")
                            objMenu.MenuName = Trim(msCtrl.Name.ToString)
                            objMenu.MenuCaption = Trim(TmpStr)
                            objMenu.MenuSeqNo = mSeqNo
                            objMenu.SaveMenu()
                    End Select
                    chkSubMenus(msCtrl)
                End If
            Next msCtrl
        End If
    End Sub

    Public Sub chkSubMenus(ByRef msCtrl As System.Windows.Forms.ToolStripMenuItem)
        Dim TmpStr As String
        Dim mstCtrl As System.Windows.Forms.ToolStripMenuItem
        Dim msiCtrl As System.Windows.Forms.ToolStripItem

        For Each msiCtrl In msCtrl.DropDownItems
            If Not TypeOf msiCtrl Is ToolStripSeparator Then
                Select Case msiCtrl.Text
                    Case "&Log In", "Log &Out", "&Exit", "-"
                    Case Else
                        mSeqNo = mSeqNo + 1
                        TmpStr = GetString(msiCtrl.Text, "&")
                        objMenu.MenuName = Trim(msiCtrl.Name.ToString)
                        objMenu.MenuCaption = Trim(TmpStr)
                        objMenu.MenuSeqNo = mSeqNo
                        objMenu.SaveMenu()
                End Select
                mstCtrl = msiCtrl
                chkSub2Menus(mstCtrl)
            End If
        Next msiCtrl
    End Sub

    Public Sub chkSub2Menus(ByRef mstCtrl As System.Windows.Forms.ToolStripMenuItem)
        Dim TmpStr As String
        'Dim mstCtrl As System.Windows.Forms.ToolStripMenuItem
        Dim msiCtrl1 As System.Windows.Forms.ToolStripItem

        For Each msiCtrl1 In mstCtrl.DropDownItems
            If Not TypeOf msiCtrl1 Is ToolStripSeparator Then
                Select Case msiCtrl1.Text
                    Case "&Log In", "Log &Out", "&Exit", "-"
                    Case Else
                        mSeqNo = mSeqNo + 1
                        TmpStr = GetString(msiCtrl1.Text, "&")
                        objMenu.MenuName = Trim(msiCtrl1.Name.ToString)
                        objMenu.MenuCaption = Trim(TmpStr)
                        objMenu.MenuSeqNo = mSeqNo
                        objMenu.SaveMenu()
                End Select
            End If
        Next msiCtrl1
    End Sub

    Public Function GetString(ByVal TmpStr As String, ByVal RepChr As String) As String
        Dim n As Integer
        Dim Length As Integer
        Dim TStr As String
        Dim TStr1 As String
        Dim TStr2 As String

        If Not (InStr(TmpStr, RepChr) > 0) Then
            GetString = TmpStr
            Exit Function
        End If

        Do While InStr(TmpStr, RepChr) > 0
            Length = Len(TmpStr)
            n = InStr(TmpStr, RepChr)
            If n = 0 And Length > 0 Then
                TStr = TmpStr
            ElseIf n = 0 Then
                TStr = ""
            ElseIf n = 1 Then
                TStr = Mid(TmpStr, n + 1, Length)
            Else
                TStr1 = Mid(TmpStr, 1, n - 1)
                TStr2 = Mid(TmpStr, n + 1, Length)
                TStr = TStr1 + TStr2
            End If
            TmpStr = TStr
        Loop
        GetString = TmpStr
        Exit Function
    End Function

    Public Sub Security()
        Dim MenuOption As String
        Dim ShowHide As Boolean

        objUsers.UserID = SysUserID

        dtSecurity = objUsers.GetSecurity
        rowNum2 = dtSecurity.Rows.Count - 1
        If rowNum2 >= 0 Then
            rowNum2 = 0
            Do While rowNum2 < dtSecurity.Rows.Count
                MenuOption = dtSecurity.Rows(rowNum2).Item("MenuName").ToString
                ShowHide = dtSecurity.Rows(rowNum2).Item("ShowMenu").ToString
                Call SetMenuSecurity(MenuOption, ShowHide)
                rowNum2 = rowNum2 + 1
            Loop
        End If
    End Sub
    Private Sub formEnabled()
        Dim MenuOption As String
        'Dim ShowHide As Boolean

        objUsers.UserID = SysUserID

        dtSecurity = objUsers.GetSecurity
        rowNum2 = dtSecurity.Rows.Count - 1
        If rowNum2 >= 0 Then
            rowNum2 = 0
            Do While rowNum2 < dtSecurity.Rows.Count
                MenuOption = dtSecurity.Rows(rowNum2).Item("MenuName").ToString
                'ShowHide = dtSecurity.Rows(rowNum2).Item("ShowMenu").ToString
                Call SetMenuSecurity(MenuOption, True)
                rowNum2 = rowNum2 + 1
            Loop
        End If
    End Sub

    Private Sub txtUserName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUserName.LostFocus
        If Trim(txtUserName.Text) <> "" Then
            txtUserName.Text = UCase(Trim(txtUserName.Text))
        End If
    End Sub

    Private Sub txtServer_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtServer.LostFocus
        txtServer.Text = UCase(txtServer.Text)
    End Sub

End Class