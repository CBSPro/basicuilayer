Imports Microsoft.Win32
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.BorderType
Imports System.Data
Imports System.Data.SqlClient
Imports Basic.DAL
Imports Basic.Constants.ProjConst
Imports Basic.DAL.Utils



Public Module Security

    Dim objConnection As Object
    Dim msCtrl As ToolStripMenuItem
    Dim Ctrl As Control
    Dim str As String
    Public mbtnAdd As Boolean
    Public mbtnEdit As Boolean
    Public mbtnDelete As Boolean
    Public mbtnPost As Boolean
    Public mbtnPrint As Boolean

    Dim objUsers As New cUsers

    Dim objMenu As New cMenu
    Dim rowNum As Integer
    Dim rowNum1 As Integer
    Dim dtLogin As New DataTable
    Dim dtSecurity As DataTable
    Dim dtMenu As DataTable
    Dim rowNum2 As Integer
    Dim mSeqNo As Integer



    Public Sub MenuSet(ByVal frm As Form, ByVal Para As Boolean)
        For Each Ctrl In frm.Controls
            If Not (TypeOf Ctrl Is PictureBox Or TypeOf Ctrl Is StatusBar) Then
                Select Case Ctrl.Text
                    Case "&File", "&Log In", "Log &Out", "E&xit", "-"
                    Case Else
                        Ctrl.Enabled = Para
                End Select
            End If
        Next Ctrl
    End Sub
    '---------------   new Menu Security -------------------
    'Public Sub SetMenuSecurity(ByVal MenuOption As String, ByVal ShowHide As Boolean)
    '    Dim TmpStr As String
    '    Dim msCtrl As System.Windows.Forms.ToolStripMenuItem

    '    For Each msCtrl In frmMdi.MenuStrip.Items
    '        If TypeOf msCtrl Is ToolStripItem Then
    '            Select Case msCtrl.Text
    '                Case "&File"
    '                Case Else
    '                    If MenuOption = msCtrl.Name Then
    '                        If ShowHide = True Then
    '                            msCtrl.Enabled = True
    '                        Else
    '                            msCtrl.Enabled = False
    '                        End If
    '                        Exit Sub
    '                        Exit Select
    '                    End If
    '            End Select
    '            chkSubMenus(msCtrl, MenuOption,showHide)
    '        End If
    '    Next msCtrl

    'End Sub

    'Private Sub chkSubMenus(ByRef msCtrl As System.Windows.Forms.ToolStripMenuItem, ByVal MenuOption As String, ByVal ShowHide As Boolean)
    '    Dim TmpStr As String
    '    Dim mstCtrl As System.Windows.Forms.ToolStripMenuItem
    '    Dim msiCtrl As System.Windows.Forms.ToolStripItem

    '    For Each msiCtrl In msCtrl.DropDownItems
    '        If Not TypeOf msiCtrl Is ToolStripSeparator Then
    '            Select Case msiCtrl.Text
    '                Case "&Log In", "Log &Out", "&Exit", "-"
    '                Case Else
    '                    If MenuOption = msiCtrl.Name Then
    '                        If ShowHide = True Then
    '                            msiCtrl.Enabled = True
    '                        Else
    '                            msiCtrl.Enabled = False
    '                        End If
    '                        Exit Sub
    '                        Exit Select
    '                    End If
    '            End Select
    '            mstCtrl = msiCtrl
    '            chkSub2Menus(mstCtrl, MenuOption, ShowHide)
    '        End If
    '    Next msiCtrl
    'End Sub
    'Private Sub chkSub2Menus(ByRef mstCtrl As System.Windows.Forms.ToolStripMenuItem, ByVal MenuOption As String, ByVal ShowHide As Boolean)
    '    Dim TmpStr As String
    '    'Dim mstCtrl As System.Windows.Forms.ToolStripMenuItem
    '    Dim msiCtrl1 As System.Windows.Forms.ToolStripItem

    '    For Each msiCtrl1 In mstCtrl.DropDownItems
    '        If Not TypeOf msiCtrl1 Is ToolStripSeparator Then
    '            Select Case msiCtrl1.Text
    '                Case "&Log In", "Log &Out", "&Exit", "-"
    '                Case Else
    '                    If MenuOption = msiCtrl1.Name Then
    '                        If ShowHide = True Then
    '                            msiCtrl1.Enabled = True
    '                        Else
    '                            msiCtrl1.Enabled = False
    '                        End If
    '                        Exit Sub
    '                        Exit Select
    '                    End If
    '            End Select
    '        End If
    '    Next msiCtrl1
    'End Sub

    Public Sub SetMenuSecurity(ByVal MenuOption As String, ByVal ShowHide As Boolean)
        'Dim msCtrl As System.Windows.Forms.ToolStripMenuItem


        'For Each msCtrl In frmMdi.MenuStrip.Items                                   'Main Menu Items Gets
        '    If TypeOf msCtrl Is ToolStripMenuItem Then
        '        Select Case msCtrl.Text
        '            Case "&File", "&Log In", "Log &Out", "E&xit", "-"
        '            Case Else
        '                If MenuOption = msCtrl.Name Then
        '                    If ShowHide = True Then
        '                        msCtrl.Enabled = True
        '                    Else
        '                        msCtrl.Enabled = False
        '                    End If
        '                    Exit Select
        '                End If
        '        End Select
        '    End If
        '    chkSubMenus(msCtrl, MenuOption, ShowHide)                               'Sub Menu Items Gets
        'Next
    End Sub
    Private Sub chkSubMenus(ByRef msCtrl As System.Windows.Forms.ToolStripMenuItem, ByVal MenuOption As String, ByVal ShowHide As Boolean)
        Dim msiCtrl As System.Windows.Forms.ToolStripItem
        Dim msiCtrl2 As System.Windows.Forms.ToolStripMenuItem

        For Each msiCtrl In msCtrl.DropDownItems                                    'Sub Menu Items Gets
            If TypeOf msiCtrl Is ToolStripMenuItem Then
                Select Case msiCtrl.Text
                    Case "&File", "&Log In", "Log &Out", "E&xit", "-"
                    Case Else
                        If MenuOption = msiCtrl.Name Then
                            If ShowHide = True Then
                                msiCtrl.Enabled = True
                            Else
                                msiCtrl.Enabled = False
                            End If
                            Exit Select
                        End If
                        If Not TypeOf msiCtrl Is ToolStripSeparator Then
                            msiCtrl2 = CType(msiCtrl, System.Windows.Forms.ToolStripMenuItem)   'Sub Menu Child Items Gets
                            If msiCtrl2.HasDropDownItems = True Then
                                chkSubMenus(msiCtrl2, MenuOption, ShowHide)
                            End If
                        End If
                End Select
            End If
            'chkchildMenus(msiCtrl, MenuOption, ShowHide)
        Next msiCtrl
    End Sub

    Private Sub chkchildMenus(ByRef msCtrl As System.Windows.Forms.ToolStripMenuItem, ByVal MenuOption As String, ByVal ShowHide As Boolean)
        Dim mscCtrl As System.Windows.Forms.ToolStripItem

        For Each mscCtrl In msCtrl.DropDownItems
            If TypeOf mscCtrl Is ToolStripMenuItem Then
                Select Case mscCtrl.Text
                    Case "&File", "&Log In", "Log &Out", "E&xit", "-"
                    Case Else
                        If MenuOption = mscCtrl.Name Then
                            If ShowHide = True Then
                                mscCtrl.Enabled = True
                            Else
                                mscCtrl.Enabled = False
                            End If
                            Exit Select
                        End If
                End Select
            End If
        Next mscCtrl
    End Sub
    Public Sub EmptyControls(ByVal frm As Form)
        Dim Ctrl As Control
        'Dim mMask As String
        Dim i As Integer
        i = 1
        For Each Ctrl In frm.Controls
            Dim CtrlName As String = Ctrl.Name
            'If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is RichTextBox Then
            If TypeOf Ctrl Is TextBox Then
                If Left(Ctrl.Name, 3) <> "TXT" Then  'These are the TEXT BOX containing some text as default
                    Ctrl.Text = ""
                End If
            ElseIf TypeOf Ctrl Is Label Then
                If Left(Ctrl.Name, 3) = "LBL" Then  'These are the labels containing Description of any key
                    Ctrl.Text = ""

                End If
            ElseIf TypeOf Ctrl Is RichTextBox Then
                Ctrl.Text = ""

            ElseIf TypeOf Ctrl Is MaskedTextBox Then
                Dim chk As CheckBox
                chk = Ctrl
                chk.Text = ""
                chk.Checked = False
                'mMask = Ctrl.Mask
                'Ctrl.Mask = ""
                'Ctrl.Text = ""
                'Ctrl.Mask = mMask
            ElseIf TypeOf Ctrl Is DateTimePicker Then
                Ctrl.Text = SySDate
            ElseIf TypeOf Ctrl Is CheckBox Then
                'Ctrl.Checked = False

                'ElseIf TypeOf Ctrl Is OptionButton Then
                '    'Ctrl.Value = 0
                '    Dim opt As optionButton
                '    chk = Ctrl
                '    chk.Checked = False
            ElseIf TypeOf Ctrl Is GroupBox Then
                For x As Integer = 0 To Ctrl.Controls.Count - 1
                    'For Each GPctrl As Control In Ctrl.Controls
                    Dim GPCtrlName As String = Ctrl.Name
                    If TypeOf Ctrl Is TextBox Then
                        Ctrl.Text = ""
                    ElseIf TypeOf Ctrl Is Label Then
                        If Left(Ctrl.Name, 3) = "LBL" Then  'These are the labels containing Description of any key
                            Ctrl.Text = ""
                        End If
                        'ElseIf TypeOf GPctrl Is MaskedTextBox Then
                        '    Dim chk As CheckBox
                        '    chk = GPctrl
                        '    chk.Text = ""
                        '    chk.Checked = False
                        'ElseIf TypeOf GPctrl Is DateTimePicker Then
                        '    GPctrl.Text = SySDate
                    End If
                    'Next
                Next

            End If

        Next
    End Sub
    'Public Sub BTNStatus(ByVal frm As Form, ByVal status As Boolean)
    '    Dim Ctrl As Control
    '    'Dim mMask As String
    '    Dim i As Integer
    '    i = 1
    '    For Each Ctrl In frm.Controls
    '        'If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is RichTextBox Then
    '        If TypeOf Ctrl Is Windows.Forms.GroupBox Then
    '            Ctrl.Enabled = False
    '        End If
    '    Next
    'End Sub
    'Public Function GetDTable(ByVal strQuery As String) As DataTable
    '    Dim DTable As New DataTable
    '    Dim DA As SqlDataAdapter
    '    Dim strConec As String

    '    Try
    '        strConec = "Data Source=" & SysServer & ";Initial Catalog=" & SysDataBase & ";User ID=" & SysUser & ";password=" & SysPassword & ";"
    '        sqlCon = New SqlConnection(strConec)
    '        sqlCon.Open()
    '        DA = New SqlDataAdapter(strQuery, sqlCon)
    '        DA.Fill(DTable)
    '        Return DTable
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Information)
    '    Finally
    '        sqlCon.Close()
    '    End Try
    '    Return DTable
    'End Function
    Public Sub SetFormSecurity(ByVal frm As Form)
        Dim mStr As String
        Dim mDt As String
        Dim DataTable As DataTable



        If SysUser = BuiltInUser Then
            frm.Tag = "1111111"
            Exit Sub
        Else
            frm.Tag = ""
        End If
        mSQL = "SELECT * FROM Security WHERE UserID = '" & SysUserID & _
                    "' AND " & "MenuName = '" & MenuName & "'"
        mStr = ""
        DataTable = GetDataTable(mSQL)

        For index As Integer = 0 To DataTable.Rows.Count - 1
            mDt = DataTable.Rows(index).Item("CanAdd").ToString()
            If mDt = True Then
                mStr = mStr & "1"
            Else
                mStr = mStr & "0"
            End If
            mDt = DataTable.Rows(index).Item("CanEdit").ToString()
            If mDt = True Then
                mStr = mStr & "1"
            Else
                mStr = mStr & "0"
            End If

            ' For Save and Cancel Buttons
            mStr = mStr & "11"
            ''
            mDt = DataTable.Rows(index).Item("CanDelete").ToString()
            If mDt = True Then
                mStr = mStr & "1"
            Else
                mStr = mStr & "0"
            End If
            mDt = DataTable.Rows(index).Item("CanPost").ToString()
            If mDt = True Then
                mStr = mStr & "1"
            Else
                mStr = mStr & "0"
            End If
            mDt = DataTable.Rows(index).Item("CanPrint").ToString()
            If mDt = True Then
                mStr = mStr & "1"
            Else
                mStr = mStr & "0"
            End If
            mDt = DataTable.Rows(index).Item("CanPreview").ToString()
            If mDt = True Then
                mStr = mStr & "1"
            Else
                mStr = mStr & "0"
            End If
            mDt = Nothing
        Next
        frm.Tag = mStr
    End Sub
    Public Sub SetButtonsSurity(ByVal frm As Form)

        If Mid(frm.Tag, 1, 1) = "1" Then                     'for ADD Button 
            mbtnAdd = True
        Else
            mbtnAdd = False
        End If
        If Mid(frm.Tag, 2, 1) = "1" Then                     'for Edit Button 
            mbtnEdit = True
        Else
            mbtnEdit = False
        End If
        If Mid(frm.Tag, 5, 1) = "1" Then                     'for Delete Button 
            mbtnDelete = True
        Else
            mbtnDelete = False
        End If
        If Mid(frm.Tag, 6, 1) = "1" Then                     'for Post Button 
            mbtnPost = True
        Else
            mbtnPost = False
        End If
        If Mid(frm.Tag, 7, 1) = "1" Then                     'for Print Button 
            mbtnPrint = True
        Else
            mbtnPrint = False
        End If

    End Sub
    Public Sub SetButtonsSeuctiry1(ByVal frm As Form, ByVal Isposted As String)

        If Mid(frm.Tag, 1, 1) = "1" Then                     'for ADD Button 
            mbtnAdd = True
        Else
            mbtnAdd = False
        End If
        If Mid(frm.Tag, 2, 1) = "1" And Isposted = "N" Then                     'for Edit Button 
            mbtnEdit = True
        Else
            mbtnEdit = False
        End If
        If Mid(frm.Tag, 5, 1) = "1" And Isposted = "N" Then                     'for Delete Button 
            mbtnDelete = True
        Else
            mbtnDelete = False
        End If
        If Mid(frm.Tag, 6, 1) = "1" And Isposted = "N" Then                     'for Post Button 
            mbtnPost = True
        Else
            mbtnPost = False
        End If
        If Mid(frm.Tag, 7, 1) = "1" Then                     'for Print Button 
            mbtnPrint = True
        Else
            mbtnPrint = False
        End If

    End Sub
    'Public Sub InitMenu()
    '    Dim TmpStr As String
    '    Dim msCtrl As System.Windows.Forms.ToolStripMenuItem
    '    mSeqNo = 0

    '    objMenu.DelMenu()

    '    dtMenu = objMenu.SelMenu
    '    rowNum1 = dtMenu.Rows.Count - 1
    '    If rowNum1 < 0 Then
    '        For Each msCtrl In frmMdi.MenuStrip.Items
    '            If TypeOf msCtrl Is ToolStripItem Then
    '                Select Case msCtrl.Text
    '                    Case "&File"
    '                    Case Else
    '                        mSeqNo = mSeqNo + 1
    '                        TmpStr = GetString(msCtrl.Text, "&")
    '                        objMenu.MenuName = Trim(msCtrl.Name.ToString)
    '                        objMenu.MenuCaption = Trim(TmpStr)
    '                        objMenu.MenuSeqNo = mSeqNo
    '                        objMenu.SaveMenu()
    '                End Select
    '                chkSubMenus(msCtrl)
    '            End If
    '        Next msCtrl
    '    End If
    'End Sub

    'Public Sub chkSubMenus(ByRef msCtrl As System.Windows.Forms.ToolStripMenuItem)
    '    Dim TmpStr As String
    '    Dim mstCtrl As System.Windows.Forms.ToolStripMenuItem
    '    Dim msiCtrl As System.Windows.Forms.ToolStripItem

    '    For Each msiCtrl In msCtrl.DropDownItems
    '        If Not TypeOf msiCtrl Is ToolStripSeparator Then
    '            Select Case msiCtrl.Text
    '                Case "&Log In", "Log &Out", "&Exit", "-"
    '                Case Else
    '                    mSeqNo = mSeqNo + 1
    '                    TmpStr = GetString(msiCtrl.Text, "&")
    '                    objMenu.MenuName = Trim(msiCtrl.Name.ToString)
    '                    objMenu.MenuCaption = Trim(TmpStr)
    '                    objMenu.MenuSeqNo = mSeqNo
    '                    objMenu.SaveMenu()
    '            End Select
    '            mstCtrl = msiCtrl
    '            chkSub2Menus(mstCtrl)
    '        End If
    '    Next msiCtrl
    'End Sub

    'Public Sub chkSub2Menus(ByRef mstCtrl As System.Windows.Forms.ToolStripMenuItem)
    '    Dim TmpStr As String
    '    'Dim mstCtrl As System.Windows.Forms.ToolStripMenuItem
    '    Dim msiCtrl1 As System.Windows.Forms.ToolStripItem

    '    For Each msiCtrl1 In mstCtrl.DropDownItems
    '        If Not TypeOf msiCtrl1 Is ToolStripSeparator Then
    '            Select Case msiCtrl1.Text
    '                Case "&Log In", "Log &Out", "&Exit", "-"
    '                Case Else
    '                    mSeqNo = mSeqNo + 1
    '                    TmpStr = GetString(msiCtrl1.Text, "&")
    '                    objMenu.MenuName = Trim(msiCtrl1.Name.ToString)
    '                    objMenu.MenuCaption = Trim(TmpStr)
    '                    objMenu.MenuSeqNo = mSeqNo
    '                    objMenu.SaveMenu()
    '            End Select
    '        End If
    '    Next msiCtrl1
    'End Sub
    'Public Function GetString(ByVal TmpStr As String, ByVal RepChr As String) As String
    '    Dim n As Integer
    '    Dim Length As Integer
    '    Dim TStr As String
    '    Dim TStr1 As String
    '    Dim TStr2 As String

    '    If Not (InStr(TmpStr, RepChr) > 0) Then
    '        GetString = TmpStr
    '        Exit Function
    '    End If

    '    Do While InStr(TmpStr, RepChr) > 0
    '        Length = Len(TmpStr)
    '        n = InStr(TmpStr, RepChr)
    '        If n = 0 And Length > 0 Then
    '            TStr = TmpStr
    '        ElseIf n = 0 Then
    '            TStr = ""
    '        ElseIf n = 1 Then
    '            TStr = Mid(TmpStr, n + 1, Length)
    '        Else
    '            TStr1 = Mid(TmpStr, 1, n - 1)
    '            TStr2 = Mid(TmpStr, n + 1, Length)
    '            TStr = TStr1 + TStr2
    '        End If
    '        TmpStr = TStr
    '    Loop
    '    GetString = TmpStr
    '    Exit Function
    'End Function

End Module
