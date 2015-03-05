
Imports System.Data.SqlClient
Imports System.IO

Public Class frmBackUp
    '    Inherits System.Windows.Forms.Form

    '    Dim objBackUp As New cBackUp

    '    Dim BakFile As String

    '#Region " Members "

    '    Private fbIgnoreClick As Boolean = False
    '    Private fbLoadingForm As Boolean = True

    '#End Region

    '#Region " Form Events "
    '#Region " ... Load "

    '    Private Sub frmBackUp_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    '        mBckDB = False
    '    End Sub
    '    Private Sub frmBackUp_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    '        Me.MdiParent = frmMdi
    '        'BakFile = "AccSys " & System.DateTime.Now
    '        BakFile = SysDsn & SySDate
    '        TurnOnHourglass()
    '        ListRootNodes()
    '        TurnOnArrow()
    '    End Sub
    '#End Region
    '#End Region

    '#Region " ... ... TurnOnHourglass "
    '    Public Sub TurnOnHourglass()
    '        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
    '    End Sub
    '#End Region

    '#Region " ... ListRootNodes "
    '    Private Sub ListRootNodes()
    '        Dim nodeText As String = ""
    '        Dim sb As New C_StringBuilder
    '        tvwLocation.Nodes.Clear()
    '        tvwLocation.BeginUpdate()
    '        With My.Computer.FileSystem
    '            For i As Integer = 0 To .Drives.Count - 1

    '                '** Build the drive's node text
    '                sb.ClearText()
    '                sb.AppendText(.Drives(i).DriveType.ToString)
    '                sb.AppendText(gtSPACE)
    '                sb.AppendText(gtLEFT_PAREN)
    '                sb.AppendText(RemoveTrailingChar(.Drives(i).Name, gtBACKSLASH))
    '                sb.AppendText(gtRIGHT_PAREN)
    '                nodeText = sb.FullText

    '                '** Add the drive to the treeview
    '                Dim driveNode As TreeNode
    '                driveNode = tvwLocation.Nodes.Add(nodeText)
    '                driveNode.ImageIndex = 0
    '                driveNode.SelectedImageIndex = 1
    '                driveNode.Tag = .Drives(i).Name
    '                If i = 0 Then
    '                    lblLocalPath.Text = CStr(driveNode.Tag)
    '                End If

    '                '** Add the next level of subfolders
    '                Try
    '                    ListLocalSubFolders(driveNode, .Drives(i).Name)
    '                Catch ex As Exception
    '                End Try
    '                driveNode = Nothing

    '            Next
    '        End With
    '        tvwLocation.EndUpdate()
    '        sb = Nothing

    '    End Sub
    '#End Region

    '#Region " ... ListLocalSubFolders "
    '    Private Sub ListLocalSubFolders(ByVal ParentNode As TreeNode, _
    '                                    ByVal ParentPath As String)
    '        Dim FolderNode As String = ""
    '        Try
    '            For Each FolderNode In Directory.GetDirectories(ParentPath)
    '                Dim childNode As TreeNode
    '                childNode = ParentNode.Nodes.Add(FilenameFromPath(FolderNode))
    '                With childNode
    '                    .ImageIndex = 0
    '                    .SelectedImageIndex = 1
    '                    .Tag = FolderNode
    '                End With
    '                childNode = Nothing
    '            Next
    '        Catch ex As Exception
    '        End Try

    '    End Sub
    '#End Region

    '#Region " Control Events "
    '#Region " ... tvwLocalFolders_AfterSelect "
    '    Private Sub tvwLocalFolders_AfterSelect(ByVal sender As Object, _
    '                                            ByVal e As System.Windows.Forms.TreeViewEventArgs) _
    '                                            Handles tvwLocation.AfterSelect
    '        If fbIgnoreClick Then
    '            Exit Sub
    '        End If

    '        '---------------------------------------------------------------------------------
    '        ' Display the path for the selected node
    '        '---------------------------------------------------------------------------------
    '        TurnOnHourglass()
    '        Dim folder As String = tvwLocation.SelectedNode.Tag
    '        lblLocalPath.Text = folder

    '        '---------------------------------------------------------------------------------
    '        ' Populate the listview with all files in the selected node
    '        '---------------------------------------------------------------------------------
    '        'ListLocalFiles(folder)
    '        TurnOnArrow()

    '    End Sub

    '#End Region
    '#Region " ... tvwLocalFolders_BeforeExpand "
    '    Private Sub tvwLocalFolders_BeforeExpand(ByVal sender As Object, _
    '                                             ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) _
    '                                             Handles tvwLocation.BeforeExpand
    '        If fbIgnoreClick Then
    '            Exit Sub
    '        End If

    '        '---------------------------------------------------------------------------------
    '        ' Display the path for the selected node
    '        '---------------------------------------------------------------------------------
    '        TurnOnHourglass()
    '        lblLocalPath.Text = e.Node.Tag

    '        '---------------------------------------------------------------------------------
    '        ' Populate all child nodes below the selected node
    '        '---------------------------------------------------------------------------------
    '        Dim parentPath As String = AddChar(e.Node.Tag)
    '        tvwLocation.BeginUpdate()
    '        Dim childNode As TreeNode = e.Node.FirstNode
    '        Do While childNode IsNot Nothing
    '            ListLocalSubFolders(childNode, parentPath & childNode.Text)
    '            childNode = childNode.NextNode
    '        Loop
    '        tvwLocation.EndUpdate()
    '        tvwLocation.Refresh()

    '        '---------------------------------------------------------------------------------
    '        ' Select the node being expanded
    '        '---------------------------------------------------------------------------------
    '        tvwLocation.SelectedNode = e.Node

    '        '---------------------------------------------------------------------------------
    '        ' Populate the listview with all files in the selected node
    '        '---------------------------------------------------------------------------------
    '        'ListLocalFiles(parentPath)
    '        TurnOnArrow()

    '    End Sub
    '#End Region
    '#End Region

    '    Private Sub btnBackup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBackup.Click
    '        If lblLocalPath.Text = "" Then
    '            MsgBox("Incomplete Path")
    '            Exit Sub
    '        End If
    '        If Mid(lblLocalPath.Text, Len(lblLocalPath.Text), 1) = "\" Then
    '            objBackUp.BackupPath = lblLocalPath.Text & BakFile
    '        Else
    '            objBackUp.BackupPath = lblLocalPath.Text & "\" & BakFile
    '        End If
    '        objBackUp.BackUp()
    '        MsgBox("Backup Done Successfully")
    '        Me.Close()
    '    End Sub


    '    Private Sub frmBackUp_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
    '        Me.WindowState = FormWindowState.Normal
    '    End Sub


End Class