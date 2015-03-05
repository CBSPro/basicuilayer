<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBackUp
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBackUp))
        Me.GpBackup = New System.Windows.Forms.GroupBox
        Me.tvwLocation = New System.Windows.Forms.TreeView
        Me.lblLocalPath = New System.Windows.Forms.Label
        Me.btnBackup = New System.Windows.Forms.Button
        Me.GpBackup.SuspendLayout()
        Me.SuspendLayout()
        '
        'GpBackup
        '
        Me.GpBackup.Controls.Add(Me.tvwLocation)
        Me.GpBackup.Controls.Add(Me.lblLocalPath)
        Me.GpBackup.Location = New System.Drawing.Point(7, 0)
        Me.GpBackup.Name = "GpBackup"
        Me.GpBackup.Size = New System.Drawing.Size(430, 221)
        Me.GpBackup.TabIndex = 0
        Me.GpBackup.TabStop = False
        Me.GpBackup.Text = "Backup Location"
        '
        'tvwLocation
        '
        Me.tvwLocation.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.tvwLocation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tvwLocation.HideSelection = False
        Me.tvwLocation.Location = New System.Drawing.Point(6, 51)
        Me.tvwLocation.Name = "tvwLocation"
        Me.tvwLocation.Size = New System.Drawing.Size(418, 164)
        Me.tvwLocation.TabIndex = 1
        '
        'lblLocalPath
        '
        Me.lblLocalPath.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.lblLocalPath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLocalPath.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLocalPath.Location = New System.Drawing.Point(6, 25)
        Me.lblLocalPath.Name = "lblLocalPath"
        Me.lblLocalPath.Size = New System.Drawing.Size(418, 23)
        Me.lblLocalPath.TabIndex = 0
        Me.lblLocalPath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnBackup
        '
        Me.btnBackup.Image = CType(resources.GetObject("btnBackup.Image"), System.Drawing.Image)
        Me.btnBackup.Location = New System.Drawing.Point(146, 227)
        Me.btnBackup.Name = "btnBackup"
        Me.btnBackup.Size = New System.Drawing.Size(114, 40)
        Me.btnBackup.TabIndex = 1
        Me.btnBackup.Text = "Start Backup"
        Me.btnBackup.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnBackup.UseVisualStyleBackColor = True
        '
        'frmBackUp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(444, 272)
        Me.Controls.Add(Me.btnBackup)
        Me.Controls.Add(Me.GpBackup)
        Me.Name = "frmBackUp"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Backup Administration"
        Me.GpBackup.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GpBackup As System.Windows.Forms.GroupBox
    Friend WithEvents lblLocalPath As System.Windows.Forms.Label
    Friend WithEvents tvwLocation As System.Windows.Forms.TreeView
    Friend WithEvents btnBackup As System.Windows.Forms.Button
End Class
