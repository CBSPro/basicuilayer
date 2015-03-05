<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLogin
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLogin))
        Me.picLogin = New System.Windows.Forms.PictureBox
        Me.txtDatabase = New System.Windows.Forms.TextBox
        Me.txtUserName = New System.Windows.Forms.TextBox
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.btnAbort = New System.Windows.Forms.Button
        Me.btnConnect = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.txtServer = New System.Windows.Forms.TextBox
        Me.lblserver = New System.Windows.Forms.Label
        Me.lblDatabase = New System.Windows.Forms.Label
        Me.lblUserName = New System.Windows.Forms.Label
        Me.lblPassward = New System.Windows.Forms.Label
        CType(Me.picLogin, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'picLogin
        '
        Me.picLogin.Image = CType(resources.GetObject("picLogin.Image"), System.Drawing.Image)
        Me.picLogin.Location = New System.Drawing.Point(3, 11)
        Me.picLogin.Name = "picLogin"
        Me.picLogin.Size = New System.Drawing.Size(155, 168)
        Me.picLogin.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogin.TabIndex = 0
        Me.picLogin.TabStop = False
        '
        'txtDatabase
        '
        Me.txtDatabase.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.txtDatabase.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDatabase.Location = New System.Drawing.Point(269, 47)
        Me.txtDatabase.Name = "txtDatabase"
        Me.txtDatabase.Size = New System.Drawing.Size(161, 20)
        Me.txtDatabase.TabIndex = 1
        Me.txtDatabase.Text = "Database Name"
        '
        'txtUserName
        '
        Me.txtUserName.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.txtUserName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUserName.Location = New System.Drawing.Point(269, 73)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(161, 20)
        Me.txtUserName.TabIndex = 2
        Me.txtUserName.Text = "User Name"
        '
        'txtPassword
        '
        Me.txtPassword.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.txtPassword.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPassword.Location = New System.Drawing.Point(269, 99)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(161, 20)
        Me.txtPassword.TabIndex = 3
        Me.txtPassword.Text = "Password"
        '
        'btnAbort
        '
        Me.btnAbort.Image = CType(resources.GetObject("btnAbort.Image"), System.Drawing.Image)
        Me.btnAbort.Location = New System.Drawing.Point(347, 135)
        Me.btnAbort.Name = "btnAbort"
        Me.btnAbort.Size = New System.Drawing.Size(89, 46)
        Me.btnAbort.TabIndex = 19
        Me.btnAbort.Text = "&Abort"
        Me.btnAbort.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btnAbort.UseVisualStyleBackColor = True
        '
        'btnConnect
        '
        Me.btnConnect.Image = CType(resources.GetObject("btnConnect.Image"), System.Drawing.Image)
        Me.btnConnect.Location = New System.Drawing.Point(254, 135)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.Size = New System.Drawing.Size(89, 46)
        Me.btnConnect.TabIndex = 18
        Me.btnConnect.Text = "&Connect"
        Me.btnConnect.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btnConnect.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1
        '
        'txtServer
        '
        Me.txtServer.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.txtServer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtServer.Location = New System.Drawing.Point(269, 21)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(161, 20)
        Me.txtServer.TabIndex = 20
        Me.txtServer.Text = "Server Name"
        '
        'lblserver
        '
        Me.lblserver.AutoSize = True
        Me.lblserver.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblserver.Location = New System.Drawing.Point(175, 21)
        Me.lblserver.Name = "lblserver"
        Me.lblserver.Size = New System.Drawing.Size(88, 13)
        Me.lblserver.TabIndex = 21
        Me.lblserver.Text = "Server Name :"
        Me.lblserver.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblDatabase
        '
        Me.lblDatabase.AutoSize = True
        Me.lblDatabase.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDatabase.Location = New System.Drawing.Point(158, 49)
        Me.lblDatabase.Name = "lblDatabase"
        Me.lblDatabase.Size = New System.Drawing.Size(105, 13)
        Me.lblDatabase.TabIndex = 22
        Me.lblDatabase.Text = "Database Name :"
        Me.lblDatabase.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblUserName
        '
        Me.lblUserName.AutoSize = True
        Me.lblUserName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserName.Location = New System.Drawing.Point(186, 75)
        Me.lblUserName.Name = "lblUserName"
        Me.lblUserName.Size = New System.Drawing.Size(77, 13)
        Me.lblUserName.TabIndex = 23
        Me.lblUserName.Text = "User Name :"
        Me.lblUserName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPassward
        '
        Me.lblPassward.AutoSize = True
        Me.lblPassward.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPassward.Location = New System.Drawing.Point(193, 101)
        Me.lblPassward.Name = "lblPassward"
        Me.lblPassward.Size = New System.Drawing.Size(69, 13)
        Me.lblPassward.TabIndex = 24
        Me.lblPassward.Text = "Passward :"
        Me.lblPassward.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frmLogin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(446, 193)
        Me.Controls.Add(Me.txtServer)
        Me.Controls.Add(Me.btnAbort)
        Me.Controls.Add(Me.btnConnect)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.txtUserName)
        Me.Controls.Add(Me.txtDatabase)
        Me.Controls.Add(Me.picLogin)
        Me.Controls.Add(Me.lblserver)
        Me.Controls.Add(Me.lblPassward)
        Me.Controls.Add(Me.lblUserName)
        Me.Controls.Add(Me.lblDatabase)
        Me.Name = "frmLogin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login"
        CType(Me.picLogin, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents picLogin As System.Windows.Forms.PictureBox
    Friend WithEvents txtDatabase As System.Windows.Forms.TextBox
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents btnAbort As System.Windows.Forms.Button
    Friend WithEvents btnConnect As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents lblserver As System.Windows.Forms.Label
    Friend WithEvents lblDatabase As System.Windows.Forms.Label
    Friend WithEvents lblUserName As System.Windows.Forms.Label
    Friend WithEvents lblPassward As System.Windows.Forms.Label
End Class
