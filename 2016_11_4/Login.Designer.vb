<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Login
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
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
        Dim Appearance13 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cancel = New Infragistics.Win.Misc.UltraButton
        Me.txtPassword = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.OK = New Infragistics.Win.Misc.UltraButton
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.txtUserName = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.Panel1.SuspendLayout()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox1.SuspendLayout()
        CType(Me.txtPassword, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtUserName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Panel1.Controls.Add(Me.UltraGroupBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(392, 171)
        Me.Panel1.TabIndex = 8
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.Controls.Add(Me.PictureBox1)
        Me.UltraGroupBox1.Controls.Add(Me.cancel)
        Me.UltraGroupBox1.Controls.Add(Me.txtPassword)
        Me.UltraGroupBox1.Controls.Add(Me.UltraLabel2)
        Me.UltraGroupBox1.Controls.Add(Me.OK)
        Me.UltraGroupBox1.Controls.Add(Me.UltraLabel1)
        Me.UltraGroupBox1.Controls.Add(Me.txtUserName)
        Me.UltraGroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(373, 148)
        Me.UltraGroupBox1.TabIndex = 10
        '
        'cancel
        '
        Me.cancel.Location = New System.Drawing.Point(262, 119)
        Me.cancel.Name = "cancel"
        Me.cancel.Size = New System.Drawing.Size(75, 23)
        Me.cancel.TabIndex = 9
        Me.cancel.Text = "&Cancel"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(141, 92)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(213, 21)
        Me.txtPassword.TabIndex = 7
        '
        'UltraLabel2
        '
        Appearance13.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel2.Appearance = Appearance13
        Me.UltraLabel2.Location = New System.Drawing.Point(141, 73)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(59, 21)
        Me.UltraLabel2.TabIndex = 9
        Me.UltraLabel2.Text = "Password"
        Me.UltraLabel2.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'OK
        '
        Me.OK.Location = New System.Drawing.Point(141, 119)
        Me.OK.Name = "OK"
        Me.OK.Size = New System.Drawing.Size(75, 23)
        Me.OK.TabIndex = 8
        Me.OK.Text = "&Ok"
        '
        'UltraLabel1
        '
        Appearance14.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel1.Appearance = Appearance14
        Me.UltraLabel1.Location = New System.Drawing.Point(141, 19)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(51, 21)
        Me.UltraLabel1.TabIndex = 8
        Me.UltraLabel1.Text = "Name"
        Me.UltraLabel1.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtUserName
        '
        Me.txtUserName.Location = New System.Drawing.Point(141, 46)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(213, 21)
        Me.txtUserName.TabIndex = 6
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = Global.DBLotVbnet.My.Resources.Resources.Community_Help
        Me.PictureBox1.Image = Global.DBLotVbnet.My.Resources.Resources.Community_Help
        Me.PictureBox1.Location = New System.Drawing.Point(6, 19)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(130, 119)
        Me.PictureBox1.TabIndex = 10
        Me.PictureBox1.TabStop = False
        '
        'Login
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(392, 171)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Login"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login"
        Me.TransparencyKey = System.Drawing.Color.Transparent
        Me.Panel1.ResumeLayout(False)
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox1.ResumeLayout(False)
        Me.UltraGroupBox1.PerformLayout()
        CType(Me.txtPassword, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtUserName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BindingSource1 As System.Windows.Forms.BindingSource
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtPassword As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents txtUserName As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents OK As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox

End Class
