<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLog
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLog))
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtPW = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.txtUname = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel16 = New Infragistics.Win.Misc.UltraLabel
        Me.cmdAdd = New Infragistics.Win.Misc.UltraButton
        Me.cmdSave = New Infragistics.Win.Misc.UltraButton
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.txtPW, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtUname, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.txtPW)
        Me.OPR0.Controls.Add(Me.UltraLabel1)
        Me.OPR0.Controls.Add(Me.txtUname)
        Me.OPR0.Controls.Add(Me.UltraLabel16)
        Me.OPR0.Location = New System.Drawing.Point(12, 4)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(285, 79)
        Me.OPR0.TabIndex = 0
        '
        'txtPW
        '
        Appearance3.ForeColor = System.Drawing.Color.Black
        Appearance3.ForeColorDisabled = System.Drawing.Color.White
        Me.txtPW.Appearance = Appearance3
        Me.txtPW.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPW.Location = New System.Drawing.Point(109, 44)
        Me.txtPW.Name = "txtPW"
        Me.txtPW.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPW.Size = New System.Drawing.Size(170, 21)
        Me.txtPW.TabIndex = 26
        '
        'UltraLabel1
        '
        Appearance23.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel1.Appearance = Appearance23
        Me.UltraLabel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel1.Location = New System.Drawing.Point(6, 44)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(97, 21)
        Me.UltraLabel1.TabIndex = 25
        Me.UltraLabel1.Text = "Password"
        Me.UltraLabel1.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtUname
        '
        Appearance2.ForeColor = System.Drawing.Color.Black
        Appearance2.ForeColorDisabled = System.Drawing.Color.White
        Me.txtUname.Appearance = Appearance2
        Me.txtUname.Location = New System.Drawing.Point(109, 17)
        Me.txtUname.Name = "txtUname"
        Me.txtUname.Size = New System.Drawing.Size(170, 21)
        Me.txtUname.TabIndex = 24
        '
        'UltraLabel16
        '
        Appearance1.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel16.Appearance = Appearance1
        Me.UltraLabel16.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel16.Location = New System.Drawing.Point(6, 17)
        Me.UltraLabel16.Name = "UltraLabel16"
        Me.UltraLabel16.Size = New System.Drawing.Size(97, 21)
        Me.UltraLabel16.TabIndex = 22
        Me.UltraLabel16.Text = "User Name"
        Me.UltraLabel16.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'cmdAdd
        '
        Appearance18.Image = Global.DBLotVbnet.My.Resources.Resources.add
        Appearance18.TextHAlignAsString = "Center"
        Me.cmdAdd.Appearance = Appearance18
        Me.cmdAdd.Location = New System.Drawing.Point(103, 89)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(85, 30)
        Me.cmdAdd.TabIndex = 1
        Me.cmdAdd.Text = "&Reset"
        '
        'cmdSave
        '
        Appearance20.Image = CType(resources.GetObject("Appearance20.Image"), Object)
        Me.cmdSave.Appearance = Appearance20
        Me.cmdSave.Enabled = False
        Me.cmdSave.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdSave.Location = New System.Drawing.Point(12, 89)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(85, 30)
        Me.cmdSave.TabIndex = 5
        Me.cmdSave.Text = "&Login"
        '
        'cmdExit
        '
        Appearance21.Image = CType(resources.GetObject("Appearance21.Image"), Object)
        Me.cmdExit.Appearance = Appearance21
        Me.cmdExit.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdExit.Location = New System.Drawing.Point(194, 89)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(103, 30)
        Me.cmdExit.TabIndex = 6
        Me.cmdExit.Text = "&Exit"
        '
        'frmLog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(309, 127)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.OPR0)
        Me.MaximizeBox = False
        Me.Name = "frmLog"
        Me.Text = "Authorize Action"
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.txtPW, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtUname, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraLabel16 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtPW As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtUname As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents cmdAdd As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdSave As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
End Class
