<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDealy_Reason
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
        Dim Appearance36 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance40 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDealy_Reason))
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance32 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtDescription = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.txtCode = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel17 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.UltraGroupBox5 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdDelete = New Infragistics.Win.Misc.UltraButton
        Me.cmdSave = New Infragistics.Win.Misc.UltraButton
        Me.cmdReset = New Infragistics.Win.Misc.UltraButton
        Me.cmdEdit = New Infragistics.Win.Misc.UltraButton
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.UltraLabel8 = New Infragistics.Win.Misc.UltraLabel
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.txtDescription, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox1.SuspendLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox5.SuspendLayout()
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox4.SuspendLayout()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.txtDescription)
        Me.OPR0.Controls.Add(Me.UltraLabel1)
        Me.OPR0.Controls.Add(Me.txtCode)
        Me.OPR0.Controls.Add(Me.UltraLabel17)
        Me.OPR0.Location = New System.Drawing.Point(28, 12)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(627, 88)
        Me.OPR0.TabIndex = 89
        '
        'txtDescription
        '
        Appearance36.BackColor = System.Drawing.Color.White
        Appearance36.ForeColor = System.Drawing.Color.Black
        Me.txtDescription.Appearance = Appearance36
        Me.txtDescription.BackColor = System.Drawing.Color.White
        Me.txtDescription.Location = New System.Drawing.Point(122, 41)
        Me.txtDescription.MaxLength = 150
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(498, 21)
        Me.txtDescription.TabIndex = 35
        '
        'UltraLabel1
        '
        Appearance40.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel1.Appearance = Appearance40
        Me.UltraLabel1.Location = New System.Drawing.Point(6, 41)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel1.TabIndex = 34
        Me.UltraLabel1.Text = "Delay Reason"
        Me.UltraLabel1.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtCode
        '
        Appearance5.ForeColor = System.Drawing.Color.Black
        Me.txtCode.Appearance = Appearance5
        Me.txtCode.Location = New System.Drawing.Point(122, 14)
        Me.txtCode.MaxLength = 10
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(127, 21)
        Me.txtCode.TabIndex = 33
        '
        'UltraLabel17
        '
        Appearance9.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel17.Appearance = Appearance9
        Me.UltraLabel17.Location = New System.Drawing.Point(6, 14)
        Me.UltraLabel17.Name = "UltraLabel17"
        Me.UltraLabel17.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel17.TabIndex = 24
        Me.UltraLabel17.Text = "Code"
        Me.UltraLabel17.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.Controls.Add(Me.UltraGrid1)
        Me.UltraGroupBox1.Location = New System.Drawing.Point(28, 107)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(626, 212)
        Me.UltraGroupBox1.TabIndex = 93
        Me.UltraGroupBox1.Text = "Delay Reason"
        '
        'UltraGrid1
        '
        Me.UltraGrid1.Location = New System.Drawing.Point(8, 21)
        Me.UltraGrid1.Name = "UltraGrid1"
        Me.UltraGrid1.Size = New System.Drawing.Size(612, 172)
        Me.UltraGrid1.TabIndex = 1
        '
        'UltraGroupBox5
        '
        Me.UltraGroupBox5.Controls.Add(Me.cmdExit)
        Me.UltraGroupBox5.Location = New System.Drawing.Point(508, 325)
        Me.UltraGroupBox5.Name = "UltraGroupBox5"
        Me.UltraGroupBox5.Size = New System.Drawing.Size(142, 69)
        Me.UltraGroupBox5.TabIndex = 154
        '
        'cmdExit
        '
        Appearance4.Image = CType(resources.GetObject("Appearance4.Image"), Object)
        Me.cmdExit.Appearance = Appearance4
        Me.cmdExit.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdExit.Location = New System.Drawing.Point(6, 13)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(130, 47)
        Me.cmdExit.TabIndex = 3
        Me.cmdExit.Text = "&Exit"
        '
        'UltraGroupBox4
        '
        Me.UltraGroupBox4.Controls.Add(Me.cmdDelete)
        Me.UltraGroupBox4.Controls.Add(Me.cmdSave)
        Me.UltraGroupBox4.Controls.Add(Me.cmdReset)
        Me.UltraGroupBox4.Controls.Add(Me.cmdEdit)
        Me.UltraGroupBox4.Location = New System.Drawing.Point(30, 325)
        Me.UltraGroupBox4.Name = "UltraGroupBox4"
        Me.UltraGroupBox4.Size = New System.Drawing.Size(472, 69)
        Me.UltraGroupBox4.TabIndex = 153
        '
        'cmdDelete
        '
        Appearance1.Image = CType(resources.GetObject("Appearance1.Image"), Object)
        Me.cmdDelete.Appearance = Appearance1
        Me.cmdDelete.Enabled = False
        Me.cmdDelete.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDelete.ImageSize = New System.Drawing.Size(32, 32)
        Me.cmdDelete.Location = New System.Drawing.Point(236, 11)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(108, 49)
        Me.cmdDelete.TabIndex = 7
        Me.cmdDelete.Text = "&Delete"
        '
        'cmdSave
        '
        Appearance20.Image = CType(resources.GetObject("Appearance20.Image"), Object)
        Me.cmdSave.Appearance = Appearance20
        Me.cmdSave.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ImageSize = New System.Drawing.Size(32, 32)
        Me.cmdSave.Location = New System.Drawing.Point(8, 11)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(108, 49)
        Me.cmdSave.TabIndex = 6
        Me.cmdSave.Text = "&Save"
        '
        'cmdReset
        '
        Appearance16.Image = CType(resources.GetObject("Appearance16.Image"), Object)
        Me.cmdReset.Appearance = Appearance16
        Me.cmdReset.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReset.ImageSize = New System.Drawing.Size(32, 32)
        Me.cmdReset.Location = New System.Drawing.Point(350, 11)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.Size = New System.Drawing.Size(108, 49)
        Me.cmdReset.TabIndex = 5
        Me.cmdReset.Text = "&Reset"
        '
        'cmdEdit
        '
        Appearance6.Image = CType(resources.GetObject("Appearance6.Image"), Object)
        Me.cmdEdit.Appearance = Appearance6
        Me.cmdEdit.Enabled = False
        Me.cmdEdit.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEdit.ImageSize = New System.Drawing.Size(32, 32)
        Me.cmdEdit.Location = New System.Drawing.Point(122, 11)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(108, 49)
        Me.cmdEdit.TabIndex = 4
        Me.cmdEdit.Text = "Edit"
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(661, 12)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(191, 154)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 157
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(858, 59)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(102, 34)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 156
        Me.PictureBox2.TabStop = False
        '
        'UltraLabel8
        '
        Appearance32.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel8.Appearance = Appearance32
        Me.UltraLabel8.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel8.Location = New System.Drawing.Point(858, 99)
        Me.UltraLabel8.Name = "UltraLabel8"
        Me.UltraLabel8.Size = New System.Drawing.Size(299, 41)
        Me.UltraLabel8.TabIndex = 155
        Me.UltraLabel8.Text = "Master Delay Reason "
        Me.UltraLabel8.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'frmDealy_Reason
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(1124, 406)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.UltraLabel8)
        Me.Controls.Add(Me.UltraGroupBox5)
        Me.Controls.Add(Me.UltraGroupBox4)
        Me.Controls.Add(Me.UltraGroupBox1)
        Me.Controls.Add(Me.OPR0)
        Me.Name = "frmDealy_Reason"
        Me.Text = "Delay Comment"
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.txtDescription, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox1.ResumeLayout(False)
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox5.ResumeLayout(False)
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox4.ResumeLayout(False)
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtDescription As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtCode As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel17 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents UltraGroupBox5 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdDelete As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdSave As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdReset As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdEdit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents UltraLabel8 As Infragistics.Win.Misc.UltraLabel
End Class
