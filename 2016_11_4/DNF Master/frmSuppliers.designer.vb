<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSuppliers
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
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSuppliers))
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtName = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.txtCode = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel17 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraGroupBox5 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdReset = New Infragistics.Win.Misc.UltraButton
        Me.cmdSave = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox3 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdDelete = New Infragistics.Win.Misc.UltraButton
        Me.cmdEdit = New Infragistics.Win.Misc.UltraButton
        Me.cmdAdd = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.txtName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox5.SuspendLayout()
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox4.SuspendLayout()
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox3.SuspendLayout()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox1.SuspendLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.txtName)
        Me.OPR0.Controls.Add(Me.UltraLabel1)
        Me.OPR0.Controls.Add(Me.txtCode)
        Me.OPR0.Controls.Add(Me.UltraLabel17)
        Me.OPR0.Enabled = False
        Me.OPR0.Location = New System.Drawing.Point(19, 12)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(626, 81)
        Me.OPR0.TabIndex = 68
        '
        'txtName
        '
        Appearance2.ForeColor = System.Drawing.Color.Black
        Me.txtName.Appearance = Appearance2
        Me.txtName.Location = New System.Drawing.Point(122, 41)
        Me.txtName.MaxLength = 80
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(498, 21)
        Me.txtName.TabIndex = 35
        '
        'UltraLabel1
        '
        Appearance7.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel1.Appearance = Appearance7
        Me.UltraLabel1.Location = New System.Drawing.Point(6, 41)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel1.TabIndex = 34
        Me.UltraLabel1.Text = "Description"
        Me.UltraLabel1.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtCode
        '
        Appearance1.BackColor = System.Drawing.Color.White
        Appearance1.ForeColor = System.Drawing.Color.Black
        Me.txtCode.Appearance = Appearance1
        Me.txtCode.BackColor = System.Drawing.Color.White
        Me.txtCode.Location = New System.Drawing.Point(122, 14)
        Me.txtCode.MaxLength = 5
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(127, 21)
        Me.txtCode.TabIndex = 33
        '
        'UltraLabel17
        '
        Appearance10.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel17.Appearance = Appearance10
        Me.UltraLabel17.Location = New System.Drawing.Point(6, 14)
        Me.UltraLabel17.Name = "UltraLabel17"
        Me.UltraLabel17.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel17.TabIndex = 24
        Me.UltraLabel17.Text = "Group Code"
        Me.UltraLabel17.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'UltraGroupBox5
        '
        Me.UltraGroupBox5.Controls.Add(Me.cmdExit)
        Me.UltraGroupBox5.Location = New System.Drawing.Point(514, 283)
        Me.UltraGroupBox5.Name = "UltraGroupBox5"
        Me.UltraGroupBox5.Size = New System.Drawing.Size(133, 50)
        Me.UltraGroupBox5.TabIndex = 78
        '
        'cmdExit
        '
        Appearance21.Image = CType(resources.GetObject("Appearance21.Image"), Object)
        Me.cmdExit.Appearance = Appearance21
        Me.cmdExit.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdExit.Location = New System.Drawing.Point(6, 10)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(121, 30)
        Me.cmdExit.TabIndex = 3
        Me.cmdExit.Text = "&Exit"
        '
        'UltraGroupBox4
        '
        Me.UltraGroupBox4.Controls.Add(Me.cmdReset)
        Me.UltraGroupBox4.Controls.Add(Me.cmdSave)
        Me.UltraGroupBox4.Location = New System.Drawing.Point(317, 284)
        Me.UltraGroupBox4.Name = "UltraGroupBox4"
        Me.UltraGroupBox4.Size = New System.Drawing.Size(192, 50)
        Me.UltraGroupBox4.TabIndex = 77
        '
        'cmdReset
        '
        Appearance16.Image = CType(resources.GetObject("Appearance16.Image"), Object)
        Me.cmdReset.Appearance = Appearance16
        Me.cmdReset.Location = New System.Drawing.Point(97, 10)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.Size = New System.Drawing.Size(85, 30)
        Me.cmdReset.TabIndex = 5
        Me.cmdReset.Text = "&Reset"
        '
        'cmdSave
        '
        Appearance20.Image = Global.DBLotVbnet.My.Resources.Resources.save_as
        Me.cmdSave.Appearance = Appearance20
        Me.cmdSave.Enabled = False
        Me.cmdSave.Location = New System.Drawing.Point(6, 10)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(85, 30)
        Me.cmdSave.TabIndex = 4
        Me.cmdSave.Text = "&Save"
        '
        'UltraGroupBox3
        '
        Me.UltraGroupBox3.Controls.Add(Me.cmdDelete)
        Me.UltraGroupBox3.Controls.Add(Me.cmdEdit)
        Me.UltraGroupBox3.Controls.Add(Me.cmdAdd)
        Me.UltraGroupBox3.Location = New System.Drawing.Point(22, 283)
        Me.UltraGroupBox3.Name = "UltraGroupBox3"
        Me.UltraGroupBox3.Size = New System.Drawing.Size(290, 50)
        Me.UltraGroupBox3.TabIndex = 76
        '
        'cmdDelete
        '
        Appearance19.Image = CType(resources.GetObject("Appearance19.Image"), Object)
        Me.cmdDelete.Appearance = Appearance19
        Me.cmdDelete.Enabled = False
        Me.cmdDelete.Location = New System.Drawing.Point(194, 12)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(85, 30)
        Me.cmdDelete.TabIndex = 2
        Me.cmdDelete.Text = "&Delete"
        '
        'cmdEdit
        '
        Appearance17.Image = CType(resources.GetObject("Appearance17.Image"), Object)
        Me.cmdEdit.Appearance = Appearance17
        Me.cmdEdit.Enabled = False
        Me.cmdEdit.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdEdit.Location = New System.Drawing.Point(100, 12)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(85, 30)
        Me.cmdEdit.TabIndex = 1
        Me.cmdEdit.Text = "&Edit"
        '
        'cmdAdd
        '
        Appearance18.Image = CType(resources.GetObject("Appearance18.Image"), Object)
        Appearance18.TextHAlignAsString = "Center"
        Me.cmdAdd.Appearance = Appearance18
        Me.cmdAdd.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdAdd.Location = New System.Drawing.Point(6, 11)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(85, 30)
        Me.cmdAdd.TabIndex = 0
        Me.cmdAdd.Text = "&Add"
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.Controls.Add(Me.UltraGrid1)
        Me.UltraGroupBox1.Location = New System.Drawing.Point(20, 101)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(626, 170)
        Me.UltraGroupBox1.TabIndex = 79
        Me.UltraGroupBox1.Text = "New Machine Group"
        '
        'UltraGrid1
        '
        Me.UltraGrid1.Location = New System.Drawing.Point(8, 21)
        Me.UltraGrid1.Name = "UltraGrid1"
        Me.UltraGrid1.Size = New System.Drawing.Size(612, 138)
        Me.UltraGrid1.TabIndex = 1
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(709, 425)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(284, 117)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 141
        Me.PictureBox1.TabStop = False
        '
        'frmSuppliers
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(998, 544)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.UltraGroupBox1)
        Me.Controls.Add(Me.UltraGroupBox5)
        Me.Controls.Add(Me.UltraGroupBox4)
        Me.Controls.Add(Me.UltraGroupBox3)
        Me.Controls.Add(Me.OPR0)
        Me.Name = "frmSuppliers"
        Me.Text = "New Machine Group"
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.txtName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox5.ResumeLayout(False)
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox4.ResumeLayout(False)
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox3.ResumeLayout(False)
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox1.ResumeLayout(False)
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtName As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtCode As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel17 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraGroupBox5 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdReset As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdSave As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox3 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdDelete As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdEdit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdAdd As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
End Class
