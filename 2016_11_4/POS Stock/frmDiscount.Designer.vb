<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDiscount
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
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDiscount))
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance40 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance36 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance24 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.UltraGroupBox5 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdReset = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox3 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdDelete = New Infragistics.Win.Misc.UltraButton
        Me.cmdEdit = New Infragistics.Win.Misc.UltraButton
        Me.cmdAdd = New Infragistics.Win.Misc.UltraButton
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtDiscount = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.txtDescription = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.txtCode = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel17 = New Infragistics.Win.Misc.UltraLabel
        Me.OPR5 = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraGrid2 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.txtFind = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel13 = New Infragistics.Win.Misc.UltraLabel
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox1.SuspendLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox5.SuspendLayout()
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox4.SuspendLayout()
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox3.SuspendLayout()
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.txtDiscount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDescription, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.OPR5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR5.SuspendLayout()
        CType(Me.UltraGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtFind, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.Controls.Add(Me.UltraGrid1)
        Me.UltraGroupBox1.Location = New System.Drawing.Point(35, 119)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(626, 212)
        Me.UltraGroupBox1.TabIndex = 102
        Me.UltraGroupBox1.Text = " "
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
        Me.UltraGroupBox5.Location = New System.Drawing.Point(529, 337)
        Me.UltraGroupBox5.Name = "UltraGroupBox5"
        Me.UltraGroupBox5.Size = New System.Drawing.Size(133, 64)
        Me.UltraGroupBox5.TabIndex = 101
        '
        'cmdExit
        '
        Appearance21.Image = CType(resources.GetObject("Appearance21.Image"), Object)
        Me.cmdExit.Appearance = Appearance21
        Me.cmdExit.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdExit.Location = New System.Drawing.Point(6, 10)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(121, 48)
        Me.cmdExit.TabIndex = 3
        Me.cmdExit.Text = "&Exit"
        '
        'UltraGroupBox4
        '
        Me.UltraGroupBox4.Controls.Add(Me.cmdReset)
        Me.UltraGroupBox4.Location = New System.Drawing.Point(405, 337)
        Me.UltraGroupBox4.Name = "UltraGroupBox4"
        Me.UltraGroupBox4.Size = New System.Drawing.Size(119, 64)
        Me.UltraGroupBox4.TabIndex = 100
        '
        'cmdReset
        '
        Appearance16.Image = CType(resources.GetObject("Appearance16.Image"), Object)
        Me.cmdReset.Appearance = Appearance16
        Me.cmdReset.ImageSize = New System.Drawing.Size(32, 32)
        Me.cmdReset.Location = New System.Drawing.Point(8, 10)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.Size = New System.Drawing.Size(105, 48)
        Me.cmdReset.TabIndex = 5
        Me.cmdReset.Text = "&Reset"
        '
        'UltraGroupBox3
        '
        Me.UltraGroupBox3.Controls.Add(Me.cmdDelete)
        Me.UltraGroupBox3.Controls.Add(Me.cmdEdit)
        Me.UltraGroupBox3.Controls.Add(Me.cmdAdd)
        Me.UltraGroupBox3.Location = New System.Drawing.Point(37, 337)
        Me.UltraGroupBox3.Name = "UltraGroupBox3"
        Me.UltraGroupBox3.Size = New System.Drawing.Size(362, 64)
        Me.UltraGroupBox3.TabIndex = 99
        '
        'cmdDelete
        '
        Appearance19.Image = CType(resources.GetObject("Appearance19.Image"), Object)
        Me.cmdDelete.Appearance = Appearance19
        Me.cmdDelete.Location = New System.Drawing.Point(243, 10)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(113, 47)
        Me.cmdDelete.TabIndex = 2
        Me.cmdDelete.Text = "&Delete"
        '
        'cmdEdit
        '
        Appearance17.Image = CType(resources.GetObject("Appearance17.Image"), Object)
        Me.cmdEdit.Appearance = Appearance17
        Me.cmdEdit.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdEdit.Location = New System.Drawing.Point(125, 10)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(113, 47)
        Me.cmdEdit.TabIndex = 1
        Me.cmdEdit.Text = "&Edit"
        '
        'cmdAdd
        '
        Appearance18.Image = CType(resources.GetObject("Appearance18.Image"), Object)
        Appearance18.TextHAlignAsString = "Center"
        Me.cmdAdd.Appearance = Appearance18
        Me.cmdAdd.ImageSize = New System.Drawing.Size(32, 32)
        Me.cmdAdd.Location = New System.Drawing.Point(6, 10)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(113, 47)
        Me.cmdAdd.TabIndex = 0
        Me.cmdAdd.Text = "&Add"
        '
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.txtDiscount)
        Me.OPR0.Controls.Add(Me.UltraLabel2)
        Me.OPR0.Controls.Add(Me.txtDescription)
        Me.OPR0.Controls.Add(Me.UltraLabel1)
        Me.OPR0.Controls.Add(Me.txtCode)
        Me.OPR0.Controls.Add(Me.UltraLabel17)
        Me.OPR0.Location = New System.Drawing.Point(35, 12)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(627, 101)
        Me.OPR0.TabIndex = 98
        '
        'txtDiscount
        '
        Appearance1.ForeColor = System.Drawing.Color.Black
        Me.txtDiscount.Appearance = Appearance1
        Me.txtDiscount.Location = New System.Drawing.Point(122, 68)
        Me.txtDiscount.MaxLength = 30
        Me.txtDiscount.Name = "txtDiscount"
        Me.txtDiscount.Size = New System.Drawing.Size(70, 21)
        Me.txtDiscount.TabIndex = 37
        '
        'UltraLabel2
        '
        Appearance40.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel2.Appearance = Appearance40
        Me.UltraLabel2.Location = New System.Drawing.Point(8, 68)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel2.TabIndex = 36
        Me.UltraLabel2.Text = "Discount %"
        Me.UltraLabel2.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtDescription
        '
        Appearance36.BackColor = System.Drawing.Color.White
        Appearance36.ForeColor = System.Drawing.Color.Black
        Me.txtDescription.Appearance = Appearance36
        Me.txtDescription.BackColor = System.Drawing.Color.White
        Me.txtDescription.Location = New System.Drawing.Point(122, 41)
        Me.txtDescription.MaxLength = 80
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(498, 21)
        Me.txtDescription.TabIndex = 35
        '
        'UltraLabel1
        '
        Appearance2.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel1.Appearance = Appearance2
        Me.UltraLabel1.Location = New System.Drawing.Point(6, 41)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel1.TabIndex = 34
        Me.UltraLabel1.Text = "Item Name"
        Me.UltraLabel1.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtCode
        '
        Appearance3.ForeColor = System.Drawing.Color.Black
        Me.txtCode.Appearance = Appearance3
        Me.txtCode.Location = New System.Drawing.Point(122, 14)
        Me.txtCode.MaxLength = 30
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
        Me.UltraLabel17.Text = "Item Code"
        Me.UltraLabel17.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'OPR5
        '
        Me.OPR5.Controls.Add(Me.UltraGrid2)
        Me.OPR5.Controls.Add(Me.txtFind)
        Me.OPR5.Controls.Add(Me.UltraLabel13)
        Me.OPR5.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OPR5.Location = New System.Drawing.Point(35, 53)
        Me.OPR5.Name = "OPR5"
        Me.OPR5.Size = New System.Drawing.Size(654, 421)
        Me.OPR5.TabIndex = 115
        Me.OPR5.Text = "Item Find ...."
        Me.OPR5.Visible = False
        '
        'UltraGrid2
        '
        Me.UltraGrid2.Location = New System.Drawing.Point(11, 54)
        Me.UltraGrid2.Name = "UltraGrid2"
        Me.UltraGrid2.Size = New System.Drawing.Size(632, 358)
        Me.UltraGrid2.TabIndex = 105
        '
        'txtFind
        '
        Appearance24.ForeColor = System.Drawing.Color.Black
        Me.txtFind.Appearance = Appearance24
        Me.txtFind.Location = New System.Drawing.Point(56, 23)
        Me.txtFind.MaxLength = 60
        Me.txtFind.Name = "txtFind"
        Me.txtFind.Size = New System.Drawing.Size(582, 21)
        Me.txtFind.TabIndex = 37
        '
        'UltraLabel13
        '
        Appearance25.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel13.Appearance = Appearance25
        Me.UltraLabel13.Location = New System.Drawing.Point(6, 23)
        Me.UltraLabel13.Name = "UltraLabel13"
        Me.UltraLabel13.Size = New System.Drawing.Size(162, 21)
        Me.UltraLabel13.TabIndex = 36
        Me.UltraLabel13.Text = "##"
        Me.UltraLabel13.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'frmDiscount
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(829, 485)
        Me.Controls.Add(Me.OPR5)
        Me.Controls.Add(Me.UltraGroupBox1)
        Me.Controls.Add(Me.UltraGroupBox5)
        Me.Controls.Add(Me.UltraGroupBox4)
        Me.Controls.Add(Me.UltraGroupBox3)
        Me.Controls.Add(Me.OPR0)
        Me.Name = "frmDiscount"
        Me.Text = "Discount Items"
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox1.ResumeLayout(False)
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox5.ResumeLayout(False)
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox4.ResumeLayout(False)
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox3.ResumeLayout(False)
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.txtDiscount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDescription, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.OPR5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR5.ResumeLayout(False)
        Me.OPR5.PerformLayout()
        CType(Me.UltraGrid2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtFind, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents UltraGroupBox5 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdReset As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox3 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdDelete As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdEdit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdAdd As Infragistics.Win.Misc.UltraButton
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtDescription As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtCode As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel17 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtDiscount As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents OPR5 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraGrid2 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents txtFind As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel13 As Infragistics.Win.Misc.UltraLabel
End Class
