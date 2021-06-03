<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDiscount_item_Uniq
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDiscount_item_Uniq))
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance40 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance36 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.ReportFilterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.RefreshToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ExitToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.DeactivateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.lblDisplay = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.ByProductCategoryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtDiscount = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.txtDescription = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.txtCode = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel17 = New Infragistics.Win.Misc.UltraLabel
        Me.MenuStrip1.SuspendLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.txtDiscount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDescription, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCode, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReportFilterToolStripMenuItem, Me.EditToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(982, 24)
        Me.MenuStrip1.TabIndex = 110
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ReportFilterToolStripMenuItem
        '
        Me.ReportFilterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1, Me.RefreshToolStripMenuItem, Me.ExitToolStripMenuItem1})
        Me.ReportFilterToolStripMenuItem.Name = "ReportFilterToolStripMenuItem"
        Me.ReportFilterToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.ReportFilterToolStripMenuItem.Text = "File"
        '
        'RefreshToolStripMenuItem
        '
        Me.RefreshToolStripMenuItem.Image = CType(resources.GetObject("RefreshToolStripMenuItem.Image"), System.Drawing.Image)
        Me.RefreshToolStripMenuItem.Name = "RefreshToolStripMenuItem"
        Me.RefreshToolStripMenuItem.Size = New System.Drawing.Size(212, 22)
        Me.RefreshToolStripMenuItem.Text = "Refresh"
        '
        'ExitToolStripMenuItem1
        '
        Me.ExitToolStripMenuItem1.Image = CType(resources.GetObject("ExitToolStripMenuItem1.Image"), System.Drawing.Image)
        Me.ExitToolStripMenuItem1.Name = "ExitToolStripMenuItem1"
        Me.ExitToolStripMenuItem1.Size = New System.Drawing.Size(152, 22)
        Me.ExitToolStripMenuItem1.Text = "Exit"
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DeactivateToolStripMenuItem, Me.ByProductCategoryToolStripMenuItem})
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(54, 20)
        Me.EditToolStripMenuItem.Text = "Search"
        '
        'DeactivateToolStripMenuItem
        '
        Me.DeactivateToolStripMenuItem.Name = "DeactivateToolStripMenuItem"
        Me.DeactivateToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.DeactivateToolStripMenuItem.Text = "# by Part No"
        '
        'lblDisplay
        '
        Me.lblDisplay.AutoSize = True
        Me.lblDisplay.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDisplay.Location = New System.Drawing.Point(22, 32)
        Me.lblDisplay.Name = "lblDisplay"
        Me.lblDisplay.Size = New System.Drawing.Size(340, 23)
        Me.lblDisplay.TabIndex = 111
        Me.lblDisplay.Text = "Unique Motors (Pvt) Ltd - Magammana"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(22, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(138, 19)
        Me.Label2.TabIndex = 112
        Me.Label2.Text = "Create Discount Item"
        '
        'UltraGrid1
        '
        Me.UltraGrid1.Location = New System.Drawing.Point(26, 77)
        Me.UltraGrid1.Name = "UltraGrid1"
        Me.UltraGrid1.Size = New System.Drawing.Size(883, 429)
        Me.UltraGrid1.TabIndex = 113
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(212, 22)
        Me.ToolStripMenuItem1.Text = "Create New Discount Item"
        '
        'ByProductCategoryToolStripMenuItem
        '
        Me.ByProductCategoryToolStripMenuItem.Name = "ByProductCategoryToolStripMenuItem"
        Me.ByProductCategoryToolStripMenuItem.Size = New System.Drawing.Size(193, 22)
        Me.ByProductCategoryToolStripMenuItem.Text = "# by Product Category"
        '
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.txtDiscount)
        Me.OPR0.Controls.Add(Me.UltraLabel2)
        Me.OPR0.Controls.Add(Me.txtDescription)
        Me.OPR0.Controls.Add(Me.UltraLabel1)
        Me.OPR0.Controls.Add(Me.txtCode)
        Me.OPR0.Controls.Add(Me.UltraLabel17)
        Me.OPR0.Location = New System.Drawing.Point(169, 113)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(627, 101)
        Me.OPR0.TabIndex = 114
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
        'frmDiscount_item_Uniq
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(982, 525)
        Me.Controls.Add(Me.OPR0)
        Me.Controls.Add(Me.UltraGrid1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblDisplay)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Name = "frmDiscount_item_Uniq"
        Me.Text = "#Discount Item"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.txtDiscount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDescription, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCode, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ReportFilterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RefreshToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DeactivateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lblDisplay As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ByProductCategoryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtDiscount As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtDescription As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtCode As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel17 As Infragistics.Win.Misc.UltraLabel
End Class
