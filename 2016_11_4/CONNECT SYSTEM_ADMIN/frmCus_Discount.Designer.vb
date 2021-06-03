<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCus_Discount
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
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance27 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCus_Discount))
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.lblDisplay = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtTotal_Discount = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel3 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraLabel8 = New Infragistics.Win.Misc.UltraLabel
        Me.txtCash = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraButton2 = New Infragistics.Win.Misc.UltraButton
        Me.cmdAdd = New Infragistics.Win.Misc.UltraButton
        Me.cmdReset = New Infragistics.Win.Misc.UltraButton
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtDiscount = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.txtDays = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraGrid2 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.ReportFilterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UsingDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CustomerSpecificDiscountToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CompanyDiscountToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.RefreshToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ExitToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.txtTotal_Discount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCash, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox1.SuspendLayout()
        CType(Me.txtDiscount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDays, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblDisplay
        '
        Me.lblDisplay.AutoSize = True
        Me.lblDisplay.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDisplay.Location = New System.Drawing.Point(22, 32)
        Me.lblDisplay.Name = "lblDisplay"
        Me.lblDisplay.Size = New System.Drawing.Size(333, 23)
        Me.lblDisplay.TabIndex = 107
        Me.lblDisplay.Text = "Connect Electricals (Pvt) Ltd - Kegalle"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(24, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(131, 19)
        Me.Label2.TabIndex = 108
        Me.Label2.Text = "Customers Discount"
        '
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.txtTotal_Discount)
        Me.OPR0.Controls.Add(Me.UltraLabel3)
        Me.OPR0.Controls.Add(Me.UltraLabel8)
        Me.OPR0.Controls.Add(Me.txtCash)
        Me.OPR0.Location = New System.Drawing.Point(28, 83)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(470, 65)
        Me.OPR0.TabIndex = 110
        Me.OPR0.Text = "Cash Discount ...."
        '
        'txtTotal_Discount
        '
        Appearance4.ForeColor = System.Drawing.Color.Black
        Me.txtTotal_Discount.Appearance = Appearance4
        Me.txtTotal_Discount.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotal_Discount.Location = New System.Drawing.Point(357, 28)
        Me.txtTotal_Discount.MaxLength = 30
        Me.txtTotal_Discount.Name = "txtTotal_Discount"
        Me.txtTotal_Discount.Size = New System.Drawing.Size(101, 23)
        Me.txtTotal_Discount.TabIndex = 50
        '
        'UltraLabel3
        '
        Appearance5.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel3.Appearance = Appearance5
        Me.UltraLabel3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel3.Location = New System.Drawing.Point(220, 30)
        Me.UltraLabel3.Name = "UltraLabel3"
        Me.UltraLabel3.Size = New System.Drawing.Size(131, 21)
        Me.UltraLabel3.TabIndex = 49
        Me.UltraLabel3.Text = "Total Invoice Discount%"
        Me.UltraLabel3.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'UltraLabel8
        '
        Appearance7.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel8.Appearance = Appearance7
        Me.UltraLabel8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel8.Location = New System.Drawing.Point(6, 30)
        Me.UltraLabel8.Name = "UltraLabel8"
        Me.UltraLabel8.Size = New System.Drawing.Size(107, 21)
        Me.UltraLabel8.TabIndex = 48
        Me.UltraLabel8.Text = "Cash Discount (%)"
        Me.UltraLabel8.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtCash
        '
        Appearance8.ForeColor = System.Drawing.Color.Black
        Me.txtCash.Appearance = Appearance8
        Me.txtCash.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCash.Location = New System.Drawing.Point(114, 28)
        Me.txtCash.MaxLength = 30
        Me.txtCash.Name = "txtCash"
        Me.txtCash.Size = New System.Drawing.Size(101, 23)
        Me.txtCash.TabIndex = 1
        '
        'UltraButton2
        '
        Appearance27.Image = CType(resources.GetObject("Appearance27.Image"), Object)
        Me.UltraButton2.Appearance = Appearance27
        Me.UltraButton2.Enabled = False
        Me.UltraButton2.ImageSize = New System.Drawing.Size(32, 32)
        Me.UltraButton2.Location = New System.Drawing.Point(147, 225)
        Me.UltraButton2.Name = "UltraButton2"
        Me.UltraButton2.Size = New System.Drawing.Size(113, 47)
        Me.UltraButton2.TabIndex = 107
        Me.UltraButton2.Text = "&Cancel"
        '
        'cmdAdd
        '
        Appearance1.Image = CType(resources.GetObject("Appearance1.Image"), Object)
        Appearance1.TextHAlignAsString = "Center"
        Me.cmdAdd.Appearance = Appearance1
        Me.cmdAdd.ImageSize = New System.Drawing.Size(32, 32)
        Me.cmdAdd.Location = New System.Drawing.Point(28, 225)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(113, 47)
        Me.cmdAdd.TabIndex = 104
        Me.cmdAdd.Text = "&Save"
        '
        'cmdReset
        '
        Appearance2.Image = CType(resources.GetObject("Appearance2.Image"), Object)
        Me.cmdReset.Appearance = Appearance2
        Me.cmdReset.ImageSize = New System.Drawing.Size(32, 32)
        Me.cmdReset.Location = New System.Drawing.Point(266, 225)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.Size = New System.Drawing.Size(113, 47)
        Me.cmdReset.TabIndex = 106
        Me.cmdReset.Text = "&Reset"
        '
        'cmdExit
        '
        Appearance17.Image = CType(resources.GetObject("Appearance17.Image"), Object)
        Me.cmdExit.Appearance = Appearance17
        Me.cmdExit.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdExit.Location = New System.Drawing.Point(385, 225)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(113, 47)
        Me.cmdExit.TabIndex = 105
        Me.cmdExit.Text = "&Exit"
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.Controls.Add(Me.txtDiscount)
        Me.UltraGroupBox1.Controls.Add(Me.UltraLabel2)
        Me.UltraGroupBox1.Controls.Add(Me.UltraLabel1)
        Me.UltraGroupBox1.Controls.Add(Me.txtDays)
        Me.UltraGroupBox1.Location = New System.Drawing.Point(28, 154)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(470, 65)
        Me.UltraGroupBox1.TabIndex = 111
        Me.UltraGroupBox1.Text = "Chque Discount ...."
        '
        'txtDiscount
        '
        Appearance3.ForeColor = System.Drawing.Color.Black
        Me.txtDiscount.Appearance = Appearance3
        Me.txtDiscount.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDiscount.Location = New System.Drawing.Point(274, 28)
        Me.txtDiscount.MaxLength = 30
        Me.txtDiscount.Name = "txtDiscount"
        Me.txtDiscount.Size = New System.Drawing.Size(101, 23)
        Me.txtDiscount.TabIndex = 50
        '
        'UltraLabel2
        '
        Appearance18.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel2.Appearance = Appearance18
        Me.UltraLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel2.Location = New System.Drawing.Point(176, 30)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel2.TabIndex = 49
        Me.UltraLabel2.Text = "Discount (%)"
        Me.UltraLabel2.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'UltraLabel1
        '
        Appearance6.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel1.Appearance = Appearance6
        Me.UltraLabel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel1.Location = New System.Drawing.Point(6, 30)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel1.TabIndex = 48
        Me.UltraLabel1.Text = "Days"
        Me.UltraLabel1.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtDays
        '
        Appearance21.ForeColor = System.Drawing.Color.Black
        Me.txtDays.Appearance = Appearance21
        Me.txtDays.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDays.Location = New System.Drawing.Point(114, 28)
        Me.txtDays.MaxLength = 30
        Me.txtDays.Name = "txtDays"
        Me.txtDays.Size = New System.Drawing.Size(56, 23)
        Me.txtDays.TabIndex = 1
        '
        'UltraGrid2
        '
        Me.UltraGrid2.Location = New System.Drawing.Point(28, 278)
        Me.UltraGrid2.Name = "UltraGrid2"
        Me.UltraGrid2.Size = New System.Drawing.Size(468, 218)
        Me.UltraGrid2.TabIndex = 112
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReportFilterToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(906, 24)
        Me.MenuStrip1.TabIndex = 113
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ReportFilterToolStripMenuItem
        '
        Me.ReportFilterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UsingDateToolStripMenuItem, Me.RefreshToolStripMenuItem, Me.ExitToolStripMenuItem1})
        Me.ReportFilterToolStripMenuItem.Name = "ReportFilterToolStripMenuItem"
        Me.ReportFilterToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.ReportFilterToolStripMenuItem.Text = "File"
        '
        'UsingDateToolStripMenuItem
        '
        Me.UsingDateToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CustomerSpecificDiscountToolStripMenuItem, Me.CompanyDiscountToolStripMenuItem})
        Me.UsingDateToolStripMenuItem.Name = "UsingDateToolStripMenuItem"
        Me.UsingDateToolStripMenuItem.Size = New System.Drawing.Size(170, 22)
        Me.UsingDateToolStripMenuItem.Text = "Advance Discount"
        '
        'CustomerSpecificDiscountToolStripMenuItem
        '
        Me.CustomerSpecificDiscountToolStripMenuItem.Name = "CustomerSpecificDiscountToolStripMenuItem"
        Me.CustomerSpecificDiscountToolStripMenuItem.Size = New System.Drawing.Size(220, 22)
        Me.CustomerSpecificDiscountToolStripMenuItem.Text = "Customer Specific Discount"
        '
        'CompanyDiscountToolStripMenuItem
        '
        Me.CompanyDiscountToolStripMenuItem.Name = "CompanyDiscountToolStripMenuItem"
        Me.CompanyDiscountToolStripMenuItem.Size = New System.Drawing.Size(220, 22)
        Me.CompanyDiscountToolStripMenuItem.Text = "Company Discount"
        '
        'RefreshToolStripMenuItem
        '
        Me.RefreshToolStripMenuItem.Image = CType(resources.GetObject("RefreshToolStripMenuItem.Image"), System.Drawing.Image)
        Me.RefreshToolStripMenuItem.Name = "RefreshToolStripMenuItem"
        Me.RefreshToolStripMenuItem.Size = New System.Drawing.Size(170, 22)
        Me.RefreshToolStripMenuItem.Text = "Refresh"
        '
        'ExitToolStripMenuItem1
        '
        Me.ExitToolStripMenuItem1.Image = CType(resources.GetObject("ExitToolStripMenuItem1.Image"), System.Drawing.Image)
        Me.ExitToolStripMenuItem1.Name = "ExitToolStripMenuItem1"
        Me.ExitToolStripMenuItem1.Size = New System.Drawing.Size(170, 22)
        Me.ExitToolStripMenuItem1.Text = "Exit"
        '
        'frmCus_Discount
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(906, 648)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.UltraGrid2)
        Me.Controls.Add(Me.UltraGroupBox1)
        Me.Controls.Add(Me.OPR0)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdReset)
        Me.Controls.Add(Me.UltraButton2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.lblDisplay)
        Me.Name = "frmCus_Discount"
        Me.Text = "Customer Discount"
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.txtTotal_Discount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCash, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox1.ResumeLayout(False)
        Me.UltraGroupBox1.PerformLayout()
        CType(Me.txtDiscount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDays, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGrid2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblDisplay As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraButton2 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdReset As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraLabel8 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtCash As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents cmdAdd As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtDiscount As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtDays As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraGrid2 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents txtTotal_Discount As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel3 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ReportFilterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UsingDateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CustomerSpecificDiscountToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RefreshToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CompanyDiscountToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
