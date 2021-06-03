<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmrptSupplier
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmrptSupplier))
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance31 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance30 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance45 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance32 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance29 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance28 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.ReportFilterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ActiveSupplierToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.InactiveSupplierToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PrintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ResetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtType = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.txtName = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.txtStatus = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.txtVAT = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel8 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraLabel7 = New Infragistics.Win.Misc.UltraLabel
        Me.txtContact = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel6 = New Infragistics.Win.Misc.UltraLabel
        Me.txtFax = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel5 = New Infragistics.Win.Misc.UltraLabel
        Me.txtTp = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel3 = New Infragistics.Win.Misc.UltraLabel
        Me.txtAdd1 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.txtAddress = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel
        Me.txtCode = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.MenuStrip1.SuspendLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.txtType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtVAT, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtContact, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtFax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAdd1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAddress, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCode, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReportFilterToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(889, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ReportFilterToolStripMenuItem
        '
        Me.ReportFilterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ActiveSupplierToolStripMenuItem, Me.InactiveSupplierToolStripMenuItem, Me.PrintToolStripMenuItem, Me.ExitToolStripMenuItem, Me.ResetToolStripMenuItem})
        Me.ReportFilterToolStripMenuItem.Name = "ReportFilterToolStripMenuItem"
        Me.ReportFilterToolStripMenuItem.Size = New System.Drawing.Size(83, 20)
        Me.ReportFilterToolStripMenuItem.Text = "Report Filter"
        '
        'ActiveSupplierToolStripMenuItem
        '
        Me.ActiveSupplierToolStripMenuItem.Name = "ActiveSupplierToolStripMenuItem"
        Me.ActiveSupplierToolStripMenuItem.Size = New System.Drawing.Size(161, 22)
        Me.ActiveSupplierToolStripMenuItem.Text = "Active Supplier"
        '
        'InactiveSupplierToolStripMenuItem
        '
        Me.InactiveSupplierToolStripMenuItem.Name = "InactiveSupplierToolStripMenuItem"
        Me.InactiveSupplierToolStripMenuItem.Size = New System.Drawing.Size(161, 22)
        Me.InactiveSupplierToolStripMenuItem.Text = "Inactive Supplier"
        '
        'PrintToolStripMenuItem
        '
        Me.PrintToolStripMenuItem.Image = CType(resources.GetObject("PrintToolStripMenuItem.Image"), System.Drawing.Image)
        Me.PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        Me.PrintToolStripMenuItem.Size = New System.Drawing.Size(161, 22)
        Me.PrintToolStripMenuItem.Text = "Print"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Image = CType(resources.GetObject("ExitToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(161, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'ResetToolStripMenuItem
        '
        Me.ResetToolStripMenuItem.Image = CType(resources.GetObject("ResetToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ResetToolStripMenuItem.Name = "ResetToolStripMenuItem"
        Me.ResetToolStripMenuItem.Size = New System.Drawing.Size(161, 22)
        Me.ResetToolStripMenuItem.Text = "Reset"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 35)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(223, 23)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "SS Distributor-Moratuwa"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 58)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(101, 19)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "List of Supplier"
        '
        'UltraGrid1
        '
        Me.UltraGrid1.Location = New System.Drawing.Point(16, 80)
        Me.UltraGrid1.Name = "UltraGrid1"
        Me.UltraGrid1.Size = New System.Drawing.Size(834, 386)
        Me.UltraGrid1.TabIndex = 3
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.OPR0)
        Me.Panel1.Location = New System.Drawing.Point(112, 137)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(638, 231)
        Me.Panel1.TabIndex = 4
        Me.Panel1.Visible = False
        '
        'OPR0
        '
        Appearance3.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.OPR0.Appearance = Appearance3
        Me.OPR0.Controls.Add(Me.txtType)
        Me.OPR0.Controls.Add(Me.txtName)
        Me.OPR0.Controls.Add(Me.txtStatus)
        Me.OPR0.Controls.Add(Me.txtVAT)
        Me.OPR0.Controls.Add(Me.UltraLabel8)
        Me.OPR0.Controls.Add(Me.UltraLabel7)
        Me.OPR0.Controls.Add(Me.txtContact)
        Me.OPR0.Controls.Add(Me.UltraLabel6)
        Me.OPR0.Controls.Add(Me.txtFax)
        Me.OPR0.Controls.Add(Me.UltraLabel5)
        Me.OPR0.Controls.Add(Me.txtTp)
        Me.OPR0.Controls.Add(Me.UltraLabel3)
        Me.OPR0.Controls.Add(Me.txtAdd1)
        Me.OPR0.Controls.Add(Me.txtAddress)
        Me.OPR0.Controls.Add(Me.UltraLabel1)
        Me.OPR0.Controls.Add(Me.UltraLabel4)
        Me.OPR0.Controls.Add(Me.txtCode)
        Me.OPR0.Controls.Add(Me.UltraLabel2)
        Me.OPR0.Location = New System.Drawing.Point(-1, -1)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(638, 232)
        Me.OPR0.TabIndex = 261
        Me.OPR0.Visible = False
        '
        'txtType
        '
        Me.txtType.Location = New System.Drawing.Point(113, 178)
        Me.txtType.MaxLength = 15
        Me.txtType.Name = "txtType"
        Me.txtType.Size = New System.Drawing.Size(180, 21)
        Me.txtType.TabIndex = 127
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(195, 42)
        Me.txtName.MaxLength = 15
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(429, 21)
        Me.txtName.TabIndex = 126
        '
        'txtStatus
        '
        Me.txtStatus.Location = New System.Drawing.Point(113, 41)
        Me.txtStatus.MaxLength = 15
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.Size = New System.Drawing.Size(76, 21)
        Me.txtStatus.TabIndex = 125
        '
        'txtVAT
        '
        Me.txtVAT.Location = New System.Drawing.Point(367, 151)
        Me.txtVAT.MaxLength = 50
        Me.txtVAT.Name = "txtVAT"
        Me.txtVAT.Size = New System.Drawing.Size(180, 21)
        Me.txtVAT.TabIndex = 124
        '
        'UltraLabel8
        '
        Appearance31.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel8.Appearance = Appearance31
        Me.UltraLabel8.Location = New System.Drawing.Point(299, 152)
        Me.UltraLabel8.Name = "UltraLabel8"
        Me.UltraLabel8.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel8.TabIndex = 123
        Me.UltraLabel8.Text = "VAT Reg No"
        Me.UltraLabel8.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'UltraLabel7
        '
        Appearance30.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel7.Appearance = Appearance30
        Me.UltraLabel7.Location = New System.Drawing.Point(15, 180)
        Me.UltraLabel7.Name = "UltraLabel7"
        Me.UltraLabel7.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel7.TabIndex = 121
        Me.UltraLabel7.Text = "Supplier Type"
        Me.UltraLabel7.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtContact
        '
        Me.txtContact.Location = New System.Drawing.Point(113, 151)
        Me.txtContact.MaxLength = 50
        Me.txtContact.Name = "txtContact"
        Me.txtContact.Size = New System.Drawing.Size(180, 21)
        Me.txtContact.TabIndex = 120
        '
        'UltraLabel6
        '
        Appearance45.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel6.Appearance = Appearance45
        Me.UltraLabel6.Location = New System.Drawing.Point(15, 152)
        Me.UltraLabel6.Name = "UltraLabel6"
        Me.UltraLabel6.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel6.TabIndex = 119
        Me.UltraLabel6.Text = "Contact Person"
        Me.UltraLabel6.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtFax
        '
        Me.txtFax.Location = New System.Drawing.Point(367, 123)
        Me.txtFax.MaxLength = 50
        Me.txtFax.Name = "txtFax"
        Me.txtFax.Size = New System.Drawing.Size(180, 21)
        Me.txtFax.TabIndex = 118
        '
        'UltraLabel5
        '
        Appearance1.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel5.Appearance = Appearance1
        Me.UltraLabel5.Location = New System.Drawing.Point(299, 124)
        Me.UltraLabel5.Name = "UltraLabel5"
        Me.UltraLabel5.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel5.TabIndex = 117
        Me.UltraLabel5.Text = "Fax No"
        Me.UltraLabel5.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtTp
        '
        Me.txtTp.Location = New System.Drawing.Point(113, 123)
        Me.txtTp.MaxLength = 50
        Me.txtTp.Name = "txtTp"
        Me.txtTp.Size = New System.Drawing.Size(180, 21)
        Me.txtTp.TabIndex = 116
        '
        'UltraLabel3
        '
        Appearance32.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel3.Appearance = Appearance32
        Me.UltraLabel3.Location = New System.Drawing.Point(15, 124)
        Me.UltraLabel3.Name = "UltraLabel3"
        Me.UltraLabel3.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel3.TabIndex = 115
        Me.UltraLabel3.Text = "Contact No"
        Me.UltraLabel3.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtAdd1
        '
        Me.txtAdd1.Location = New System.Drawing.Point(113, 97)
        Me.txtAdd1.MaxLength = 150
        Me.txtAdd1.Name = "txtAdd1"
        Me.txtAdd1.Size = New System.Drawing.Size(511, 21)
        Me.txtAdd1.TabIndex = 114
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(113, 70)
        Me.txtAddress.MaxLength = 150
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(511, 21)
        Me.txtAddress.TabIndex = 113
        '
        'UltraLabel1
        '
        Appearance29.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel1.Appearance = Appearance29
        Me.UltraLabel1.Location = New System.Drawing.Point(15, 70)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel1.TabIndex = 110
        Me.UltraLabel1.Text = "Address"
        Me.UltraLabel1.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'UltraLabel4
        '
        Appearance28.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel4.Appearance = Appearance28
        Me.UltraLabel4.Location = New System.Drawing.Point(15, 43)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel4.TabIndex = 56
        Me.UltraLabel4.Text = "Supplier Name"
        Me.UltraLabel4.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtCode
        '
        Me.txtCode.Location = New System.Drawing.Point(113, 15)
        Me.txtCode.MaxLength = 15
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(127, 21)
        Me.txtCode.TabIndex = 51
        '
        'UltraLabel2
        '
        Appearance2.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel2.Appearance = Appearance2
        Me.UltraLabel2.Location = New System.Drawing.Point(15, 15)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(109, 21)
        Me.UltraLabel2.TabIndex = 36
        Me.UltraLabel2.Text = "Acc Code"
        Me.UltraLabel2.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'frmrptSupplier
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(889, 492)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.UltraGrid1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmrptSupplier"
        Me.Text = "Report View"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.txtType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtStatus, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtVAT, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtContact, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtFax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAdd1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAddress, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCode, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ReportFilterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ActiveSupplierToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents InactiveSupplierToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents PrintToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ResetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtVAT As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel8 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraLabel7 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtContact As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel6 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtFax As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel5 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtTp As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel3 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtAdd1 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents txtAddress As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtCode As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtType As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents txtName As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents txtStatus As Infragistics.Win.UltraWinEditors.UltraTextEditor
End Class
