<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGrn_T_Acc
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
        Dim Appearance99 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance100 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance101 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance102 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance103 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance104 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance105 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance106 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance107 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance108 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance109 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance110 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance111 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGrn_T_Acc))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.ReportFilterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.DetailsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.BySupplierToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ByDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.UltraButton8 = New Infragistics.Win.Misc.UltraButton
        Me.cboAccount = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.Label8 = New System.Windows.Forms.Label
        Me.PrintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ResetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MenuStrip1.SuspendLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        CType(Me.cboAccount, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReportFilterToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1077, 24)
        Me.MenuStrip1.TabIndex = 2
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ReportFilterToolStripMenuItem
        '
        Me.ReportFilterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DetailsToolStripMenuItem, Me.PrintToolStripMenuItem, Me.ResetToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.ReportFilterToolStripMenuItem.Name = "ReportFilterToolStripMenuItem"
        Me.ReportFilterToolStripMenuItem.Size = New System.Drawing.Size(45, 20)
        Me.ReportFilterToolStripMenuItem.Text = "Filter"
        '
        'DetailsToolStripMenuItem
        '
        Me.DetailsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BySupplierToolStripMenuItem, Me.ByDateToolStripMenuItem})
        Me.DetailsToolStripMenuItem.Name = "DetailsToolStripMenuItem"
        Me.DetailsToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.DetailsToolStripMenuItem.Text = "by Supplier"
        '
        'BySupplierToolStripMenuItem
        '
        Me.BySupplierToolStripMenuItem.Name = "BySupplierToolStripMenuItem"
        Me.BySupplierToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.BySupplierToolStripMenuItem.Text = "All Transaction"
        '
        'ByDateToolStripMenuItem
        '
        Me.ByDateToolStripMenuItem.Name = "ByDateToolStripMenuItem"
        Me.ByDateToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.ByDateToolStripMenuItem.Text = "by Date"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(17, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(138, 23)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "SS Distributors"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(17, 57)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(123, 19)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Creditors Statment"
        '
        'UltraGrid1
        '
        Me.UltraGrid1.Location = New System.Drawing.Point(12, 93)
        Me.UltraGrid1.Name = "UltraGrid1"
        Me.UltraGrid1.Size = New System.Drawing.Size(1045, 386)
        Me.UltraGrid1.TabIndex = 22
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.UltraButton8)
        Me.Panel3.Controls.Add(Me.cboAccount)
        Me.Panel3.Controls.Add(Me.Label8)
        Me.Panel3.Location = New System.Drawing.Point(315, 124)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(437, 64)
        Me.Panel3.TabIndex = 123
        Me.Panel3.Visible = False
        '
        'UltraButton8
        '
        Appearance99.Image = CType(resources.GetObject("Appearance99.Image"), Object)
        Appearance99.ImageHAlign = Infragistics.Win.HAlign.Center
        Me.UltraButton8.Appearance = Appearance99
        Me.UltraButton8.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton8.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton8.Location = New System.Drawing.Point(392, 16)
        Me.UltraButton8.Name = "UltraButton8"
        Me.UltraButton8.Size = New System.Drawing.Size(33, 34)
        Me.UltraButton8.TabIndex = 278
        '
        'cboAccount
        '
        Me.cboAccount.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Suggest
        Me.cboAccount.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Appearance100.BackColor = System.Drawing.SystemColors.Window
        Appearance100.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.cboAccount.DisplayLayout.Appearance = Appearance100
        Me.cboAccount.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.cboAccount.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance101.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance101.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance101.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance101.BorderColor = System.Drawing.SystemColors.Window
        Me.cboAccount.DisplayLayout.GroupByBox.Appearance = Appearance101
        Appearance102.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboAccount.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance102
        Me.cboAccount.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance103.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance103.BackColor2 = System.Drawing.SystemColors.Control
        Appearance103.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance103.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboAccount.DisplayLayout.GroupByBox.PromptAppearance = Appearance103
        Me.cboAccount.DisplayLayout.MaxColScrollRegions = 1
        Me.cboAccount.DisplayLayout.MaxRowScrollRegions = 1
        Appearance104.BackColor = System.Drawing.SystemColors.Window
        Appearance104.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cboAccount.DisplayLayout.Override.ActiveCellAppearance = Appearance104
        Appearance105.BackColor = System.Drawing.SystemColors.Highlight
        Appearance105.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.cboAccount.DisplayLayout.Override.ActiveRowAppearance = Appearance105
        Me.cboAccount.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.cboAccount.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance106.BackColor = System.Drawing.SystemColors.Window
        Me.cboAccount.DisplayLayout.Override.CardAreaAppearance = Appearance106
        Appearance107.BorderColor = System.Drawing.Color.Silver
        Appearance107.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.cboAccount.DisplayLayout.Override.CellAppearance = Appearance107
        Me.cboAccount.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.cboAccount.DisplayLayout.Override.CellPadding = 0
        Appearance108.BackColor = System.Drawing.SystemColors.Control
        Appearance108.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance108.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance108.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance108.BorderColor = System.Drawing.SystemColors.Window
        Me.cboAccount.DisplayLayout.Override.GroupByRowAppearance = Appearance108
        Appearance109.TextHAlignAsString = "Left"
        Me.cboAccount.DisplayLayout.Override.HeaderAppearance = Appearance109
        Me.cboAccount.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.cboAccount.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance110.BackColor = System.Drawing.SystemColors.Window
        Appearance110.BorderColor = System.Drawing.Color.Silver
        Me.cboAccount.DisplayLayout.Override.RowAppearance = Appearance110
        Me.cboAccount.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance111.BackColor = System.Drawing.SystemColors.ControlLight
        Me.cboAccount.DisplayLayout.Override.TemplateAddRowAppearance = Appearance111
        Me.cboAccount.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.cboAccount.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.cboAccount.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.cboAccount.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.[Default]
        Me.cboAccount.Location = New System.Drawing.Point(79, 21)
        Me.cboAccount.Name = "cboAccount"
        Me.cboAccount.Size = New System.Drawing.Size(307, 22)
        Me.cboAccount.TabIndex = 277
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 24)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(76, 13)
        Me.Label8.TabIndex = 276
        Me.Label8.Text = "Supplier Name"
        '
        'PrintToolStripMenuItem
        '
        Me.PrintToolStripMenuItem.Image = CType(resources.GetObject("PrintToolStripMenuItem.Image"), System.Drawing.Image)
        Me.PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        Me.PrintToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.PrintToolStripMenuItem.Text = "Print Preview"
        '
        'ResetToolStripMenuItem
        '
        Me.ResetToolStripMenuItem.Image = CType(resources.GetObject("ResetToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ResetToolStripMenuItem.Name = "ResetToolStripMenuItem"
        Me.ResetToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.ResetToolStripMenuItem.Text = "Reset"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Image = CType(resources.GetObject("ExitToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'frmGrn_T_Acc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1077, 446)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.UltraGrid1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Name = "frmGrn_T_Acc"
        Me.Text = "##Creditors Statment"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.cboAccount, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ReportFilterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DetailsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BySupplierToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrintToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ResetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents UltraButton8 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cboAccount As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents ByDateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
