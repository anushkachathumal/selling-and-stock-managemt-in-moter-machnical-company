<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmrptRowMaterial
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmrptRowMaterial))
        Dim Appearance544 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance356 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance357 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance358 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance359 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance360 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance361 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance362 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance363 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance364 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance365 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance381 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance384 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.ReportFilterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AllItemsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ByCategoryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PrintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.RefreshToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.UltraButton3 = New Infragistics.Win.Misc.UltraButton
        Me.cboCategory = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.Label3 = New System.Windows.Forms.Label
        Me.MenuStrip1.SuspendLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.cboCategory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReportFilterToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(955, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ReportFilterToolStripMenuItem
        '
        Me.ReportFilterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AllItemsToolStripMenuItem, Me.ByCategoryToolStripMenuItem, Me.PrintToolStripMenuItem, Me.RefreshToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.ReportFilterToolStripMenuItem.Name = "ReportFilterToolStripMenuItem"
        Me.ReportFilterToolStripMenuItem.Size = New System.Drawing.Size(83, 20)
        Me.ReportFilterToolStripMenuItem.Text = "Report Filter"
        '
        'AllItemsToolStripMenuItem
        '
        Me.AllItemsToolStripMenuItem.Name = "AllItemsToolStripMenuItem"
        Me.AllItemsToolStripMenuItem.Size = New System.Drawing.Size(138, 22)
        Me.AllItemsToolStripMenuItem.Text = "All Items"
        '
        'ByCategoryToolStripMenuItem
        '
        Me.ByCategoryToolStripMenuItem.Name = "ByCategoryToolStripMenuItem"
        Me.ByCategoryToolStripMenuItem.Size = New System.Drawing.Size(138, 22)
        Me.ByCategoryToolStripMenuItem.Text = "By Category"
        '
        'PrintToolStripMenuItem
        '
        Me.PrintToolStripMenuItem.Image = CType(resources.GetObject("PrintToolStripMenuItem.Image"), System.Drawing.Image)
        Me.PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        Me.PrintToolStripMenuItem.Size = New System.Drawing.Size(138, 22)
        Me.PrintToolStripMenuItem.Text = "Print"
        '
        'RefreshToolStripMenuItem
        '
        Me.RefreshToolStripMenuItem.Image = CType(resources.GetObject("RefreshToolStripMenuItem.Image"), System.Drawing.Image)
        Me.RefreshToolStripMenuItem.Name = "RefreshToolStripMenuItem"
        Me.RefreshToolStripMenuItem.Size = New System.Drawing.Size(138, 22)
        Me.RefreshToolStripMenuItem.Text = "Refresh"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Image = CType(resources.GetObject("ExitToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(138, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(131, 19)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "List of Row material"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(182, 23)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Apsco Shoes Pvt Ltd"
        '
        'UltraGrid1
        '
        Me.UltraGrid1.Location = New System.Drawing.Point(16, 91)
        Me.UltraGrid1.Name = "UltraGrid1"
        Me.UltraGrid1.Size = New System.Drawing.Size(834, 386)
        Me.UltraGrid1.TabIndex = 5
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.UltraButton3)
        Me.Panel1.Controls.Add(Me.cboCategory)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Location = New System.Drawing.Point(209, 180)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(320, 59)
        Me.Panel1.TabIndex = 6
        Me.Panel1.Visible = False
        '
        'UltraButton3
        '
        Appearance544.Image = CType(resources.GetObject("Appearance544.Image"), Object)
        Appearance544.ImageHAlign = Infragistics.Win.HAlign.Center
        Me.UltraButton3.Appearance = Appearance544
        Me.UltraButton3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton3.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton3.Location = New System.Drawing.Point(275, 12)
        Me.UltraButton3.Name = "UltraButton3"
        Me.UltraButton3.Size = New System.Drawing.Size(33, 34)
        Me.UltraButton3.TabIndex = 261
        '
        'cboCategory
        '
        Me.cboCategory.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Suggest
        Me.cboCategory.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Appearance356.BackColor = System.Drawing.SystemColors.Window
        Appearance356.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.cboCategory.DisplayLayout.Appearance = Appearance356
        Me.cboCategory.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.cboCategory.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance357.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance357.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance357.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance357.BorderColor = System.Drawing.SystemColors.Window
        Me.cboCategory.DisplayLayout.GroupByBox.Appearance = Appearance357
        Appearance358.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboCategory.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance358
        Me.cboCategory.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance359.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance359.BackColor2 = System.Drawing.SystemColors.Control
        Appearance359.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance359.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboCategory.DisplayLayout.GroupByBox.PromptAppearance = Appearance359
        Me.cboCategory.DisplayLayout.MaxColScrollRegions = 1
        Me.cboCategory.DisplayLayout.MaxRowScrollRegions = 1
        Appearance360.BackColor = System.Drawing.SystemColors.Window
        Appearance360.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cboCategory.DisplayLayout.Override.ActiveCellAppearance = Appearance360
        Appearance361.BackColor = System.Drawing.SystemColors.Highlight
        Appearance361.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.cboCategory.DisplayLayout.Override.ActiveRowAppearance = Appearance361
        Me.cboCategory.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.cboCategory.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance362.BackColor = System.Drawing.SystemColors.Window
        Me.cboCategory.DisplayLayout.Override.CardAreaAppearance = Appearance362
        Appearance363.BorderColor = System.Drawing.Color.Silver
        Appearance363.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.cboCategory.DisplayLayout.Override.CellAppearance = Appearance363
        Me.cboCategory.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.cboCategory.DisplayLayout.Override.CellPadding = 0
        Appearance364.BackColor = System.Drawing.SystemColors.Control
        Appearance364.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance364.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance364.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance364.BorderColor = System.Drawing.SystemColors.Window
        Me.cboCategory.DisplayLayout.Override.GroupByRowAppearance = Appearance364
        Appearance365.TextHAlignAsString = "Left"
        Me.cboCategory.DisplayLayout.Override.HeaderAppearance = Appearance365
        Me.cboCategory.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.cboCategory.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance381.BackColor = System.Drawing.SystemColors.Window
        Appearance381.BorderColor = System.Drawing.Color.Silver
        Me.cboCategory.DisplayLayout.Override.RowAppearance = Appearance381
        Me.cboCategory.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance384.BackColor = System.Drawing.SystemColors.ControlLight
        Me.cboCategory.DisplayLayout.Override.TemplateAddRowAppearance = Appearance384
        Me.cboCategory.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.cboCategory.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.cboCategory.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.cboCategory.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.[Default]
        Me.cboCategory.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList
        Me.cboCategory.Location = New System.Drawing.Point(89, 18)
        Me.cboCategory.Name = "cboCategory"
        Me.cboCategory.Size = New System.Drawing.Size(180, 22)
        Me.cboCategory.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 21)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Category Name"
        '
        'frmrptRowMaterial
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(955, 558)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.UltraGrid1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmrptRowMaterial"
        Me.Text = "List of Row Material"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.cboCategory, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ReportFilterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AllItemsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ByCategoryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrintToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RefreshToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboCategory As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents UltraButton3 As Infragistics.Win.Misc.UltraButton
End Class
