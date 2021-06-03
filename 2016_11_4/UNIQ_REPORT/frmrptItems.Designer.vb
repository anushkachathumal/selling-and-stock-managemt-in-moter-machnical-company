<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmrptItems
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmrptItems))
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance24 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance26 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance27 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.ReportFilterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UsingDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem
        Me.CategoryOnlyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem3 = New System.Windows.Forms.ToolStripMenuItem
        Me.RefreshToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ExitToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.lblDisplay = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.UltraGrid2 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.UltraButton1 = New Infragistics.Win.Misc.UltraButton
        Me.cboCategory = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.Label1 = New System.Windows.Forms.Label
        Me.MenuStrip1.SuspendLayout()
        CType(Me.UltraGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.cboCategory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReportFilterToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1031, 24)
        Me.MenuStrip1.TabIndex = 112
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ReportFilterToolStripMenuItem
        '
        Me.ReportFilterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UsingDateToolStripMenuItem, Me.ToolStripMenuItem2, Me.ToolStripMenuItem3, Me.RefreshToolStripMenuItem, Me.ExitToolStripMenuItem1})
        Me.ReportFilterToolStripMenuItem.Name = "ReportFilterToolStripMenuItem"
        Me.ReportFilterToolStripMenuItem.Size = New System.Drawing.Size(54, 20)
        Me.ReportFilterToolStripMenuItem.Text = "Search"
        '
        'UsingDateToolStripMenuItem
        '
        Me.UsingDateToolStripMenuItem.Name = "UsingDateToolStripMenuItem"
        Me.UsingDateToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.UsingDateToolStripMenuItem.Text = "# All Items"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CategoryOnlyToolStripMenuItem})
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(159, 22)
        Me.ToolStripMenuItem2.Text = "# Item Category"
        '
        'CategoryOnlyToolStripMenuItem
        '
        Me.CategoryOnlyToolStripMenuItem.Name = "CategoryOnlyToolStripMenuItem"
        Me.CategoryOnlyToolStripMenuItem.Size = New System.Drawing.Size(254, 22)
        Me.CategoryOnlyToolStripMenuItem.Text = "Category Only"
        '
        'ToolStripMenuItem3
        '
        Me.ToolStripMenuItem3.Image = CType(resources.GetObject("ToolStripMenuItem3.Image"), System.Drawing.Image)
        Me.ToolStripMenuItem3.Name = "ToolStripMenuItem3"
        Me.ToolStripMenuItem3.Size = New System.Drawing.Size(159, 22)
        Me.ToolStripMenuItem3.Text = "Print"
        '
        'RefreshToolStripMenuItem
        '
        Me.RefreshToolStripMenuItem.Image = CType(resources.GetObject("RefreshToolStripMenuItem.Image"), System.Drawing.Image)
        Me.RefreshToolStripMenuItem.Name = "RefreshToolStripMenuItem"
        Me.RefreshToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.RefreshToolStripMenuItem.Text = "Refresh"
        '
        'ExitToolStripMenuItem1
        '
        Me.ExitToolStripMenuItem1.Image = CType(resources.GetObject("ExitToolStripMenuItem1.Image"), System.Drawing.Image)
        Me.ExitToolStripMenuItem1.Name = "ExitToolStripMenuItem1"
        Me.ExitToolStripMenuItem1.Size = New System.Drawing.Size(159, 22)
        Me.ExitToolStripMenuItem1.Text = "Exit"
        '
        'lblDisplay
        '
        Me.lblDisplay.AutoSize = True
        Me.lblDisplay.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDisplay.Location = New System.Drawing.Point(12, 39)
        Me.lblDisplay.Name = "lblDisplay"
        Me.lblDisplay.Size = New System.Drawing.Size(340, 23)
        Me.lblDisplay.TabIndex = 115
        Me.lblDisplay.Text = "Unique Motors (Pvt) Ltd - Magammana"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 62)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(82, 19)
        Me.Label2.TabIndex = 116
        Me.Label2.Text = "Item Report"
        '
        'UltraGrid2
        '
        Me.UltraGrid2.Location = New System.Drawing.Point(16, 100)
        Me.UltraGrid2.Name = "UltraGrid2"
        Me.UltraGrid2.Size = New System.Drawing.Size(978, 389)
        Me.UltraGrid2.TabIndex = 122
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.UltraButton1)
        Me.Panel1.Controls.Add(Me.cboCategory)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(304, 149)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(388, 62)
        Me.Panel1.TabIndex = 128
        Me.Panel1.Visible = False
        '
        'UltraButton1
        '
        Appearance3.Image = CType(resources.GetObject("Appearance3.Image"), Object)
        Appearance3.ImageHAlign = Infragistics.Win.HAlign.Center
        Me.UltraButton1.Appearance = Appearance3
        Me.UltraButton1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton1.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton1.Location = New System.Drawing.Point(348, 15)
        Me.UltraButton1.Name = "UltraButton1"
        Me.UltraButton1.Size = New System.Drawing.Size(33, 34)
        Me.UltraButton1.TabIndex = 280
        '
        'cboCategory
        '
        Me.cboCategory.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Suggest
        Me.cboCategory.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Appearance4.BackColor = System.Drawing.SystemColors.Window
        Appearance4.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.cboCategory.DisplayLayout.Appearance = Appearance4
        Me.cboCategory.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.cboCategory.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance8.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance8.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance8.BorderColor = System.Drawing.SystemColors.Window
        Me.cboCategory.DisplayLayout.GroupByBox.Appearance = Appearance8
        Appearance18.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboCategory.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance18
        Me.cboCategory.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance19.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance19.BackColor2 = System.Drawing.SystemColors.Control
        Appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance19.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboCategory.DisplayLayout.GroupByBox.PromptAppearance = Appearance19
        Me.cboCategory.DisplayLayout.MaxColScrollRegions = 1
        Me.cboCategory.DisplayLayout.MaxRowScrollRegions = 1
        Appearance20.BackColor = System.Drawing.SystemColors.Window
        Appearance20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cboCategory.DisplayLayout.Override.ActiveCellAppearance = Appearance20
        Appearance21.BackColor = System.Drawing.SystemColors.Highlight
        Appearance21.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.cboCategory.DisplayLayout.Override.ActiveRowAppearance = Appearance21
        Me.cboCategory.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.cboCategory.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance22.BackColor = System.Drawing.SystemColors.Window
        Me.cboCategory.DisplayLayout.Override.CardAreaAppearance = Appearance22
        Appearance23.BorderColor = System.Drawing.Color.Silver
        Appearance23.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.cboCategory.DisplayLayout.Override.CellAppearance = Appearance23
        Me.cboCategory.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.cboCategory.DisplayLayout.Override.CellPadding = 0
        Appearance24.BackColor = System.Drawing.SystemColors.Control
        Appearance24.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance24.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance24.BorderColor = System.Drawing.SystemColors.Window
        Me.cboCategory.DisplayLayout.Override.GroupByRowAppearance = Appearance24
        Appearance25.TextHAlignAsString = "Left"
        Me.cboCategory.DisplayLayout.Override.HeaderAppearance = Appearance25
        Me.cboCategory.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.cboCategory.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance26.BackColor = System.Drawing.SystemColors.Window
        Appearance26.BorderColor = System.Drawing.Color.Silver
        Me.cboCategory.DisplayLayout.Override.RowAppearance = Appearance26
        Me.cboCategory.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance27.BackColor = System.Drawing.SystemColors.ControlLight
        Me.cboCategory.DisplayLayout.Override.TemplateAddRowAppearance = Appearance27
        Me.cboCategory.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.cboCategory.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.cboCategory.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.cboCategory.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2007
        Me.cboCategory.Location = New System.Drawing.Point(67, 20)
        Me.cboCategory.Name = "cboCategory"
        Me.cboCategory.Size = New System.Drawing.Size(275, 22)
        Me.cboCategory.TabIndex = 279
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 13)
        Me.Label1.TabIndex = 276
        Me.Label1.Text = "Category"
        '
        'frmrptItems
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1031, 484)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.UltraGrid2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblDisplay)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Name = "frmrptItems"
        Me.Text = "#Report Viwer"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.UltraGrid2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.cboCategory, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ReportFilterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UsingDateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CategoryOnlyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem3 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RefreshToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lblDisplay As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents UltraGrid2 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents UltraButton1 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cboCategory As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
