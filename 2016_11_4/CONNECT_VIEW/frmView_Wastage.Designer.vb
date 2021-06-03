<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmView_Wastage
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmView_Wastage))
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim DateButton13 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim DateButton14 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
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
        Dim DateButton15 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim DateButton16 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.ReportFilterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UsingDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ByDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ByPartNoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.ByDateToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.BySupplierToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.RefreshToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ExitToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.lblDisplay = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.UltraGrid2 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.UltraButton2 = New Infragistics.Win.Misc.UltraButton
        Me.txtDate2 = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtDate1 = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label8 = New System.Windows.Forms.Label
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.UltraButton3 = New Infragistics.Win.Misc.UltraButton
        Me.cboItem = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtC2 = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtC1 = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label9 = New System.Windows.Forms.Label
        Me.MenuStrip1.SuspendLayout()
        CType(Me.UltraGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.txtDate2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDate1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        CType(Me.cboItem, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReportFilterToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(865, 24)
        Me.MenuStrip1.TabIndex = 111
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ReportFilterToolStripMenuItem
        '
        Me.ReportFilterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UsingDateToolStripMenuItem, Me.ToolStripMenuItem1, Me.RefreshToolStripMenuItem, Me.ExitToolStripMenuItem1})
        Me.ReportFilterToolStripMenuItem.Name = "ReportFilterToolStripMenuItem"
        Me.ReportFilterToolStripMenuItem.Size = New System.Drawing.Size(54, 20)
        Me.ReportFilterToolStripMenuItem.Text = "Search"
        '
        'UsingDateToolStripMenuItem
        '
        Me.UsingDateToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ByDateToolStripMenuItem, Me.ByPartNoToolStripMenuItem})
        Me.UsingDateToolStripMenuItem.Name = "UsingDateToolStripMenuItem"
        Me.UsingDateToolStripMenuItem.Size = New System.Drawing.Size(196, 22)
        Me.UsingDateToolStripMenuItem.Text = "Active Wastage Note"
        '
        'ByDateToolStripMenuItem
        '
        Me.ByDateToolStripMenuItem.Name = "ByDateToolStripMenuItem"
        Me.ByDateToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.ByDateToolStripMenuItem.Text = "# by Date"
        '
        'ByPartNoToolStripMenuItem
        '
        Me.ByPartNoToolStripMenuItem.Name = "ByPartNoToolStripMenuItem"
        Me.ByPartNoToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.ByPartNoToolStripMenuItem.Text = "# by Part No"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ByDateToolStripMenuItem1, Me.BySupplierToolStripMenuItem1})
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(196, 22)
        Me.ToolStripMenuItem1.Text = "Deactive Wastage Note"
        '
        'ByDateToolStripMenuItem1
        '
        Me.ByDateToolStripMenuItem1.Name = "ByDateToolStripMenuItem1"
        Me.ByDateToolStripMenuItem1.Size = New System.Drawing.Size(152, 22)
        Me.ByDateToolStripMenuItem1.Text = "# by Date"
        '
        'BySupplierToolStripMenuItem1
        '
        Me.BySupplierToolStripMenuItem1.Name = "BySupplierToolStripMenuItem1"
        Me.BySupplierToolStripMenuItem1.Size = New System.Drawing.Size(152, 22)
        Me.BySupplierToolStripMenuItem1.Text = "# by Part No"
        '
        'RefreshToolStripMenuItem
        '
        Me.RefreshToolStripMenuItem.Image = CType(resources.GetObject("RefreshToolStripMenuItem.Image"), System.Drawing.Image)
        Me.RefreshToolStripMenuItem.Name = "RefreshToolStripMenuItem"
        Me.RefreshToolStripMenuItem.Size = New System.Drawing.Size(196, 22)
        Me.RefreshToolStripMenuItem.Text = "Refresh"
        '
        'ExitToolStripMenuItem1
        '
        Me.ExitToolStripMenuItem1.Image = CType(resources.GetObject("ExitToolStripMenuItem1.Image"), System.Drawing.Image)
        Me.ExitToolStripMenuItem1.Name = "ExitToolStripMenuItem1"
        Me.ExitToolStripMenuItem1.Size = New System.Drawing.Size(196, 22)
        Me.ExitToolStripMenuItem1.Text = "Exit"
        '
        'lblDisplay
        '
        Me.lblDisplay.AutoSize = True
        Me.lblDisplay.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDisplay.Location = New System.Drawing.Point(12, 39)
        Me.lblDisplay.Name = "lblDisplay"
        Me.lblDisplay.Size = New System.Drawing.Size(340, 23)
        Me.lblDisplay.TabIndex = 113
        Me.lblDisplay.Text = "Unique Motors (Pvt) Ltd - Magammana"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 62)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(132, 19)
        Me.Label2.TabIndex = 114
        Me.Label2.Text = "Wastage Entry Note"
        '
        'UltraGrid2
        '
        Me.UltraGrid2.Location = New System.Drawing.Point(16, 84)
        Me.UltraGrid2.Name = "UltraGrid2"
        Me.UltraGrid2.Size = New System.Drawing.Size(843, 437)
        Me.UltraGrid2.TabIndex = 121
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.UltraButton2)
        Me.Panel1.Controls.Add(Me.txtDate2)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.txtDate1)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Location = New System.Drawing.Point(223, 155)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(388, 58)
        Me.Panel1.TabIndex = 123
        Me.Panel1.Visible = False
        '
        'UltraButton2
        '
        Appearance2.Image = CType(resources.GetObject("Appearance2.Image"), Object)
        Appearance2.ImageHAlign = Infragistics.Win.HAlign.Center
        Me.UltraButton2.Appearance = Appearance2
        Me.UltraButton2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton2.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton2.Location = New System.Drawing.Point(338, 13)
        Me.UltraButton2.Name = "UltraButton2"
        Me.UltraButton2.Size = New System.Drawing.Size(33, 34)
        Me.UltraButton2.TabIndex = 278
        '
        'txtDate2
        '
        Me.txtDate2.BackColor = System.Drawing.SystemColors.Window
        Me.txtDate2.DateButtons.Add(DateButton13)
        Me.txtDate2.Location = New System.Drawing.Point(232, 18)
        Me.txtDate2.Name = "txtDate2"
        Me.txtDate2.NonAutoSizeHeight = 21
        Me.txtDate2.Size = New System.Drawing.Size(100, 21)
        Me.txtDate2.TabIndex = 275
        Me.txtDate2.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(186, 21)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(46, 13)
        Me.Label4.TabIndex = 274
        Me.Label4.Text = "To Date"
        '
        'txtDate1
        '
        Me.txtDate1.BackColor = System.Drawing.SystemColors.Window
        Me.txtDate1.DateButtons.Add(DateButton14)
        Me.txtDate1.Location = New System.Drawing.Point(80, 18)
        Me.txtDate1.Name = "txtDate1"
        Me.txtDate1.NonAutoSizeHeight = 21
        Me.txtDate1.Size = New System.Drawing.Size(100, 21)
        Me.txtDate1.TabIndex = 273
        Me.txtDate1.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 21)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(56, 13)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "From Date"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.UltraButton3)
        Me.Panel3.Controls.Add(Me.cboItem)
        Me.Panel3.Controls.Add(Me.Label6)
        Me.Panel3.Controls.Add(Me.txtC2)
        Me.Panel3.Controls.Add(Me.Label7)
        Me.Panel3.Controls.Add(Me.txtC1)
        Me.Panel3.Controls.Add(Me.Label9)
        Me.Panel3.Location = New System.Drawing.Point(183, 145)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(388, 79)
        Me.Panel3.TabIndex = 125
        Me.Panel3.Visible = False
        '
        'UltraButton3
        '
        Appearance1.Image = CType(resources.GetObject("Appearance1.Image"), Object)
        Appearance1.ImageHAlign = Infragistics.Win.HAlign.Center
        Me.UltraButton3.Appearance = Appearance1
        Me.UltraButton3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton3.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton3.Location = New System.Drawing.Point(348, 23)
        Me.UltraButton3.Name = "UltraButton3"
        Me.UltraButton3.Size = New System.Drawing.Size(33, 34)
        Me.UltraButton3.TabIndex = 278
        '
        'cboItem
        '
        Me.cboItem.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Suggest
        Me.cboItem.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Appearance356.BackColor = System.Drawing.SystemColors.Window
        Appearance356.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.cboItem.DisplayLayout.Appearance = Appearance356
        Me.cboItem.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.cboItem.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance357.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance357.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance357.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance357.BorderColor = System.Drawing.SystemColors.Window
        Me.cboItem.DisplayLayout.GroupByBox.Appearance = Appearance357
        Appearance358.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboItem.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance358
        Me.cboItem.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance359.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance359.BackColor2 = System.Drawing.SystemColors.Control
        Appearance359.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance359.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboItem.DisplayLayout.GroupByBox.PromptAppearance = Appearance359
        Me.cboItem.DisplayLayout.MaxColScrollRegions = 1
        Me.cboItem.DisplayLayout.MaxRowScrollRegions = 1
        Appearance360.BackColor = System.Drawing.SystemColors.Window
        Appearance360.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cboItem.DisplayLayout.Override.ActiveCellAppearance = Appearance360
        Appearance361.BackColor = System.Drawing.SystemColors.Highlight
        Appearance361.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.cboItem.DisplayLayout.Override.ActiveRowAppearance = Appearance361
        Me.cboItem.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.cboItem.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance362.BackColor = System.Drawing.SystemColors.Window
        Me.cboItem.DisplayLayout.Override.CardAreaAppearance = Appearance362
        Appearance363.BorderColor = System.Drawing.Color.Silver
        Appearance363.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.cboItem.DisplayLayout.Override.CellAppearance = Appearance363
        Me.cboItem.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.cboItem.DisplayLayout.Override.CellPadding = 0
        Appearance364.BackColor = System.Drawing.SystemColors.Control
        Appearance364.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance364.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance364.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance364.BorderColor = System.Drawing.SystemColors.Window
        Me.cboItem.DisplayLayout.Override.GroupByRowAppearance = Appearance364
        Appearance365.TextHAlignAsString = "Left"
        Me.cboItem.DisplayLayout.Override.HeaderAppearance = Appearance365
        Me.cboItem.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.cboItem.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance381.BackColor = System.Drawing.SystemColors.Window
        Appearance381.BorderColor = System.Drawing.Color.Silver
        Me.cboItem.DisplayLayout.Override.RowAppearance = Appearance381
        Me.cboItem.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance384.BackColor = System.Drawing.SystemColors.ControlLight
        Me.cboItem.DisplayLayout.Override.TemplateAddRowAppearance = Appearance384
        Me.cboItem.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.cboItem.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.cboItem.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.cboItem.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.[Default]
        Me.cboItem.Location = New System.Drawing.Point(79, 42)
        Me.cboItem.Name = "cboItem"
        Me.cboItem.Size = New System.Drawing.Size(252, 22)
        Me.cboItem.TabIndex = 277
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 42)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(43, 13)
        Me.Label6.TabIndex = 276
        Me.Label6.Text = "Part No"
        '
        'txtC2
        '
        Me.txtC2.BackColor = System.Drawing.SystemColors.Window
        Me.txtC2.DateButtons.Add(DateButton15)
        Me.txtC2.Location = New System.Drawing.Point(232, 18)
        Me.txtC2.Name = "txtC2"
        Me.txtC2.NonAutoSizeHeight = 21
        Me.txtC2.Size = New System.Drawing.Size(100, 21)
        Me.txtC2.TabIndex = 275
        Me.txtC2.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(186, 21)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(46, 13)
        Me.Label7.TabIndex = 274
        Me.Label7.Text = "To Date"
        '
        'txtC1
        '
        Me.txtC1.BackColor = System.Drawing.SystemColors.Window
        Me.txtC1.DateButtons.Add(DateButton16)
        Me.txtC1.Location = New System.Drawing.Point(80, 18)
        Me.txtC1.Name = "txtC1"
        Me.txtC1.NonAutoSizeHeight = 21
        Me.txtC1.Size = New System.Drawing.Size(100, 21)
        Me.txtC1.TabIndex = 273
        Me.txtC1.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(3, 21)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(56, 13)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "From Date"
        '
        'frmView_Wastage
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(865, 534)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.UltraGrid2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblDisplay)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Name = "frmView_Wastage"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "#View Wastage Entry Note"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.UltraGrid2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.txtDate2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDate1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.cboItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ReportFilterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UsingDateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ByDateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ByPartNoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ByDateToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BySupplierToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RefreshToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lblDisplay As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents UltraGrid2 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents UltraButton2 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents txtDate2 As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtDate1 As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents UltraButton3 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cboItem As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtC2 As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtC1 As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label9 As System.Windows.Forms.Label
End Class
