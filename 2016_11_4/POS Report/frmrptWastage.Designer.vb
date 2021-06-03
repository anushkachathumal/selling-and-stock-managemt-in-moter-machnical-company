<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmrptWastage
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmrptWastage))
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim DateButton1 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim DateButton2 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
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
        Dim DateButton3 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim DateButton4 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.ReportFilterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UsingDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UsingItemNameToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PrintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ResetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.UltraButton2 = New Infragistics.Win.Misc.UltraButton
        Me.txtDate4 = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtDate3 = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label6 = New System.Windows.Forms.Label
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.UltraButton4 = New Infragistics.Win.Misc.UltraButton
        Me.cboItem = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtC2 = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtC1 = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label13 = New System.Windows.Forms.Label
        Me.MenuStrip1.SuspendLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        CType(Me.txtDate4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDate3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel4.SuspendLayout()
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
        Me.MenuStrip1.Size = New System.Drawing.Size(1060, 24)
        Me.MenuStrip1.TabIndex = 23
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ReportFilterToolStripMenuItem
        '
        Me.ReportFilterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UsingDateToolStripMenuItem, Me.UsingItemNameToolStripMenuItem, Me.PrintToolStripMenuItem, Me.ResetToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.ReportFilterToolStripMenuItem.Name = "ReportFilterToolStripMenuItem"
        Me.ReportFilterToolStripMenuItem.Size = New System.Drawing.Size(83, 20)
        Me.ReportFilterToolStripMenuItem.Text = "Report Filter"
        '
        'UsingDateToolStripMenuItem
        '
        Me.UsingDateToolStripMenuItem.Name = "UsingDateToolStripMenuItem"
        Me.UsingDateToolStripMenuItem.Size = New System.Drawing.Size(166, 22)
        Me.UsingDateToolStripMenuItem.Text = "All Transaction"
        '
        'UsingItemNameToolStripMenuItem
        '
        Me.UsingItemNameToolStripMenuItem.Name = "UsingItemNameToolStripMenuItem"
        Me.UsingItemNameToolStripMenuItem.Size = New System.Drawing.Size(166, 22)
        Me.UsingItemNameToolStripMenuItem.Text = "Using Item Name"
        '
        'PrintToolStripMenuItem
        '
        Me.PrintToolStripMenuItem.Image = CType(resources.GetObject("PrintToolStripMenuItem.Image"), System.Drawing.Image)
        Me.PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        Me.PrintToolStripMenuItem.Size = New System.Drawing.Size(166, 22)
        Me.PrintToolStripMenuItem.Text = "Print"
        '
        'ResetToolStripMenuItem
        '
        Me.ResetToolStripMenuItem.Image = CType(resources.GetObject("ResetToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ResetToolStripMenuItem.Name = "ResetToolStripMenuItem"
        Me.ResetToolStripMenuItem.Size = New System.Drawing.Size(166, 22)
        Me.ResetToolStripMenuItem.Text = "Reset"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Image = CType(resources.GetObject("ExitToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(166, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'UltraGrid1
        '
        Me.UltraGrid1.Location = New System.Drawing.Point(12, 97)
        Me.UltraGrid1.Name = "UltraGrid1"
        Me.UltraGrid1.Size = New System.Drawing.Size(902, 386)
        Me.UltraGrid1.TabIndex = 24
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(223, 23)
        Me.Label1.TabIndex = 25
        Me.Label1.Text = "SS Distributor-Moratuwa"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 63)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(111, 19)
        Me.Label2.TabIndex = 26
        Me.Label2.Text = "Wastage  Report"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.UltraButton2)
        Me.Panel3.Controls.Add(Me.txtDate4)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Controls.Add(Me.txtDate3)
        Me.Panel3.Controls.Add(Me.Label6)
        Me.Panel3.Location = New System.Drawing.Point(308, 144)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(365, 59)
        Me.Panel3.TabIndex = 27
        Me.Panel3.Visible = False
        '
        'UltraButton2
        '
        Appearance2.Image = CType(resources.GetObject("Appearance2.Image"), Object)
        Appearance2.ImageHAlign = Infragistics.Win.HAlign.Center
        Me.UltraButton2.Appearance = Appearance2
        Me.UltraButton2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton2.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton2.Location = New System.Drawing.Point(324, 11)
        Me.UltraButton2.Name = "UltraButton2"
        Me.UltraButton2.Size = New System.Drawing.Size(33, 34)
        Me.UltraButton2.TabIndex = 278
        '
        'txtDate4
        '
        Me.txtDate4.BackColor = System.Drawing.SystemColors.Window
        Me.txtDate4.DateButtons.Add(DateButton1)
        Me.txtDate4.Location = New System.Drawing.Point(217, 18)
        Me.txtDate4.Name = "txtDate4"
        Me.txtDate4.NonAutoSizeHeight = 21
        Me.txtDate4.Size = New System.Drawing.Size(100, 21)
        Me.txtDate4.TabIndex = 275
        Me.txtDate4.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(171, 21)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(46, 13)
        Me.Label5.TabIndex = 274
        Me.Label5.Text = "To Date"
        '
        'txtDate3
        '
        Me.txtDate3.BackColor = System.Drawing.SystemColors.Window
        Me.txtDate3.DateButtons.Add(DateButton2)
        Me.txtDate3.Location = New System.Drawing.Point(65, 18)
        Me.txtDate3.Name = "txtDate3"
        Me.txtDate3.NonAutoSizeHeight = 21
        Me.txtDate3.Size = New System.Drawing.Size(100, 21)
        Me.txtDate3.TabIndex = 273
        Me.txtDate3.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 21)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(56, 13)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "From Date"
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Controls.Add(Me.UltraButton4)
        Me.Panel4.Controls.Add(Me.cboItem)
        Me.Panel4.Controls.Add(Me.Label11)
        Me.Panel4.Controls.Add(Me.txtC2)
        Me.Panel4.Controls.Add(Me.Label12)
        Me.Panel4.Controls.Add(Me.txtC1)
        Me.Panel4.Controls.Add(Me.Label13)
        Me.Panel4.Location = New System.Drawing.Point(308, 144)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(383, 79)
        Me.Panel4.TabIndex = 28
        Me.Panel4.Visible = False
        '
        'UltraButton4
        '
        Appearance1.Image = CType(resources.GetObject("Appearance1.Image"), Object)
        Appearance1.ImageHAlign = Infragistics.Win.HAlign.Center
        Me.UltraButton4.Appearance = Appearance1
        Me.UltraButton4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton4.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton4.Location = New System.Drawing.Point(343, 21)
        Me.UltraButton4.Name = "UltraButton4"
        Me.UltraButton4.Size = New System.Drawing.Size(33, 34)
        Me.UltraButton4.TabIndex = 278
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
        Me.cboItem.Location = New System.Drawing.Point(85, 42)
        Me.cboItem.Name = "cboItem"
        Me.cboItem.Size = New System.Drawing.Size(258, 22)
        Me.cboItem.TabIndex = 277
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 42)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(55, 13)
        Me.Label11.TabIndex = 276
        Me.Label11.Text = "Item Code"
        '
        'txtC2
        '
        Me.txtC2.BackColor = System.Drawing.SystemColors.Window
        Me.txtC2.DateButtons.Add(DateButton3)
        Me.txtC2.Location = New System.Drawing.Point(243, 18)
        Me.txtC2.Name = "txtC2"
        Me.txtC2.NonAutoSizeHeight = 21
        Me.txtC2.Size = New System.Drawing.Size(100, 21)
        Me.txtC2.TabIndex = 275
        Me.txtC2.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(191, 21)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(46, 13)
        Me.Label12.TabIndex = 274
        Me.Label12.Text = "To Date"
        '
        'txtC1
        '
        Me.txtC1.BackColor = System.Drawing.SystemColors.Window
        Me.txtC1.DateButtons.Add(DateButton4)
        Me.txtC1.Location = New System.Drawing.Point(85, 18)
        Me.txtC1.Name = "txtC1"
        Me.txtC1.NonAutoSizeHeight = 21
        Me.txtC1.Size = New System.Drawing.Size(100, 21)
        Me.txtC1.TabIndex = 273
        Me.txtC1.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(3, 21)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(56, 13)
        Me.Label13.TabIndex = 0
        Me.Label13.Text = "From Date"
        '
        'frmrptWastage
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1060, 474)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.UltraGrid1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Name = "frmrptWastage"
        Me.Text = "##View Report"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.txtDate4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDate3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        CType(Me.cboItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ReportFilterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UsingDateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UsingItemNameToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrintToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ResetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents UltraButton2 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents txtDate4 As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtDate3 As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents UltraButton4 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cboItem As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtC2 As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtC1 As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label13 As System.Windows.Forms.Label
End Class
