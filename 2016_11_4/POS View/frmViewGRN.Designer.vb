<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmViewGRN
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
        Dim DateButton1 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim DateButton2 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim DateButton3 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim DateButton4 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmViewGRN))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.FilterByToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ByDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.BySupplierToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.VATInvoiceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UltraGrid3 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.txtTo = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtFrom = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label1 = New System.Windows.Forms.Label
        Me.UltraButton3 = New Infragistics.Win.Misc.UltraButton
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.txtDate2 = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtDate1 = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label4 = New System.Windows.Forms.Label
        Me.UltraButton1 = New Infragistics.Win.Misc.UltraButton
        Me.cboSupp = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.Label5 = New System.Windows.Forms.Label
        Me.AZToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ZAToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ResetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MenuStrip1.SuspendLayout()
        CType(Me.UltraGrid3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.txtTo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtFrom, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        CType(Me.txtDate2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDate1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboSupp, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FilterByToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(904, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FilterByToolStripMenuItem
        '
        Me.FilterByToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ByDateToolStripMenuItem, Me.BySupplierToolStripMenuItem, Me.ResetToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.FilterByToolStripMenuItem.Name = "FilterByToolStripMenuItem"
        Me.FilterByToolStripMenuItem.Size = New System.Drawing.Size(61, 20)
        Me.FilterByToolStripMenuItem.Text = "Filter by"
        '
        'ByDateToolStripMenuItem
        '
        Me.ByDateToolStripMenuItem.Name = "ByDateToolStripMenuItem"
        Me.ByDateToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.ByDateToolStripMenuItem.Text = "by Date"
        '
        'BySupplierToolStripMenuItem
        '
        Me.BySupplierToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AZToolStripMenuItem, Me.ZAToolStripMenuItem, Me.VATInvoiceToolStripMenuItem})
        Me.BySupplierToolStripMenuItem.Name = "BySupplierToolStripMenuItem"
        Me.BySupplierToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.BySupplierToolStripMenuItem.Text = "by Supplier"
        '
        'VATInvoiceToolStripMenuItem
        '
        Me.VATInvoiceToolStripMenuItem.Name = "VATInvoiceToolStripMenuItem"
        Me.VATInvoiceToolStripMenuItem.Size = New System.Drawing.Size(137, 22)
        Me.VATInvoiceToolStripMenuItem.Text = "VAT Invoice"
        '
        'UltraGrid3
        '
        Me.UltraGrid3.Location = New System.Drawing.Point(11, 30)
        Me.UltraGrid3.Name = "UltraGrid3"
        Me.UltraGrid3.Size = New System.Drawing.Size(880, 364)
        Me.UltraGrid3.TabIndex = 107
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.txtTo)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.txtFrom)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.UltraButton3)
        Me.Panel1.Location = New System.Drawing.Point(252, 166)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(382, 61)
        Me.Panel1.TabIndex = 108
        Me.Panel1.Visible = False
        '
        'txtTo
        '
        Me.txtTo.BackColor = System.Drawing.SystemColors.Window
        Me.txtTo.DateButtons.Add(DateButton1)
        Me.txtTo.Location = New System.Drawing.Point(235, 19)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.NonAutoSizeHeight = 21
        Me.txtTo.Size = New System.Drawing.Size(100, 21)
        Me.txtTo.TabIndex = 276
        Me.txtTo.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(183, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 13)
        Me.Label2.TabIndex = 275
        Me.Label2.Text = "To Date"
        '
        'txtFrom
        '
        Me.txtFrom.BackColor = System.Drawing.SystemColors.Window
        Me.txtFrom.DateButtons.Add(DateButton2)
        Me.txtFrom.Location = New System.Drawing.Point(77, 19)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.NonAutoSizeHeight = 21
        Me.txtFrom.Size = New System.Drawing.Size(100, 21)
        Me.txtFrom.TabIndex = 274
        Me.txtFrom.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 13)
        Me.Label1.TabIndex = 262
        Me.Label1.Text = "From Date"
        '
        'UltraButton3
        '
        Appearance1.Image = CType(resources.GetObject("Appearance1.Image"), Object)
        Appearance1.ImageHAlign = Infragistics.Win.HAlign.Center
        Me.UltraButton3.Appearance = Appearance1
        Me.UltraButton3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton3.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton3.Location = New System.Drawing.Point(341, 11)
        Me.UltraButton3.Name = "UltraButton3"
        Me.UltraButton3.Size = New System.Drawing.Size(33, 34)
        Me.UltraButton3.TabIndex = 261
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.txtDate2)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.txtDate1)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.UltraButton1)
        Me.Panel2.Controls.Add(Me.cboSupp)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Location = New System.Drawing.Point(252, 167)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(401, 83)
        Me.Panel2.TabIndex = 109
        Me.Panel2.Visible = False
        '
        'txtDate2
        '
        Me.txtDate2.BackColor = System.Drawing.SystemColors.Window
        Me.txtDate2.DateButtons.Add(DateButton3)
        Me.txtDate2.Location = New System.Drawing.Point(245, 46)
        Me.txtDate2.Name = "txtDate2"
        Me.txtDate2.NonAutoSizeHeight = 21
        Me.txtDate2.Size = New System.Drawing.Size(100, 21)
        Me.txtDate2.TabIndex = 276
        Me.txtDate2.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(195, 49)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 13)
        Me.Label3.TabIndex = 275
        Me.Label3.Text = "To Date"
        '
        'txtDate1
        '
        Me.txtDate1.BackColor = System.Drawing.SystemColors.Window
        Me.txtDate1.DateButtons.Add(DateButton4)
        Me.txtDate1.Location = New System.Drawing.Point(89, 46)
        Me.txtDate1.Name = "txtDate1"
        Me.txtDate1.NonAutoSizeHeight = 21
        Me.txtDate1.Size = New System.Drawing.Size(100, 21)
        Me.txtDate1.TabIndex = 274
        Me.txtDate1.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 47)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 13)
        Me.Label4.TabIndex = 262
        Me.Label4.Text = "From Date"
        '
        'UltraButton1
        '
        Appearance544.Image = CType(resources.GetObject("Appearance544.Image"), Object)
        Appearance544.ImageHAlign = Infragistics.Win.HAlign.Center
        Me.UltraButton1.Appearance = Appearance544
        Me.UltraButton1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton1.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton1.Location = New System.Drawing.Point(351, 25)
        Me.UltraButton1.Name = "UltraButton1"
        Me.UltraButton1.Size = New System.Drawing.Size(33, 34)
        Me.UltraButton1.TabIndex = 261
        '
        'cboSupp
        '
        Me.cboSupp.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Suggest
        Me.cboSupp.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Appearance356.BackColor = System.Drawing.SystemColors.Window
        Appearance356.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.cboSupp.DisplayLayout.Appearance = Appearance356
        Me.cboSupp.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.cboSupp.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance357.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance357.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance357.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance357.BorderColor = System.Drawing.SystemColors.Window
        Me.cboSupp.DisplayLayout.GroupByBox.Appearance = Appearance357
        Appearance358.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboSupp.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance358
        Me.cboSupp.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance359.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance359.BackColor2 = System.Drawing.SystemColors.Control
        Appearance359.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance359.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboSupp.DisplayLayout.GroupByBox.PromptAppearance = Appearance359
        Me.cboSupp.DisplayLayout.MaxColScrollRegions = 1
        Me.cboSupp.DisplayLayout.MaxRowScrollRegions = 1
        Appearance360.BackColor = System.Drawing.SystemColors.Window
        Appearance360.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cboSupp.DisplayLayout.Override.ActiveCellAppearance = Appearance360
        Appearance361.BackColor = System.Drawing.SystemColors.Highlight
        Appearance361.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.cboSupp.DisplayLayout.Override.ActiveRowAppearance = Appearance361
        Me.cboSupp.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.cboSupp.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance362.BackColor = System.Drawing.SystemColors.Window
        Me.cboSupp.DisplayLayout.Override.CardAreaAppearance = Appearance362
        Appearance363.BorderColor = System.Drawing.Color.Silver
        Appearance363.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.cboSupp.DisplayLayout.Override.CellAppearance = Appearance363
        Me.cboSupp.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.cboSupp.DisplayLayout.Override.CellPadding = 0
        Appearance364.BackColor = System.Drawing.SystemColors.Control
        Appearance364.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance364.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance364.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance364.BorderColor = System.Drawing.SystemColors.Window
        Me.cboSupp.DisplayLayout.Override.GroupByRowAppearance = Appearance364
        Appearance365.TextHAlignAsString = "Left"
        Me.cboSupp.DisplayLayout.Override.HeaderAppearance = Appearance365
        Me.cboSupp.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.cboSupp.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance381.BackColor = System.Drawing.SystemColors.Window
        Appearance381.BorderColor = System.Drawing.Color.Silver
        Me.cboSupp.DisplayLayout.Override.RowAppearance = Appearance381
        Me.cboSupp.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance384.BackColor = System.Drawing.SystemColors.ControlLight
        Me.cboSupp.DisplayLayout.Override.TemplateAddRowAppearance = Appearance384
        Me.cboSupp.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.cboSupp.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.cboSupp.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.cboSupp.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.[Default]
        Me.cboSupp.Location = New System.Drawing.Point(89, 18)
        Me.cboSupp.Name = "cboSupp"
        Me.cboSupp.Size = New System.Drawing.Size(256, 22)
        Me.cboSupp.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 21)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(45, 13)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Supplier"
        '
        'AZToolStripMenuItem
        '
        Me.AZToolStripMenuItem.Image = CType(resources.GetObject("AZToolStripMenuItem.Image"), System.Drawing.Image)
        Me.AZToolStripMenuItem.Name = "AZToolStripMenuItem"
        Me.AZToolStripMenuItem.Size = New System.Drawing.Size(137, 22)
        Me.AZToolStripMenuItem.Text = "A-Z"
        '
        'ZAToolStripMenuItem
        '
        Me.ZAToolStripMenuItem.Image = CType(resources.GetObject("ZAToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ZAToolStripMenuItem.Name = "ZAToolStripMenuItem"
        Me.ZAToolStripMenuItem.Size = New System.Drawing.Size(137, 22)
        Me.ZAToolStripMenuItem.Text = "Z-A"
        '
        'ResetToolStripMenuItem
        '
        Me.ResetToolStripMenuItem.Image = CType(resources.GetObject("ResetToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ResetToolStripMenuItem.Name = "ResetToolStripMenuItem"
        Me.ResetToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.ResetToolStripMenuItem.Text = "Reset"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Image = CType(resources.GetObject("ExitToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'frmViewGRN
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(904, 415)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.UltraGrid3)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmViewGRN"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "## View GRN"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.UltraGrid3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.txtTo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtFrom, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.txtDate2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDate1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboSupp, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FilterByToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ByDateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BySupplierToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AZToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ZAToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents VATInvoiceToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ResetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UltraGrid3 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtTo As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtFrom As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents UltraButton3 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents txtDate2 As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDate1 As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents UltraButton1 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cboSupp As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents Label5 As System.Windows.Forms.Label
End Class
