<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmview_Outstanding
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmview_Outstanding))
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim DateButton1 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim DateButton2 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance13 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance15 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.ReportFilterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UsingDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ByDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.BySupplierToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
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
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.UltraButton1 = New Infragistics.Win.Misc.UltraButton
        Me.cboCustomer = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.Label3 = New System.Windows.Forms.Label
        Me.MenuStrip1.SuspendLayout()
        CType(Me.UltraGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.txtDate2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDate1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        CType(Me.cboCustomer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReportFilterToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(835, 24)
        Me.MenuStrip1.TabIndex = 112
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
        Me.UsingDateToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ByDateToolStripMenuItem, Me.BySupplierToolStripMenuItem})
        Me.UsingDateToolStripMenuItem.Name = "UsingDateToolStripMenuItem"
        Me.UsingDateToolStripMenuItem.Size = New System.Drawing.Size(229, 22)
        Me.UsingDateToolStripMenuItem.Text = "Outstanding Payment"
        '
        'ByDateToolStripMenuItem
        '
        Me.ByDateToolStripMenuItem.Name = "ByDateToolStripMenuItem"
        Me.ByDateToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.ByDateToolStripMenuItem.Text = "# by Date"
        '
        'BySupplierToolStripMenuItem
        '
        Me.BySupplierToolStripMenuItem.Name = "BySupplierToolStripMenuItem"
        Me.BySupplierToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.BySupplierToolStripMenuItem.Text = "# by Customer"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ByDateToolStripMenuItem1, Me.BySupplierToolStripMenuItem1})
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(229, 22)
        Me.ToolStripMenuItem1.Text = "Cancel Outstanding Payment"
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
        Me.BySupplierToolStripMenuItem1.Text = "# by Customer"
        '
        'RefreshToolStripMenuItem
        '
        Me.RefreshToolStripMenuItem.Image = CType(resources.GetObject("RefreshToolStripMenuItem.Image"), System.Drawing.Image)
        Me.RefreshToolStripMenuItem.Name = "RefreshToolStripMenuItem"
        Me.RefreshToolStripMenuItem.Size = New System.Drawing.Size(229, 22)
        Me.RefreshToolStripMenuItem.Text = "Refresh"
        '
        'ExitToolStripMenuItem1
        '
        Me.ExitToolStripMenuItem1.Image = CType(resources.GetObject("ExitToolStripMenuItem1.Image"), System.Drawing.Image)
        Me.ExitToolStripMenuItem1.Name = "ExitToolStripMenuItem1"
        Me.ExitToolStripMenuItem1.Size = New System.Drawing.Size(229, 22)
        Me.ExitToolStripMenuItem1.Text = "Exit"
        '
        'lblDisplay
        '
        Me.lblDisplay.AutoSize = True
        Me.lblDisplay.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDisplay.Location = New System.Drawing.Point(12, 39)
        Me.lblDisplay.Name = "lblDisplay"
        Me.lblDisplay.Size = New System.Drawing.Size(340, 23)
        Me.lblDisplay.TabIndex = 114
        Me.lblDisplay.Text = "Unique Motors (Pvt) Ltd - Magammana"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 62)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(146, 19)
        Me.Label2.TabIndex = 115
        Me.Label2.Text = "Outstanding Collection"
        '
        'UltraGrid2
        '
        Me.UltraGrid2.Location = New System.Drawing.Point(16, 94)
        Me.UltraGrid2.Name = "UltraGrid2"
        Me.UltraGrid2.Size = New System.Drawing.Size(807, 368)
        Me.UltraGrid2.TabIndex = 122
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
        Me.Panel1.Location = New System.Drawing.Point(248, 134)
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
        Me.txtDate2.DateButtons.Add(DateButton1)
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
        Me.txtDate1.DateButtons.Add(DateButton2)
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
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.UltraButton1)
        Me.Panel2.Controls.Add(Me.cboCustomer)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Location = New System.Drawing.Point(232, 134)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(508, 58)
        Me.Panel2.TabIndex = 124
        Me.Panel2.Visible = False
        '
        'UltraButton1
        '
        Appearance3.Image = CType(resources.GetObject("Appearance3.Image"), Object)
        Appearance3.ImageHAlign = Infragistics.Win.HAlign.Center
        Me.UltraButton1.Appearance = Appearance3
        Me.UltraButton1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton1.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton1.Location = New System.Drawing.Point(460, 13)
        Me.UltraButton1.Name = "UltraButton1"
        Me.UltraButton1.Size = New System.Drawing.Size(33, 34)
        Me.UltraButton1.TabIndex = 278
        '
        'cboCustomer
        '
        Me.cboCustomer.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Suggest
        Me.cboCustomer.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Appearance4.BackColor = System.Drawing.SystemColors.Window
        Appearance4.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.cboCustomer.DisplayLayout.Appearance = Appearance4
        Me.cboCustomer.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.cboCustomer.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance5.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance5.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance5.BorderColor = System.Drawing.SystemColors.Window
        Me.cboCustomer.DisplayLayout.GroupByBox.Appearance = Appearance5
        Appearance6.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboCustomer.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance6
        Me.cboCustomer.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance7.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance7.BackColor2 = System.Drawing.SystemColors.Control
        Appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance7.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboCustomer.DisplayLayout.GroupByBox.PromptAppearance = Appearance7
        Me.cboCustomer.DisplayLayout.MaxColScrollRegions = 1
        Me.cboCustomer.DisplayLayout.MaxRowScrollRegions = 1
        Appearance8.BackColor = System.Drawing.SystemColors.Window
        Appearance8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cboCustomer.DisplayLayout.Override.ActiveCellAppearance = Appearance8
        Appearance9.BackColor = System.Drawing.SystemColors.Highlight
        Appearance9.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.cboCustomer.DisplayLayout.Override.ActiveRowAppearance = Appearance9
        Me.cboCustomer.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.cboCustomer.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance10.BackColor = System.Drawing.SystemColors.Window
        Me.cboCustomer.DisplayLayout.Override.CardAreaAppearance = Appearance10
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Appearance11.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.cboCustomer.DisplayLayout.Override.CellAppearance = Appearance11
        Me.cboCustomer.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.cboCustomer.DisplayLayout.Override.CellPadding = 0
        Appearance12.BackColor = System.Drawing.SystemColors.Control
        Appearance12.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance12.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance12.BorderColor = System.Drawing.SystemColors.Window
        Me.cboCustomer.DisplayLayout.Override.GroupByRowAppearance = Appearance12
        Appearance13.TextHAlignAsString = "Left"
        Me.cboCustomer.DisplayLayout.Override.HeaderAppearance = Appearance13
        Me.cboCustomer.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.cboCustomer.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance14.BackColor = System.Drawing.SystemColors.Window
        Appearance14.BorderColor = System.Drawing.Color.Silver
        Me.cboCustomer.DisplayLayout.Override.RowAppearance = Appearance14
        Me.cboCustomer.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance15.BackColor = System.Drawing.SystemColors.ControlLight
        Me.cboCustomer.DisplayLayout.Override.TemplateAddRowAppearance = Appearance15
        Me.cboCustomer.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.cboCustomer.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.cboCustomer.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.cboCustomer.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.[Default]
        Me.cboCustomer.Location = New System.Drawing.Point(96, 18)
        Me.cboCustomer.Name = "cboCustomer"
        Me.cboCustomer.Size = New System.Drawing.Size(358, 22)
        Me.cboCustomer.TabIndex = 277
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 21)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(82, 13)
        Me.Label3.TabIndex = 276
        Me.Label3.Text = "Customer Name"
        '
        'frmview_Outstanding
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(835, 478)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.UltraGrid2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblDisplay)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MaximizeBox = False
        Me.Name = "frmview_Outstanding"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "#View Outstanding Collection"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.UltraGrid2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.txtDate2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDate1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.cboCustomer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ReportFilterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UsingDateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ByDateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BySupplierToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
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
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents UltraButton1 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cboCustomer As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
