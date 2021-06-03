<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRow_Material
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
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance31 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance542 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRow_Material))
        Dim Appearance102 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance543 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance544 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cboCategory = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.txtName = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.txtCode = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel
        Me.cmdSave = New Infragistics.Win.Misc.UltraButton
        Me.UltraButton4 = New Infragistics.Win.Misc.UltraButton
        Me.UltraButton2 = New Infragistics.Win.Misc.UltraButton
        Me.UltraButton3 = New Infragistics.Win.Misc.UltraButton
        Me.OPR4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.cboCategory, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.OPR4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR4.SuspendLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.cboCategory)
        Me.OPR0.Controls.Add(Me.UltraLabel1)
        Me.OPR0.Controls.Add(Me.txtName)
        Me.OPR0.Controls.Add(Me.UltraLabel2)
        Me.OPR0.Controls.Add(Me.txtCode)
        Me.OPR0.Controls.Add(Me.UltraLabel4)
        Me.OPR0.Location = New System.Drawing.Point(25, 12)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(641, 134)
        Me.OPR0.TabIndex = 128
        Me.OPR0.Text = " "
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
        Me.cboCategory.Location = New System.Drawing.Point(85, 56)
        Me.cboCategory.Name = "cboCategory"
        Me.cboCategory.Size = New System.Drawing.Size(149, 22)
        Me.cboCategory.TabIndex = 0
        '
        'UltraLabel1
        '
        Appearance17.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel1.Appearance = Appearance17
        Me.UltraLabel1.Location = New System.Drawing.Point(6, 57)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel1.TabIndex = 117
        Me.UltraLabel1.Text = "Category"
        Me.UltraLabel1.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(85, 84)
        Me.txtName.MaxLength = 120
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(550, 21)
        Me.txtName.TabIndex = 1
        '
        'UltraLabel2
        '
        Appearance1.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel2.Appearance = Appearance1
        Me.UltraLabel2.Location = New System.Drawing.Point(6, 84)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel2.TabIndex = 115
        Me.UltraLabel2.Text = "Item Name"
        Me.UltraLabel2.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtCode
        '
        Me.txtCode.Location = New System.Drawing.Point(85, 30)
        Me.txtCode.MaxLength = 15
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(120, 21)
        Me.txtCode.TabIndex = 2
        '
        'UltraLabel4
        '
        Appearance31.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel4.Appearance = Appearance31
        Me.UltraLabel4.Location = New System.Drawing.Point(6, 30)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel4.TabIndex = 56
        Me.UltraLabel4.Text = "Item Code"
        Me.UltraLabel4.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'cmdSave
        '
        Appearance542.Image = CType(resources.GetObject("Appearance542.Image"), Object)
        Me.cmdSave.Appearance = Appearance542
        Me.cmdSave.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ImageSize = New System.Drawing.Size(22, 22)
        Me.cmdSave.Location = New System.Drawing.Point(26, 385)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(134, 44)
        Me.cmdSave.TabIndex = 263
        Me.cmdSave.Text = "&Save"
        '
        'UltraButton4
        '
        Appearance102.Image = CType(resources.GetObject("Appearance102.Image"), Object)
        Me.UltraButton4.Appearance = Appearance102
        Me.UltraButton4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton4.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton4.Location = New System.Drawing.Point(165, 385)
        Me.UltraButton4.Name = "UltraButton4"
        Me.UltraButton4.Size = New System.Drawing.Size(134, 44)
        Me.UltraButton4.TabIndex = 262
        Me.UltraButton4.Text = "&Delete"
        '
        'UltraButton2
        '
        Appearance543.Image = CType(resources.GetObject("Appearance543.Image"), Object)
        Me.UltraButton2.Appearance = Appearance543
        Me.UltraButton2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton2.ImageSize = New System.Drawing.Size(32, 32)
        Me.UltraButton2.Location = New System.Drawing.Point(305, 384)
        Me.UltraButton2.Name = "UltraButton2"
        Me.UltraButton2.Size = New System.Drawing.Size(134, 45)
        Me.UltraButton2.TabIndex = 261
        Me.UltraButton2.Text = "&Reset"
        '
        'UltraButton3
        '
        Appearance544.Image = CType(resources.GetObject("Appearance544.Image"), Object)
        Me.UltraButton3.Appearance = Appearance544
        Me.UltraButton3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton3.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton3.Location = New System.Drawing.Point(532, 384)
        Me.UltraButton3.Name = "UltraButton3"
        Me.UltraButton3.Size = New System.Drawing.Size(134, 45)
        Me.UltraButton3.TabIndex = 260
        Me.UltraButton3.Text = "&Exit"
        '
        'OPR4
        '
        Me.OPR4.Controls.Add(Me.UltraGrid1)
        Me.OPR4.Location = New System.Drawing.Point(25, 152)
        Me.OPR4.Name = "OPR4"
        Me.OPR4.Size = New System.Drawing.Size(641, 226)
        Me.OPR4.TabIndex = 259
        '
        'UltraGrid1
        '
        Me.UltraGrid1.Location = New System.Drawing.Point(6, 19)
        Me.UltraGrid1.Name = "UltraGrid1"
        Me.UltraGrid1.Size = New System.Drawing.Size(629, 201)
        Me.UltraGrid1.TabIndex = 1
        '
        'frmRow_Material
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(840, 480)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.UltraButton4)
        Me.Controls.Add(Me.UltraButton2)
        Me.Controls.Add(Me.UltraButton3)
        Me.Controls.Add(Me.OPR4)
        Me.Controls.Add(Me.OPR0)
        Me.Name = "frmRow_Material"
        Me.Text = "Row Material"
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.cboCategory, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.OPR4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR4.ResumeLayout(False)
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtName As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtCode As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents cmdSave As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraButton4 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraButton2 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraButton3 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents OPR4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents cboCategory As Infragistics.Win.UltraWinGrid.UltraCombo
End Class
