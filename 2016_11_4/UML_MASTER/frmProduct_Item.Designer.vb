<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProduct_Item
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
        Dim Appearance48 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance49 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance75 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance103 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance104 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance105 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance106 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance107 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance108 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance109 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance110 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance111 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance28 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance102 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProduct_Item))
        Dim Appearance542 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance543 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance544 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cboName = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel
        Me.txtCode = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.OPR4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.cmdDelete = New Infragistics.Win.Misc.UltraButton
        Me.cmdSave = New Infragistics.Win.Misc.UltraButton
        Me.UltraButton2 = New Infragistics.Win.Misc.UltraButton
        Me.UltraButton3 = New Infragistics.Win.Misc.UltraButton
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.cboName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.OPR4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR4.SuspendLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.cboName)
        Me.OPR0.Controls.Add(Me.UltraLabel4)
        Me.OPR0.Controls.Add(Me.txtCode)
        Me.OPR0.Controls.Add(Me.UltraLabel2)
        Me.OPR0.Location = New System.Drawing.Point(21, 17)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(640, 98)
        Me.OPR0.TabIndex = 260
        '
        'cboName
        '
        Me.cboName.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Suggest
        Me.cboName.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Appearance48.BackColor = System.Drawing.SystemColors.Window
        Appearance48.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.cboName.DisplayLayout.Appearance = Appearance48
        Me.cboName.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.cboName.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance49.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance49.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance49.BorderColor = System.Drawing.SystemColors.Window
        Me.cboName.DisplayLayout.GroupByBox.Appearance = Appearance49
        Appearance75.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboName.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance75
        Me.cboName.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance103.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance103.BackColor2 = System.Drawing.SystemColors.Control
        Appearance103.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance103.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboName.DisplayLayout.GroupByBox.PromptAppearance = Appearance103
        Me.cboName.DisplayLayout.MaxColScrollRegions = 1
        Me.cboName.DisplayLayout.MaxRowScrollRegions = 1
        Appearance104.BackColor = System.Drawing.SystemColors.Window
        Appearance104.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cboName.DisplayLayout.Override.ActiveCellAppearance = Appearance104
        Appearance105.BackColor = System.Drawing.SystemColors.Highlight
        Appearance105.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.cboName.DisplayLayout.Override.ActiveRowAppearance = Appearance105
        Me.cboName.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.cboName.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance106.BackColor = System.Drawing.SystemColors.Window
        Me.cboName.DisplayLayout.Override.CardAreaAppearance = Appearance106
        Appearance107.BorderColor = System.Drawing.Color.Silver
        Appearance107.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.cboName.DisplayLayout.Override.CellAppearance = Appearance107
        Me.cboName.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.cboName.DisplayLayout.Override.CellPadding = 0
        Appearance108.BackColor = System.Drawing.SystemColors.Control
        Appearance108.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance108.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance108.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance108.BorderColor = System.Drawing.SystemColors.Window
        Me.cboName.DisplayLayout.Override.GroupByRowAppearance = Appearance108
        Appearance109.TextHAlignAsString = "Left"
        Me.cboName.DisplayLayout.Override.HeaderAppearance = Appearance109
        Me.cboName.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.cboName.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance110.BackColor = System.Drawing.SystemColors.Window
        Appearance110.BorderColor = System.Drawing.Color.Silver
        Me.cboName.DisplayLayout.Override.RowAppearance = Appearance110
        Me.cboName.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance111.BackColor = System.Drawing.SystemColors.ControlLight
        Me.cboName.DisplayLayout.Override.TemplateAddRowAppearance = Appearance111
        Me.cboName.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.cboName.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.cboName.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.cboName.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.[Default]
        Me.cboName.Location = New System.Drawing.Point(113, 41)
        Me.cboName.Name = "cboName"
        Me.cboName.Size = New System.Drawing.Size(511, 22)
        Me.cboName.TabIndex = 109
        '
        'UltraLabel4
        '
        Appearance28.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel4.Appearance = Appearance28
        Me.UltraLabel4.Location = New System.Drawing.Point(15, 43)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel4.TabIndex = 56
        Me.UltraLabel4.Text = "Box Name"
        Me.UltraLabel4.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtCode
        '
        Me.txtCode.Location = New System.Drawing.Point(113, 15)
        Me.txtCode.MaxLength = 15
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(100, 21)
        Me.txtCode.TabIndex = 51
        '
        'UltraLabel2
        '
        Appearance8.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel2.Appearance = Appearance8
        Me.UltraLabel2.Location = New System.Drawing.Point(15, 15)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(109, 21)
        Me.UltraLabel2.TabIndex = 36
        Me.UltraLabel2.Text = "Ref Code"
        Me.UltraLabel2.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'OPR4
        '
        Me.OPR4.Controls.Add(Me.UltraGrid1)
        Me.OPR4.Location = New System.Drawing.Point(21, 125)
        Me.OPR4.Name = "OPR4"
        Me.OPR4.Size = New System.Drawing.Size(641, 226)
        Me.OPR4.TabIndex = 261
        '
        'UltraGrid1
        '
        Me.UltraGrid1.Location = New System.Drawing.Point(6, 19)
        Me.UltraGrid1.Name = "UltraGrid1"
        Me.UltraGrid1.Size = New System.Drawing.Size(629, 201)
        Me.UltraGrid1.TabIndex = 1
        '
        'cmdDelete
        '
        Appearance102.Image = CType(resources.GetObject("Appearance102.Image"), Object)
        Me.cmdDelete.Appearance = Appearance102
        Me.cmdDelete.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDelete.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdDelete.Location = New System.Drawing.Point(161, 366)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(134, 43)
        Me.cmdDelete.TabIndex = 268
        Me.cmdDelete.Text = "&Delete"
        '
        'cmdSave
        '
        Appearance542.Image = CType(resources.GetObject("Appearance542.Image"), Object)
        Me.cmdSave.Appearance = Appearance542
        Me.cmdSave.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ImageSize = New System.Drawing.Size(22, 22)
        Me.cmdSave.Location = New System.Drawing.Point(21, 366)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(134, 43)
        Me.cmdSave.TabIndex = 267
        Me.cmdSave.Text = "&Save"
        '
        'UltraButton2
        '
        Appearance543.Image = CType(resources.GetObject("Appearance543.Image"), Object)
        Me.UltraButton2.Appearance = Appearance543
        Me.UltraButton2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton2.ImageSize = New System.Drawing.Size(32, 32)
        Me.UltraButton2.Location = New System.Drawing.Point(301, 365)
        Me.UltraButton2.Name = "UltraButton2"
        Me.UltraButton2.Size = New System.Drawing.Size(134, 44)
        Me.UltraButton2.TabIndex = 266
        Me.UltraButton2.Text = "&Reset"
        '
        'UltraButton3
        '
        Appearance544.Image = CType(resources.GetObject("Appearance544.Image"), Object)
        Me.UltraButton3.Appearance = Appearance544
        Me.UltraButton3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton3.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton3.Location = New System.Drawing.Point(528, 365)
        Me.UltraButton3.Name = "UltraButton3"
        Me.UltraButton3.Size = New System.Drawing.Size(134, 44)
        Me.UltraButton3.TabIndex = 265
        Me.UltraButton3.Text = "&Exit"
        '
        'frmProduct_Item
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(761, 476)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.UltraButton2)
        Me.Controls.Add(Me.UltraButton3)
        Me.Controls.Add(Me.OPR4)
        Me.Controls.Add(Me.OPR0)
        Me.Name = "frmProduct_Item"
        Me.Text = "Packing Box"
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.cboName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.OPR4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR4.ResumeLayout(False)
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cboName As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtCode As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents OPR4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents cmdDelete As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdSave As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraButton2 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraButton3 As Infragistics.Win.Misc.UltraButton
End Class
