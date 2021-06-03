<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmExaminner_Eff
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
        Dim Appearance141 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance28 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance30 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance13 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance32 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance33 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance35 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance24 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance26 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance27 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance181 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmExaminner_Eff))
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.OPR2 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtYear = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel16 = New Infragistics.Win.Misc.UltraLabel
        Me.cboMonth = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.UltraLabel28 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraGroupBox5 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdReset = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox3 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdEdit = New Infragistics.Win.Misc.UltraButton
        Me.cmdAdd = New Infragistics.Win.Misc.UltraButton
        CType(Me.OPR2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR2.SuspendLayout()
        CType(Me.txtYear, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboMonth, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox5.SuspendLayout()
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox4.SuspendLayout()
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'OPR2
        '
        Me.OPR2.Controls.Add(Me.txtYear)
        Me.OPR2.Controls.Add(Me.UltraLabel16)
        Me.OPR2.Controls.Add(Me.cboMonth)
        Me.OPR2.Controls.Add(Me.UltraLabel28)
        Me.OPR2.Enabled = False
        Me.OPR2.Location = New System.Drawing.Point(22, 15)
        Me.OPR2.Name = "OPR2"
        Me.OPR2.Size = New System.Drawing.Size(442, 60)
        Me.OPR2.TabIndex = 142
        '
        'txtYear
        '
        Me.txtYear.Location = New System.Drawing.Point(332, 14)
        Me.txtYear.MaxLength = 150
        Me.txtYear.Name = "txtYear"
        Me.txtYear.Size = New System.Drawing.Size(93, 21)
        Me.txtYear.TabIndex = 77
        '
        'UltraLabel16
        '
        Appearance141.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel16.Appearance = Appearance141
        Me.UltraLabel16.Location = New System.Drawing.Point(266, 14)
        Me.UltraLabel16.Name = "UltraLabel16"
        Me.UltraLabel16.Size = New System.Drawing.Size(82, 21)
        Me.UltraLabel16.TabIndex = 62
        Me.UltraLabel16.Text = "Year"
        Me.UltraLabel16.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'cboMonth
        '
        Me.cboMonth.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Appearance28.BackColor = System.Drawing.SystemColors.Window
        Appearance28.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.cboMonth.DisplayLayout.Appearance = Appearance28
        Me.cboMonth.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.cboMonth.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance30.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance30.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance30.BorderColor = System.Drawing.SystemColors.Window
        Me.cboMonth.DisplayLayout.GroupByBox.Appearance = Appearance30
        Appearance12.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboMonth.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance12
        Me.cboMonth.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance13.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance13.BackColor2 = System.Drawing.SystemColors.Control
        Appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance13.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboMonth.DisplayLayout.GroupByBox.PromptAppearance = Appearance13
        Me.cboMonth.DisplayLayout.MaxColScrollRegions = 1
        Me.cboMonth.DisplayLayout.MaxRowScrollRegions = 1
        Appearance32.BackColor = System.Drawing.SystemColors.Window
        Appearance32.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cboMonth.DisplayLayout.Override.ActiveCellAppearance = Appearance32
        Appearance33.BackColor = System.Drawing.SystemColors.Highlight
        Appearance33.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.cboMonth.DisplayLayout.Override.ActiveRowAppearance = Appearance33
        Me.cboMonth.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.cboMonth.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance35.BackColor = System.Drawing.SystemColors.Window
        Me.cboMonth.DisplayLayout.Override.CardAreaAppearance = Appearance35
        Appearance23.BorderColor = System.Drawing.Color.Silver
        Appearance23.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.cboMonth.DisplayLayout.Override.CellAppearance = Appearance23
        Me.cboMonth.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.cboMonth.DisplayLayout.Override.CellPadding = 0
        Appearance24.BackColor = System.Drawing.SystemColors.Control
        Appearance24.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance24.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance24.BorderColor = System.Drawing.SystemColors.Window
        Me.cboMonth.DisplayLayout.Override.GroupByRowAppearance = Appearance24
        Appearance25.TextHAlignAsString = "Left"
        Me.cboMonth.DisplayLayout.Override.HeaderAppearance = Appearance25
        Me.cboMonth.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.cboMonth.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance26.BackColor = System.Drawing.SystemColors.Window
        Appearance26.BorderColor = System.Drawing.Color.Silver
        Me.cboMonth.DisplayLayout.Override.RowAppearance = Appearance26
        Me.cboMonth.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance27.BackColor = System.Drawing.SystemColors.ControlLight
        Me.cboMonth.DisplayLayout.Override.TemplateAddRowAppearance = Appearance27
        Me.cboMonth.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.cboMonth.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.cboMonth.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.cboMonth.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.[Default]
        Me.cboMonth.Location = New System.Drawing.Point(69, 14)
        Me.cboMonth.Name = "cboMonth"
        Me.cboMonth.Size = New System.Drawing.Size(140, 22)
        Me.cboMonth.TabIndex = 14
        '
        'UltraLabel28
        '
        Appearance181.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel28.Appearance = Appearance181
        Me.UltraLabel28.Location = New System.Drawing.Point(8, 16)
        Me.UltraLabel28.Name = "UltraLabel28"
        Me.UltraLabel28.Size = New System.Drawing.Size(73, 21)
        Me.UltraLabel28.TabIndex = 60
        Me.UltraLabel28.Text = "Month"
        Me.UltraLabel28.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'UltraGroupBox5
        '
        Me.UltraGroupBox5.Controls.Add(Me.cmdExit)
        Me.UltraGroupBox5.Location = New System.Drawing.Point(331, 82)
        Me.UltraGroupBox5.Name = "UltraGroupBox5"
        Me.UltraGroupBox5.Size = New System.Drawing.Size(133, 50)
        Me.UltraGroupBox5.TabIndex = 145
        '
        'cmdExit
        '
        Appearance21.Image = CType(resources.GetObject("Appearance21.Image"), Object)
        Me.cmdExit.Appearance = Appearance21
        Me.cmdExit.ImageSize = New System.Drawing.Size(18, 18)
        Me.cmdExit.Location = New System.Drawing.Point(6, 10)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(121, 30)
        Me.cmdExit.TabIndex = 3
        Me.cmdExit.Text = "&Exit"
        '
        'UltraGroupBox4
        '
        Me.UltraGroupBox4.Controls.Add(Me.cmdReset)
        Me.UltraGroupBox4.Location = New System.Drawing.Point(222, 82)
        Me.UltraGroupBox4.Name = "UltraGroupBox4"
        Me.UltraGroupBox4.Size = New System.Drawing.Size(99, 50)
        Me.UltraGroupBox4.TabIndex = 144
        '
        'cmdReset
        '
        Appearance16.Image = CType(resources.GetObject("Appearance16.Image"), Object)
        Me.cmdReset.Appearance = Appearance16
        Me.cmdReset.Location = New System.Drawing.Point(6, 9)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.Size = New System.Drawing.Size(85, 30)
        Me.cmdReset.TabIndex = 5
        Me.cmdReset.Text = "&Reset"
        '
        'UltraGroupBox3
        '
        Me.UltraGroupBox3.Controls.Add(Me.cmdEdit)
        Me.UltraGroupBox3.Controls.Add(Me.cmdAdd)
        Me.UltraGroupBox3.Location = New System.Drawing.Point(23, 82)
        Me.UltraGroupBox3.Name = "UltraGroupBox3"
        Me.UltraGroupBox3.Size = New System.Drawing.Size(193, 50)
        Me.UltraGroupBox3.TabIndex = 143
        '
        'cmdEdit
        '
        Appearance17.Image = CType(resources.GetObject("Appearance17.Image"), Object)
        Me.cmdEdit.Appearance = Appearance17
        Me.cmdEdit.Enabled = False
        Me.cmdEdit.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdEdit.Location = New System.Drawing.Point(100, 12)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(85, 30)
        Me.cmdEdit.TabIndex = 1
        Me.cmdEdit.Text = "&Print"
        '
        'cmdAdd
        '
        Appearance18.Image = CType(resources.GetObject("Appearance18.Image"), Object)
        Appearance18.TextHAlignAsString = "Center"
        Me.cmdAdd.Appearance = Appearance18
        Me.cmdAdd.ImageSize = New System.Drawing.Size(18, 18)
        Me.cmdAdd.Location = New System.Drawing.Point(6, 11)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(85, 30)
        Me.cmdAdd.TabIndex = 0
        Me.cmdAdd.Text = "&Add"
        '
        'frmExaminner_Eff
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(922, 262)
        Me.Controls.Add(Me.UltraGroupBox5)
        Me.Controls.Add(Me.UltraGroupBox4)
        Me.Controls.Add(Me.UltraGroupBox3)
        Me.Controls.Add(Me.OPR2)
        Me.Name = "frmExaminner_Eff"
        Me.Text = "Examinner wise  Efficiency "
        CType(Me.OPR2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR2.ResumeLayout(False)
        Me.OPR2.PerformLayout()
        CType(Me.txtYear, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboMonth, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox5.ResumeLayout(False)
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox4.ResumeLayout(False)
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents OPR2 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraLabel16 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents cboMonth As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents UltraLabel28 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtYear As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraGroupBox5 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdReset As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox3 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdEdit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdAdd As Infragistics.Win.Misc.UltraButton
End Class
