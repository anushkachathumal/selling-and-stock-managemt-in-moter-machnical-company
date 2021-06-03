<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmQualityReport
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
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmQualityReport))
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance74 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim DateButton7 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim Appearance73 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim DateButton8 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance75 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance15 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraGroupBox5 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdReset = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox3 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdEdit = New Infragistics.Win.Misc.UltraButton
        Me.cmdAdd = New Infragistics.Win.Misc.UltraButton
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtToTime = New System.Windows.Forms.DateTimePicker
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel
        Me.txtTime1 = New System.Windows.Forms.DateTimePicker
        Me.UltraLabel3 = New Infragistics.Win.Misc.UltraLabel
        Me.txtTo = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.txtDate = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.chk1 = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.UltraLabel5 = New Infragistics.Win.Misc.UltraLabel
        Me.cboFrom = New Infragistics.Win.UltraWinGrid.UltraCombo
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox5.SuspendLayout()
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox4.SuspendLayout()
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox3.SuspendLayout()
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.txtTo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboFrom, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UltraGroupBox5
        '
        Me.UltraGroupBox5.Controls.Add(Me.cmdExit)
        Me.UltraGroupBox5.Location = New System.Drawing.Point(336, 152)
        Me.UltraGroupBox5.Name = "UltraGroupBox5"
        Me.UltraGroupBox5.Size = New System.Drawing.Size(133, 50)
        Me.UltraGroupBox5.TabIndex = 140
        '
        'cmdExit
        '
        Appearance4.Image = CType(resources.GetObject("Appearance4.Image"), Object)
        Me.cmdExit.Appearance = Appearance4
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
        Me.UltraGroupBox4.Location = New System.Drawing.Point(225, 153)
        Me.UltraGroupBox4.Name = "UltraGroupBox4"
        Me.UltraGroupBox4.Size = New System.Drawing.Size(99, 50)
        Me.UltraGroupBox4.TabIndex = 139
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
        Me.UltraGroupBox3.Location = New System.Drawing.Point(26, 152)
        Me.UltraGroupBox3.Name = "UltraGroupBox3"
        Me.UltraGroupBox3.Size = New System.Drawing.Size(193, 50)
        Me.UltraGroupBox3.TabIndex = 138
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
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.cboFrom)
        Me.OPR0.Controls.Add(Me.UltraLabel5)
        Me.OPR0.Controls.Add(Me.txtToTime)
        Me.OPR0.Controls.Add(Me.UltraLabel4)
        Me.OPR0.Controls.Add(Me.txtTime1)
        Me.OPR0.Controls.Add(Me.UltraLabel3)
        Me.OPR0.Controls.Add(Me.txtTo)
        Me.OPR0.Controls.Add(Me.UltraLabel2)
        Me.OPR0.Controls.Add(Me.txtDate)
        Me.OPR0.Controls.Add(Me.UltraLabel1)
        Me.OPR0.Enabled = False
        Me.OPR0.Location = New System.Drawing.Point(26, 31)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(441, 106)
        Me.OPR0.TabIndex = 137
        '
        'txtToTime
        '
        Me.txtToTime.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.txtToTime.Location = New System.Drawing.Point(306, 43)
        Me.txtToTime.Name = "txtToTime"
        Me.txtToTime.Size = New System.Drawing.Size(125, 20)
        Me.txtToTime.TabIndex = 88
        '
        'UltraLabel4
        '
        Appearance74.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel4.Appearance = Appearance74
        Me.UltraLabel4.Location = New System.Drawing.Point(230, 43)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel4.TabIndex = 87
        Me.UltraLabel4.Text = "To Time"
        Me.UltraLabel4.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtTime1
        '
        Me.txtTime1.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.txtTime1.Location = New System.Drawing.Point(90, 44)
        Me.txtTime1.Name = "txtTime1"
        Me.txtTime1.Size = New System.Drawing.Size(125, 20)
        Me.txtTime1.TabIndex = 86
        '
        'UltraLabel3
        '
        Appearance3.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel3.Appearance = Appearance3
        Me.UltraLabel3.Location = New System.Drawing.Point(6, 43)
        Me.UltraLabel3.Name = "UltraLabel3"
        Me.UltraLabel3.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel3.TabIndex = 85
        Me.UltraLabel3.Text = "From Time"
        Me.UltraLabel3.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtTo
        '
        Me.txtTo.BackColor = System.Drawing.SystemColors.Window
        Me.txtTo.DateButtons.Add(DateButton7)
        Me.txtTo.Location = New System.Drawing.Point(306, 16)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.NonAutoSizeHeight = 21
        Me.txtTo.Size = New System.Drawing.Size(125, 21)
        Me.txtTo.TabIndex = 80
        Me.txtTo.Value = New Date(2009, 3, 1, 0, 0, 0, 0)
        '
        'UltraLabel2
        '
        Appearance73.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel2.Appearance = Appearance73
        Me.UltraLabel2.Location = New System.Drawing.Point(230, 19)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel2.TabIndex = 79
        Me.UltraLabel2.Text = "To Date"
        Me.UltraLabel2.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtDate
        '
        Me.txtDate.BackColor = System.Drawing.SystemColors.Window
        Me.txtDate.DateButtons.Add(DateButton8)
        Me.txtDate.Location = New System.Drawing.Point(90, 16)
        Me.txtDate.Name = "txtDate"
        Me.txtDate.NonAutoSizeHeight = 21
        Me.txtDate.Size = New System.Drawing.Size(125, 21)
        Me.txtDate.TabIndex = 78
        Me.txtDate.Value = New Date(2009, 3, 1, 0, 0, 0, 0)
        '
        'UltraLabel1
        '
        Appearance1.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel1.Appearance = Appearance1
        Me.UltraLabel1.Location = New System.Drawing.Point(6, 16)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel1.TabIndex = 34
        Me.UltraLabel1.Text = "From Date"
        Me.UltraLabel1.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(3, 208)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(284, 176)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 141
        Me.PictureBox1.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(567, 23)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(178, 90)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 142
        Me.PictureBox2.TabStop = False
        '
        'chk1
        '
        Appearance75.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Appearance75.FontData.BoldAsString = "False"
        Me.chk1.Appearance = Appearance75
        Me.chk1.BackColor = System.Drawing.Color.Transparent
        Me.chk1.BackColorInternal = System.Drawing.Color.Transparent
        Me.chk1.Location = New System.Drawing.Point(26, 9)
        Me.chk1.Name = "chk1"
        Me.chk1.Size = New System.Drawing.Size(152, 16)
        Me.chk1.TabIndex = 145
        Me.chk1.Text = "Date and Time"
        '
        'UltraLabel5
        '
        Appearance2.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel5.Appearance = Appearance2
        Me.UltraLabel5.Location = New System.Drawing.Point(6, 70)
        Me.UltraLabel5.Name = "UltraLabel5"
        Me.UltraLabel5.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel5.TabIndex = 89
        Me.UltraLabel5.Text = "Fault Code"
        Me.UltraLabel5.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'cboFrom
        '
        Me.cboFrom.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.SuggestAppend
        Me.cboFrom.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Appearance5.BackColor = System.Drawing.SystemColors.Window
        Appearance5.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.cboFrom.DisplayLayout.Appearance = Appearance5
        Me.cboFrom.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.cboFrom.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance6.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance6.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance6.BorderColor = System.Drawing.SystemColors.Window
        Me.cboFrom.DisplayLayout.GroupByBox.Appearance = Appearance6
        Appearance7.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboFrom.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance7
        Me.cboFrom.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance8.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance8.BackColor2 = System.Drawing.SystemColors.Control
        Appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance8.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cboFrom.DisplayLayout.GroupByBox.PromptAppearance = Appearance8
        Me.cboFrom.DisplayLayout.MaxColScrollRegions = 1
        Me.cboFrom.DisplayLayout.MaxRowScrollRegions = 1
        Appearance9.BackColor = System.Drawing.SystemColors.Window
        Appearance9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cboFrom.DisplayLayout.Override.ActiveCellAppearance = Appearance9
        Appearance10.BackColor = System.Drawing.SystemColors.Highlight
        Appearance10.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.cboFrom.DisplayLayout.Override.ActiveRowAppearance = Appearance10
        Me.cboFrom.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.cboFrom.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Me.cboFrom.DisplayLayout.Override.CardAreaAppearance = Appearance11
        Appearance14.BorderColor = System.Drawing.Color.Silver
        Appearance14.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.cboFrom.DisplayLayout.Override.CellAppearance = Appearance14
        Me.cboFrom.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.cboFrom.DisplayLayout.Override.CellPadding = 0
        Appearance15.BackColor = System.Drawing.SystemColors.Control
        Appearance15.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance15.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance15.BorderColor = System.Drawing.SystemColors.Window
        Me.cboFrom.DisplayLayout.Override.GroupByRowAppearance = Appearance15
        Appearance19.TextHAlignAsString = "Left"
        Me.cboFrom.DisplayLayout.Override.HeaderAppearance = Appearance19
        Me.cboFrom.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.cboFrom.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance20.BackColor = System.Drawing.SystemColors.Window
        Appearance20.BorderColor = System.Drawing.Color.Silver
        Me.cboFrom.DisplayLayout.Override.RowAppearance = Appearance20
        Me.cboFrom.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance22.BackColor = System.Drawing.SystemColors.ControlLight
        Me.cboFrom.DisplayLayout.Override.TemplateAddRowAppearance = Appearance22
        Me.cboFrom.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.cboFrom.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.cboFrom.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.cboFrom.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.[Default]
        Me.cboFrom.Location = New System.Drawing.Point(90, 69)
        Me.cboFrom.Name = "cboFrom"
        Me.cboFrom.Size = New System.Drawing.Size(341, 22)
        Me.cboFrom.TabIndex = 91
        '
        'frmQualityReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(862, 484)
        Me.Controls.Add(Me.chk1)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.UltraGroupBox5)
        Me.Controls.Add(Me.UltraGroupBox4)
        Me.Controls.Add(Me.UltraGroupBox3)
        Me.Controls.Add(Me.OPR0)
        Me.Name = "frmQualityReport"
        Me.Text = "Weekly Quality wise Report Group by Fault"
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox5.ResumeLayout(False)
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox4.ResumeLayout(False)
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox3.ResumeLayout(False)
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.txtTo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboFrom, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents UltraGroupBox5 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdReset As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox3 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdEdit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdAdd As Infragistics.Win.Misc.UltraButton
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtToTime As System.Windows.Forms.DateTimePicker
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtTime1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents UltraLabel3 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtTo As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtDate As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents chk1 As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents UltraLabel5 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents cboFrom As Infragistics.Win.UltraWinGrid.UltraCombo
End Class
