<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOrder_Placement
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
        Dim Appearance38 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim DateButton2 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance42 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance40 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance36 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOrder_Placement))
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtTodate = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.txtFromDate = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.UltraLabel5 = New Infragistics.Win.Misc.UltraLabel
        Me.chkPTL = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.chkOCI = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.chkIN = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.UltraGroupBox5 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraButton6 = New Infragistics.Win.Misc.UltraButton
        Me.cmdReset = New Infragistics.Win.Misc.UltraButton
        Me.cmdSave = New Infragistics.Win.Misc.UltraButton
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.UltraLabel8 = New Infragistics.Win.Misc.UltraLabel
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.txtTodate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtFromDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox5.SuspendLayout()
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox4.SuspendLayout()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.txtTodate)
        Me.OPR0.Controls.Add(Me.UltraLabel1)
        Me.OPR0.Controls.Add(Me.txtFromDate)
        Me.OPR0.Controls.Add(Me.UltraLabel5)
        Me.OPR0.Location = New System.Drawing.Point(12, 34)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(537, 58)
        Me.OPR0.TabIndex = 114
        '
        'txtTodate
        '
        Me.txtTodate.BackColor = System.Drawing.SystemColors.Window
        Me.txtTodate.DateButtons.Add(DateButton1)
        Me.txtTodate.Location = New System.Drawing.Point(416, 16)
        Me.txtTodate.Name = "txtTodate"
        Me.txtTodate.NonAutoSizeHeight = 21
        Me.txtTodate.Size = New System.Drawing.Size(107, 21)
        Me.txtTodate.TabIndex = 82
        Me.txtTodate.Value = New Date(2008, 12, 1, 0, 0, 0, 0)
        '
        'UltraLabel1
        '
        Appearance38.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel1.Appearance = Appearance38
        Me.UltraLabel1.Location = New System.Drawing.Point(336, 16)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(103, 19)
        Me.UltraLabel1.TabIndex = 81
        Me.UltraLabel1.Text = "To Date"
        Me.UltraLabel1.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtFromDate
        '
        Me.txtFromDate.BackColor = System.Drawing.SystemColors.Window
        Me.txtFromDate.DateButtons.Add(DateButton2)
        Me.txtFromDate.Location = New System.Drawing.Point(101, 18)
        Me.txtFromDate.Name = "txtFromDate"
        Me.txtFromDate.NonAutoSizeHeight = 21
        Me.txtFromDate.Size = New System.Drawing.Size(107, 21)
        Me.txtFromDate.TabIndex = 80
        Me.txtFromDate.Value = New Date(2008, 12, 1, 0, 0, 0, 0)
        '
        'UltraLabel5
        '
        Appearance2.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel5.Appearance = Appearance2
        Me.UltraLabel5.Location = New System.Drawing.Point(6, 19)
        Me.UltraLabel5.Name = "UltraLabel5"
        Me.UltraLabel5.Size = New System.Drawing.Size(103, 19)
        Me.UltraLabel5.TabIndex = 79
        Me.UltraLabel5.Text = "From Date"
        Me.UltraLabel5.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'chkPTL
        '
        Appearance42.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Appearance42.FontData.BoldAsString = "True"
        Me.chkPTL.Appearance = Appearance42
        Me.chkPTL.BackColor = System.Drawing.Color.Transparent
        Me.chkPTL.BackColorInternal = System.Drawing.Color.Transparent
        Me.chkPTL.Location = New System.Drawing.Point(207, 8)
        Me.chkPTL.Name = "chkPTL"
        Me.chkPTL.Size = New System.Drawing.Size(121, 20)
        Me.chkPTL.TabIndex = 166
        Me.chkPTL.Text = "PTL"
        Me.chkPTL.Visible = False
        '
        'chkOCI
        '
        Appearance40.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Appearance40.FontData.BoldAsString = "True"
        Me.chkOCI.Appearance = Appearance40
        Me.chkOCI.BackColor = System.Drawing.Color.Transparent
        Me.chkOCI.BackColorInternal = System.Drawing.Color.Transparent
        Me.chkOCI.Location = New System.Drawing.Point(120, 8)
        Me.chkOCI.Name = "chkOCI"
        Me.chkOCI.Size = New System.Drawing.Size(121, 20)
        Me.chkOCI.TabIndex = 165
        Me.chkOCI.Text = "Out Source"
        '
        'chkIN
        '
        Appearance36.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Appearance36.FontData.BoldAsString = "True"
        Me.chkIN.Appearance = Appearance36
        Me.chkIN.BackColor = System.Drawing.Color.Transparent
        Me.chkIN.BackColorInternal = System.Drawing.Color.Transparent
        Me.chkIN.Location = New System.Drawing.Point(14, 8)
        Me.chkIN.Name = "chkIN"
        Me.chkIN.Size = New System.Drawing.Size(121, 20)
        Me.chkIN.TabIndex = 164
        Me.chkIN.Text = "IN House"
        '
        'UltraGroupBox5
        '
        Me.UltraGroupBox5.Controls.Add(Me.cmdExit)
        Me.UltraGroupBox5.Location = New System.Drawing.Point(366, 98)
        Me.UltraGroupBox5.Name = "UltraGroupBox5"
        Me.UltraGroupBox5.Size = New System.Drawing.Size(183, 69)
        Me.UltraGroupBox5.TabIndex = 180
        '
        'cmdExit
        '
        Appearance4.Image = CType(resources.GetObject("Appearance4.Image"), Object)
        Me.cmdExit.Appearance = Appearance4
        Me.cmdExit.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdExit.Location = New System.Drawing.Point(6, 13)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(163, 47)
        Me.cmdExit.TabIndex = 3
        Me.cmdExit.Text = "&Exit"
        '
        'UltraGroupBox4
        '
        Me.UltraGroupBox4.Controls.Add(Me.UltraButton6)
        Me.UltraGroupBox4.Controls.Add(Me.cmdReset)
        Me.UltraGroupBox4.Controls.Add(Me.cmdSave)
        Me.UltraGroupBox4.Location = New System.Drawing.Point(12, 98)
        Me.UltraGroupBox4.Name = "UltraGroupBox4"
        Me.UltraGroupBox4.Size = New System.Drawing.Size(353, 69)
        Me.UltraGroupBox4.TabIndex = 179
        '
        'UltraButton6
        '
        Appearance20.Image = CType(resources.GetObject("Appearance20.Image"), Object)
        Me.UltraButton6.Appearance = Appearance20
        Me.UltraButton6.Enabled = False
        Me.UltraButton6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton6.ImageSize = New System.Drawing.Size(32, 32)
        Me.UltraButton6.Location = New System.Drawing.Point(120, 10)
        Me.UltraButton6.Name = "UltraButton6"
        Me.UltraButton6.Size = New System.Drawing.Size(108, 49)
        Me.UltraButton6.TabIndex = 6
        Me.UltraButton6.Text = "&Save"
        '
        'cmdReset
        '
        Appearance16.Image = CType(resources.GetObject("Appearance16.Image"), Object)
        Me.cmdReset.Appearance = Appearance16
        Me.cmdReset.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReset.ImageSize = New System.Drawing.Size(32, 32)
        Me.cmdReset.Location = New System.Drawing.Point(234, 11)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.Size = New System.Drawing.Size(108, 49)
        Me.cmdReset.TabIndex = 5
        Me.cmdReset.Text = "&Reset"
        '
        'cmdSave
        '
        Appearance1.Image = CType(resources.GetObject("Appearance1.Image"), Object)
        Me.cmdSave.Appearance = Appearance1
        Me.cmdSave.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ImageSize = New System.Drawing.Size(32, 32)
        Me.cmdSave.Location = New System.Drawing.Point(6, 10)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(108, 49)
        Me.cmdSave.TabIndex = 4
        Me.cmdSave.Text = "&Preview"
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(263, 172)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(238, 111)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 178
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(515, 213)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(102, 34)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 177
        Me.PictureBox2.TabStop = False
        Me.PictureBox2.Visible = False
        '
        'UltraLabel8
        '
        Appearance5.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel8.Appearance = Appearance5
        Me.UltraLabel8.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel8.Location = New System.Drawing.Point(507, 253)
        Me.UltraLabel8.Name = "UltraLabel8"
        Me.UltraLabel8.Size = New System.Drawing.Size(361, 31)
        Me.UltraLabel8.TabIndex = 176
        Me.UltraLabel8.Text = "Order Placement Report"
        Me.UltraLabel8.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'frmOrder_Placement
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(790, 287)
        Me.Controls.Add(Me.UltraGroupBox5)
        Me.Controls.Add(Me.UltraGroupBox4)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.UltraLabel8)
        Me.Controls.Add(Me.chkPTL)
        Me.Controls.Add(Me.chkOCI)
        Me.Controls.Add(Me.chkIN)
        Me.Controls.Add(Me.OPR0)
        Me.Name = "frmOrder_Placement"
        Me.Text = "Order Placement"
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.txtTodate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtFromDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox5.ResumeLayout(False)
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox4.ResumeLayout(False)
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtTodate As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtFromDate As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents UltraLabel5 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents chkPTL As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents chkOCI As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents chkIN As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents UltraGroupBox5 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraButton6 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdReset As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdSave As Infragistics.Win.Misc.UltraButton
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents UltraLabel8 As Infragistics.Win.Misc.UltraLabel
End Class
