<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOrder_Status_Report
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
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOrder_Status_Report))
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim DateButton1 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim Appearance57 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim DateButton2 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim Appearance41 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance35 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraGroupBox5 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdReset = New Infragistics.Win.Misc.UltraButton
        Me.cmdSave = New Infragistics.Win.Misc.UltraButton
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtTodate = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.txtFromDate = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.UltraLabel5 = New Infragistics.Win.Misc.UltraLabel
        Me.OPR1 = New Infragistics.Win.Misc.UltraGroupBox
        Me.lblPro = New Infragistics.Win.Misc.UltraLabel
        Me.pbCount = New Infragistics.Win.UltraWinProgressBar.UltraProgressBar
        Me.UltraButton1 = New Infragistics.Win.Misc.UltraButton
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox5.SuspendLayout()
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox4.SuspendLayout()
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.txtTodate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtFromDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.OPR1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR1.SuspendLayout()
        Me.SuspendLayout()
        '
        'UltraGroupBox5
        '
        Me.UltraGroupBox5.Controls.Add(Me.cmdExit)
        Me.UltraGroupBox5.Location = New System.Drawing.Point(421, 156)
        Me.UltraGroupBox5.Name = "UltraGroupBox5"
        Me.UltraGroupBox5.Size = New System.Drawing.Size(142, 69)
        Me.UltraGroupBox5.TabIndex = 111
        '
        'cmdExit
        '
        Appearance21.Image = CType(resources.GetObject("Appearance21.Image"), Object)
        Me.cmdExit.Appearance = Appearance21
        Me.cmdExit.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdExit.Location = New System.Drawing.Point(6, 10)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(130, 47)
        Me.cmdExit.TabIndex = 3
        Me.cmdExit.Text = "&Exit"
        '
        'UltraGroupBox4
        '
        Me.UltraGroupBox4.Controls.Add(Me.UltraButton1)
        Me.UltraGroupBox4.Controls.Add(Me.cmdReset)
        Me.UltraGroupBox4.Controls.Add(Me.cmdSave)
        Me.UltraGroupBox4.Location = New System.Drawing.Point(26, 156)
        Me.UltraGroupBox4.Name = "UltraGroupBox4"
        Me.UltraGroupBox4.Size = New System.Drawing.Size(353, 69)
        Me.UltraGroupBox4.TabIndex = 110
        '
        'cmdReset
        '
        Appearance16.Image = CType(resources.GetObject("Appearance16.Image"), Object)
        Me.cmdReset.Appearance = Appearance16
        Me.cmdReset.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReset.ImageSize = New System.Drawing.Size(32, 32)
        Me.cmdReset.Location = New System.Drawing.Point(234, 8)
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
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.txtTodate)
        Me.OPR0.Controls.Add(Me.UltraLabel1)
        Me.OPR0.Controls.Add(Me.txtFromDate)
        Me.OPR0.Controls.Add(Me.UltraLabel5)
        Me.OPR0.Location = New System.Drawing.Point(25, 12)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(537, 55)
        Me.OPR0.TabIndex = 108
        '
        'txtTodate
        '
        Me.txtTodate.BackColor = System.Drawing.SystemColors.Window
        Me.txtTodate.DateButtons.Add(DateButton1)
        Me.txtTodate.Location = New System.Drawing.Point(424, 16)
        Me.txtTodate.Name = "txtTodate"
        Me.txtTodate.NonAutoSizeHeight = 21
        Me.txtTodate.Size = New System.Drawing.Size(107, 21)
        Me.txtTodate.TabIndex = 82
        Me.txtTodate.Value = New Date(2008, 12, 1, 0, 0, 0, 0)
        '
        'UltraLabel1
        '
        Appearance57.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel1.Appearance = Appearance57
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
        Appearance41.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel5.Appearance = Appearance41
        Me.UltraLabel5.Location = New System.Drawing.Point(6, 19)
        Me.UltraLabel5.Name = "UltraLabel5"
        Me.UltraLabel5.Size = New System.Drawing.Size(103, 19)
        Me.UltraLabel5.TabIndex = 79
        Me.UltraLabel5.Text = "From Date"
        Me.UltraLabel5.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'OPR1
        '
        Me.OPR1.Controls.Add(Me.lblPro)
        Me.OPR1.Controls.Add(Me.pbCount)
        Me.OPR1.Enabled = False
        Me.OPR1.Location = New System.Drawing.Point(26, 73)
        Me.OPR1.Name = "OPR1"
        Me.OPR1.Size = New System.Drawing.Size(537, 72)
        Me.OPR1.TabIndex = 112
        '
        'lblPro
        '
        Appearance35.BackColor = System.Drawing.Color.White
        Appearance35.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.lblPro.Appearance = Appearance35
        Me.lblPro.Location = New System.Drawing.Point(6, 47)
        Me.lblPro.Name = "lblPro"
        Me.lblPro.Size = New System.Drawing.Size(396, 19)
        Me.lblPro.TabIndex = 85
        Me.lblPro.Text = "Progress ...."
        Me.lblPro.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'pbCount
        '
        Appearance22.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Appearance22.FontData.BoldAsString = "True"
        Appearance22.ForeColorDisabled = System.Drawing.Color.Black
        Me.pbCount.Appearance = Appearance22
        Appearance23.FontData.BoldAsString = "True"
        Me.pbCount.FillAppearance = Appearance23
        Me.pbCount.Location = New System.Drawing.Point(6, 19)
        Me.pbCount.Name = "pbCount"
        Me.pbCount.Size = New System.Drawing.Size(525, 21)
        Me.pbCount.TabIndex = 48
        Me.pbCount.Text = "[Formatted]"
        '
        'UltraButton1
        '
        Appearance20.Image = CType(resources.GetObject("Appearance20.Image"), Object)
        Me.UltraButton1.Appearance = Appearance20
        Me.UltraButton1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton1.ImageSize = New System.Drawing.Size(32, 32)
        Me.UltraButton1.Location = New System.Drawing.Point(120, 8)
        Me.UltraButton1.Name = "UltraButton1"
        Me.UltraButton1.Size = New System.Drawing.Size(108, 49)
        Me.UltraButton1.TabIndex = 6
        Me.UltraButton1.Text = "&Print"
        '
        'frmOrder_Status_Report
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(596, 288)
        Me.Controls.Add(Me.OPR1)
        Me.Controls.Add(Me.UltraGroupBox5)
        Me.Controls.Add(Me.UltraGroupBox4)
        Me.Controls.Add(Me.OPR0)
        Me.Name = "frmOrder_Status_Report"
        Me.Text = "Order Status Report"
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox5.ResumeLayout(False)
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox4.ResumeLayout(False)
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.txtTodate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtFromDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.OPR1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraGroupBox5 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdReset As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdSave As Infragistics.Win.Misc.UltraButton
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtTodate As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtFromDate As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents UltraLabel5 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents OPR1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents lblPro As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents pbCount As Infragistics.Win.UltraWinProgressBar.UltraProgressBar
    Friend WithEvents UltraButton1 As Infragistics.Win.Misc.UltraButton
End Class
