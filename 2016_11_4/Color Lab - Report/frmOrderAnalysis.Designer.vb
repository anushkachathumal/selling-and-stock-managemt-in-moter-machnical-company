<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOrderAnalysis
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOrderAnalysis))
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance35 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim DateButton1 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim Appearance57 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim DateButton2 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim Appearance41 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraGroupBox5 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdReset = New Infragistics.Win.Misc.UltraButton
        Me.cmdSave = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox3 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdEdit = New Infragistics.Win.Misc.UltraButton
        Me.cmdAdd = New Infragistics.Win.Misc.UltraButton
        Me.OPR1 = New Infragistics.Win.Misc.UltraGroupBox
        Me.lblPro = New Infragistics.Win.Misc.UltraLabel
        Me.pbCount = New Infragistics.Win.UltraWinProgressBar.UltraProgressBar
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtTodate = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.txtFromDate = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.UltraLabel5 = New Infragistics.Win.Misc.UltraLabel
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox5.SuspendLayout()
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox4.SuspendLayout()
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox3.SuspendLayout()
        CType(Me.OPR1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR1.SuspendLayout()
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.txtTodate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtFromDate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UltraGroupBox5
        '
        Me.UltraGroupBox5.Controls.Add(Me.cmdExit)
        Me.UltraGroupBox5.Location = New System.Drawing.Point(417, 158)
        Me.UltraGroupBox5.Name = "UltraGroupBox5"
        Me.UltraGroupBox5.Size = New System.Drawing.Size(142, 50)
        Me.UltraGroupBox5.TabIndex = 101
        '
        'cmdExit
        '
        Appearance21.Image = CType(resources.GetObject("Appearance21.Image"), Object)
        Me.cmdExit.Appearance = Appearance21
        Me.cmdExit.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdExit.Location = New System.Drawing.Point(6, 10)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(130, 30)
        Me.cmdExit.TabIndex = 3
        Me.cmdExit.Text = "&Exit"
        '
        'UltraGroupBox4
        '
        Me.UltraGroupBox4.Controls.Add(Me.cmdReset)
        Me.UltraGroupBox4.Controls.Add(Me.cmdSave)
        Me.UltraGroupBox4.Location = New System.Drawing.Point(219, 160)
        Me.UltraGroupBox4.Name = "UltraGroupBox4"
        Me.UltraGroupBox4.Size = New System.Drawing.Size(192, 50)
        Me.UltraGroupBox4.TabIndex = 100
        '
        'cmdReset
        '
        Appearance16.Image = CType(resources.GetObject("Appearance16.Image"), Object)
        Me.cmdReset.Appearance = Appearance16
        Me.cmdReset.Location = New System.Drawing.Point(97, 10)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.Size = New System.Drawing.Size(85, 30)
        Me.cmdReset.TabIndex = 5
        Me.cmdReset.Text = "&Reset"
        '
        'cmdSave
        '
        Appearance20.Image = Global.DBLotVbnet.My.Resources.Resources.save_as
        Me.cmdSave.Appearance = Appearance20
        Me.cmdSave.Enabled = False
        Me.cmdSave.Location = New System.Drawing.Point(6, 10)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(85, 30)
        Me.cmdSave.TabIndex = 4
        Me.cmdSave.Text = "&Print"
        '
        'UltraGroupBox3
        '
        Me.UltraGroupBox3.Controls.Add(Me.cmdEdit)
        Me.UltraGroupBox3.Controls.Add(Me.cmdAdd)
        Me.UltraGroupBox3.Location = New System.Drawing.Point(22, 160)
        Me.UltraGroupBox3.Name = "UltraGroupBox3"
        Me.UltraGroupBox3.Size = New System.Drawing.Size(192, 50)
        Me.UltraGroupBox3.TabIndex = 99
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
        Me.cmdEdit.Text = "&Edit"
        '
        'cmdAdd
        '
        Appearance18.Image = CType(resources.GetObject("Appearance18.Image"), Object)
        Appearance18.TextHAlignAsString = "Center"
        Me.cmdAdd.Appearance = Appearance18
        Me.cmdAdd.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdAdd.Location = New System.Drawing.Point(6, 11)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(85, 30)
        Me.cmdAdd.TabIndex = 0
        Me.cmdAdd.Text = "&Add"
        '
        'OPR1
        '
        Me.OPR1.Controls.Add(Me.lblPro)
        Me.OPR1.Controls.Add(Me.pbCount)
        Me.OPR1.Enabled = False
        Me.OPR1.Location = New System.Drawing.Point(22, 74)
        Me.OPR1.Name = "OPR1"
        Me.OPR1.Size = New System.Drawing.Size(537, 72)
        Me.OPR1.TabIndex = 98
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
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.txtTodate)
        Me.OPR0.Controls.Add(Me.UltraLabel1)
        Me.OPR0.Controls.Add(Me.txtFromDate)
        Me.OPR0.Controls.Add(Me.UltraLabel5)
        Me.OPR0.Enabled = False
        Me.OPR0.Location = New System.Drawing.Point(22, 16)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(537, 55)
        Me.OPR0.TabIndex = 97
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
        'frmOrderAnalysis
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(701, 262)
        Me.Controls.Add(Me.UltraGroupBox5)
        Me.Controls.Add(Me.UltraGroupBox4)
        Me.Controls.Add(Me.UltraGroupBox3)
        Me.Controls.Add(Me.OPR1)
        Me.Controls.Add(Me.OPR0)
        Me.Name = "frmOrderAnalysis"
        Me.Text = "Order Analysis Report"
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox5.ResumeLayout(False)
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox4.ResumeLayout(False)
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox3.ResumeLayout(False)
        CType(Me.OPR1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR1.ResumeLayout(False)
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.txtTodate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtFromDate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraGroupBox5 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdReset As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdSave As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox3 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdEdit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdAdd As Infragistics.Win.Misc.UltraButton
    Friend WithEvents OPR1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents lblPro As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents pbCount As Infragistics.Win.UltraWinProgressBar.UltraProgressBar
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtTodate As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtFromDate As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents UltraLabel5 As Infragistics.Win.Misc.UltraLabel
End Class
