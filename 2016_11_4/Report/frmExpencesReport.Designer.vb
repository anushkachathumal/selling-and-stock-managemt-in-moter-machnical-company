<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmExpencesReport
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmExpencesReport))
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim DateButton1 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim DateButton2 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraGroupBox5 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdReset = New Infragistics.Win.Misc.UltraButton
        Me.cmdSave = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox3 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdAdd = New Infragistics.Win.Misc.UltraButton
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtTo = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.txtDate = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
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
        Me.SuspendLayout()
        '
        'UltraGroupBox5
        '
        Me.UltraGroupBox5.Controls.Add(Me.cmdExit)
        Me.UltraGroupBox5.Location = New System.Drawing.Point(332, 103)
        Me.UltraGroupBox5.Name = "UltraGroupBox5"
        Me.UltraGroupBox5.Size = New System.Drawing.Size(130, 50)
        Me.UltraGroupBox5.TabIndex = 106
        '
        'cmdExit
        '
        Appearance21.Image = CType(resources.GetObject("Appearance21.Image"), Object)
        Me.cmdExit.Appearance = Appearance21
        Me.cmdExit.Location = New System.Drawing.Point(6, 10)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(118, 30)
        Me.cmdExit.TabIndex = 3
        Me.cmdExit.Text = "&Exit"
        '
        'UltraGroupBox4
        '
        Me.UltraGroupBox4.Controls.Add(Me.cmdReset)
        Me.UltraGroupBox4.Controls.Add(Me.cmdSave)
        Me.UltraGroupBox4.Location = New System.Drawing.Point(134, 104)
        Me.UltraGroupBox4.Name = "UltraGroupBox4"
        Me.UltraGroupBox4.Size = New System.Drawing.Size(192, 50)
        Me.UltraGroupBox4.TabIndex = 105
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
        Me.UltraGroupBox3.Controls.Add(Me.cmdAdd)
        Me.UltraGroupBox3.Location = New System.Drawing.Point(23, 103)
        Me.UltraGroupBox3.Name = "UltraGroupBox3"
        Me.UltraGroupBox3.Size = New System.Drawing.Size(104, 50)
        Me.UltraGroupBox3.TabIndex = 104
        '
        'cmdAdd
        '
        Appearance18.Image = CType(resources.GetObject("Appearance18.Image"), Object)
        Appearance18.TextHAlignAsString = "Center"
        Me.cmdAdd.Appearance = Appearance18
        Me.cmdAdd.Location = New System.Drawing.Point(10, 11)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(85, 30)
        Me.cmdAdd.TabIndex = 0
        Me.cmdAdd.Text = "&Add"
        '
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.txtTo)
        Me.OPR0.Controls.Add(Me.UltraLabel1)
        Me.OPR0.Controls.Add(Me.txtDate)
        Me.OPR0.Controls.Add(Me.UltraLabel2)
        Me.OPR0.Enabled = False
        Me.OPR0.Location = New System.Drawing.Point(21, 23)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(441, 61)
        Me.OPR0.TabIndex = 103
        '
        'txtTo
        '
        Me.txtTo.BackColor = System.Drawing.SystemColors.Window
        Me.txtTo.DateButtons.Add(DateButton1)
        Me.txtTo.Location = New System.Drawing.Point(310, 17)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.NonAutoSizeHeight = 21
        Me.txtTo.Size = New System.Drawing.Size(122, 21)
        Me.txtTo.TabIndex = 79
        Me.txtTo.Value = New Date(2009, 4, 1, 0, 0, 0, 0)
        '
        'UltraLabel1
        '
        Appearance5.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel1.Appearance = Appearance5
        Me.UltraLabel1.Location = New System.Drawing.Point(222, 17)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel1.TabIndex = 78
        Me.UltraLabel1.Text = "To Date"
        Me.UltraLabel1.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtDate
        '
        Me.txtDate.BackColor = System.Drawing.SystemColors.Window
        Me.txtDate.DateButtons.Add(DateButton2)
        Me.txtDate.Location = New System.Drawing.Point(86, 17)
        Me.txtDate.Name = "txtDate"
        Me.txtDate.NonAutoSizeHeight = 21
        Me.txtDate.Size = New System.Drawing.Size(122, 21)
        Me.txtDate.TabIndex = 77
        Me.txtDate.Value = New Date(2009, 4, 1, 0, 0, 0, 0)
        '
        'UltraLabel2
        '
        Appearance1.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel2.Appearance = Appearance1
        Me.UltraLabel2.Location = New System.Drawing.Point(6, 20)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel2.TabIndex = 36
        Me.UltraLabel2.Text = "From Date"
        Me.UltraLabel2.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'frmExpencesReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(535, 338)
        Me.Controls.Add(Me.UltraGroupBox5)
        Me.Controls.Add(Me.UltraGroupBox4)
        Me.Controls.Add(Me.UltraGroupBox3)
        Me.Controls.Add(Me.OPR0)
        Me.Name = "frmExpencesReport"
        Me.Text = "Collection Report"
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
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraGroupBox5 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdReset As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdSave As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox3 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdAdd As Infragistics.Win.Misc.UltraButton
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtTo As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtDate As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
End Class
