<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRollwt
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRollwt))
        Dim DateButton1 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim Appearance73 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim DateButton2 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim Appearance74 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraGroupBox5 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtTo = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.txtDate = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraGroupBox3 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdEdit = New Infragistics.Win.Misc.UltraButton
        Me.cmdAdd = New Infragistics.Win.Misc.UltraButton
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.UltraGroupBox4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdReset = New Infragistics.Win.Misc.UltraButton
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox5.SuspendLayout()
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.txtTo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox3.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'UltraGroupBox5
        '
        Me.UltraGroupBox5.Controls.Add(Me.cmdExit)
        Me.UltraGroupBox5.Location = New System.Drawing.Point(326, 98)
        Me.UltraGroupBox5.Name = "UltraGroupBox5"
        Me.UltraGroupBox5.Size = New System.Drawing.Size(133, 50)
        Me.UltraGroupBox5.TabIndex = 139
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
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.txtTo)
        Me.OPR0.Controls.Add(Me.UltraLabel2)
        Me.OPR0.Controls.Add(Me.txtDate)
        Me.OPR0.Controls.Add(Me.UltraLabel1)
        Me.OPR0.Enabled = False
        Me.OPR0.Location = New System.Drawing.Point(18, 24)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(441, 68)
        Me.OPR0.TabIndex = 137
        '
        'txtTo
        '
        Me.txtTo.BackColor = System.Drawing.SystemColors.Window
        Me.txtTo.DateButtons.Add(DateButton1)
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
        Me.txtDate.DateButtons.Add(DateButton2)
        Me.txtDate.Location = New System.Drawing.Point(90, 16)
        Me.txtDate.Name = "txtDate"
        Me.txtDate.NonAutoSizeHeight = 21
        Me.txtDate.Size = New System.Drawing.Size(125, 21)
        Me.txtDate.TabIndex = 78
        Me.txtDate.Value = New Date(2009, 3, 1, 0, 0, 0, 0)
        '
        'UltraLabel1
        '
        Appearance74.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel1.Appearance = Appearance74
        Me.UltraLabel1.Location = New System.Drawing.Point(6, 16)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel1.TabIndex = 34
        Me.UltraLabel1.Text = "From Date"
        Me.UltraLabel1.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'UltraGroupBox3
        '
        Me.UltraGroupBox3.Controls.Add(Me.cmdEdit)
        Me.UltraGroupBox3.Controls.Add(Me.cmdAdd)
        Me.UltraGroupBox3.Location = New System.Drawing.Point(18, 98)
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
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(614, 22)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(178, 90)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 141
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(-5, 154)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(284, 176)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 140
        Me.PictureBox1.TabStop = False
        '
        'UltraGroupBox4
        '
        Me.UltraGroupBox4.Controls.Add(Me.cmdReset)
        Me.UltraGroupBox4.Location = New System.Drawing.Point(221, 98)
        Me.UltraGroupBox4.Name = "UltraGroupBox4"
        Me.UltraGroupBox4.Size = New System.Drawing.Size(99, 50)
        Me.UltraGroupBox4.TabIndex = 142
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
        'frmRollwt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(802, 390)
        Me.Controls.Add(Me.UltraGroupBox4)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.UltraGroupBox5)
        Me.Controls.Add(Me.OPR0)
        Me.Controls.Add(Me.UltraGroupBox3)
        Me.Name = "frmRollwt"
        Me.Text = "Roll Weight Report"
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox5.ResumeLayout(False)
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.txtTo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox3.ResumeLayout(False)
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents UltraGroupBox5 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtTo As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtDate As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraGroupBox3 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdEdit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdAdd As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdReset As Infragistics.Win.Misc.UltraButton
End Class
