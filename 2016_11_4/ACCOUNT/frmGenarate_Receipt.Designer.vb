<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGenarate_Receipt
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
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGenarate_Receipt))
        Dim DateButton1 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim DateButton2 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim Appearance542 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance543 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance544 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.UltraButton1 = New Infragistics.Win.Misc.UltraButton
        Me.txtDate4 = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtDate3 = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label6 = New System.Windows.Forms.Label
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.cmdSave = New Infragistics.Win.Misc.UltraButton
        Me.UltraButton2 = New Infragistics.Win.Misc.UltraButton
        Me.UltraButton3 = New Infragistics.Win.Misc.UltraButton
        Me.Panel2.SuspendLayout()
        CType(Me.txtDate4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDate3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.UltraButton1)
        Me.Panel2.Controls.Add(Me.txtDate4)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.txtDate3)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Location = New System.Drawing.Point(13, 12)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(361, 56)
        Me.Panel2.TabIndex = 16
        '
        'UltraButton1
        '
        Appearance1.Image = CType(resources.GetObject("Appearance1.Image"), Object)
        Appearance1.ImageHAlign = Infragistics.Win.HAlign.Center
        Me.UltraButton1.Appearance = Appearance1
        Me.UltraButton1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton1.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton1.Location = New System.Drawing.Point(324, 14)
        Me.UltraButton1.Name = "UltraButton1"
        Me.UltraButton1.Size = New System.Drawing.Size(33, 34)
        Me.UltraButton1.TabIndex = 278
        '
        'txtDate4
        '
        Me.txtDate4.BackColor = System.Drawing.SystemColors.Window
        Me.txtDate4.DateButtons.Add(DateButton1)
        Me.txtDate4.Location = New System.Drawing.Point(217, 18)
        Me.txtDate4.Name = "txtDate4"
        Me.txtDate4.NonAutoSizeHeight = 21
        Me.txtDate4.Size = New System.Drawing.Size(100, 21)
        Me.txtDate4.TabIndex = 275
        Me.txtDate4.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(171, 21)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(46, 13)
        Me.Label5.TabIndex = 274
        Me.Label5.Text = "To Date"
        '
        'txtDate3
        '
        Me.txtDate3.BackColor = System.Drawing.SystemColors.Window
        Me.txtDate3.DateButtons.Add(DateButton2)
        Me.txtDate3.Location = New System.Drawing.Point(65, 18)
        Me.txtDate3.Name = "txtDate3"
        Me.txtDate3.NonAutoSizeHeight = 21
        Me.txtDate3.Size = New System.Drawing.Size(100, 21)
        Me.txtDate3.TabIndex = 273
        Me.txtDate3.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 21)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(56, 13)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "From Date"
        '
        'UltraGrid1
        '
        Me.UltraGrid1.Location = New System.Drawing.Point(12, 74)
        Me.UltraGrid1.Name = "UltraGrid1"
        Me.UltraGrid1.Size = New System.Drawing.Size(834, 322)
        Me.UltraGrid1.TabIndex = 17
        '
        'cmdSave
        '
        Appearance542.Image = CType(resources.GetObject("Appearance542.Image"), Object)
        Me.cmdSave.Appearance = Appearance542
        Me.cmdSave.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ImageSize = New System.Drawing.Size(22, 22)
        Me.cmdSave.Location = New System.Drawing.Point(12, 403)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(134, 43)
        Me.cmdSave.TabIndex = 270
        Me.cmdSave.Text = "&Save"
        '
        'UltraButton2
        '
        Appearance543.Image = CType(resources.GetObject("Appearance543.Image"), Object)
        Me.UltraButton2.Appearance = Appearance543
        Me.UltraButton2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton2.ImageSize = New System.Drawing.Size(32, 32)
        Me.UltraButton2.Location = New System.Drawing.Point(151, 403)
        Me.UltraButton2.Name = "UltraButton2"
        Me.UltraButton2.Size = New System.Drawing.Size(134, 44)
        Me.UltraButton2.TabIndex = 269
        Me.UltraButton2.Text = "&Reset"
        '
        'UltraButton3
        '
        Appearance544.Image = CType(resources.GetObject("Appearance544.Image"), Object)
        Me.UltraButton3.Appearance = Appearance544
        Me.UltraButton3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton3.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton3.Location = New System.Drawing.Point(291, 403)
        Me.UltraButton3.Name = "UltraButton3"
        Me.UltraButton3.Size = New System.Drawing.Size(134, 44)
        Me.UltraButton3.TabIndex = 268
        Me.UltraButton3.Text = "&Exit"
        '
        'frmGenarate_Receipt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(874, 476)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.UltraButton2)
        Me.Controls.Add(Me.UltraButton3)
        Me.Controls.Add(Me.UltraGrid1)
        Me.Controls.Add(Me.Panel2)
        Me.Name = "frmGenarate_Receipt"
        Me.Text = "Genarate Receipt"
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.txtDate4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDate3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents UltraButton1 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents txtDate4 As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtDate3 As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents cmdSave As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraButton2 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraButton3 As Infragistics.Win.Misc.UltraButton
End Class
