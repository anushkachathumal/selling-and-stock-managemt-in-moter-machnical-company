<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmVarions_rpt
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVarions_rpt))
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim DateButton1 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim DateButton2 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.ReportFilterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UsingDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PrintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ResetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.UltraButton4 = New Infragistics.Win.Misc.UltraButton
        Me.txtC2 = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtC1 = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label13 = New System.Windows.Forms.Label
        Me.MenuStrip1.SuspendLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel4.SuspendLayout()
        CType(Me.txtC2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReportFilterToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(937, 24)
        Me.MenuStrip1.TabIndex = 23
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ReportFilterToolStripMenuItem
        '
        Me.ReportFilterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UsingDateToolStripMenuItem, Me.PrintToolStripMenuItem, Me.ResetToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.ReportFilterToolStripMenuItem.Name = "ReportFilterToolStripMenuItem"
        Me.ReportFilterToolStripMenuItem.Size = New System.Drawing.Size(83, 20)
        Me.ReportFilterToolStripMenuItem.Text = "Report Filter"
        '
        'UsingDateToolStripMenuItem
        '
        Me.UsingDateToolStripMenuItem.Name = "UsingDateToolStripMenuItem"
        Me.UsingDateToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.UsingDateToolStripMenuItem.Text = "All Transaction"
        '
        'PrintToolStripMenuItem
        '
        Me.PrintToolStripMenuItem.Image = CType(resources.GetObject("PrintToolStripMenuItem.Image"), System.Drawing.Image)
        Me.PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        Me.PrintToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.PrintToolStripMenuItem.Text = "Print"
        '
        'ResetToolStripMenuItem
        '
        Me.ResetToolStripMenuItem.Image = CType(resources.GetObject("ResetToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ResetToolStripMenuItem.Name = "ResetToolStripMenuItem"
        Me.ResetToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.ResetToolStripMenuItem.Text = "Reset"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Image = CType(resources.GetObject("ExitToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'UltraGrid1
        '
        Me.UltraGrid1.Location = New System.Drawing.Point(12, 94)
        Me.UltraGrid1.Name = "UltraGrid1"
        Me.UltraGrid1.Size = New System.Drawing.Size(992, 386)
        Me.UltraGrid1.TabIndex = 24
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 39)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(223, 23)
        Me.Label1.TabIndex = 25
        Me.Label1.Text = "SS Distributor-Moratuwa"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 65)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 19)
        Me.Label2.TabIndex = 26
        Me.Label2.Text = "Varions Report"
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Controls.Add(Me.UltraButton4)
        Me.Panel4.Controls.Add(Me.txtC2)
        Me.Panel4.Controls.Add(Me.Label12)
        Me.Panel4.Controls.Add(Me.txtC1)
        Me.Panel4.Controls.Add(Me.Label13)
        Me.Panel4.Location = New System.Drawing.Point(241, 130)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(383, 59)
        Me.Panel4.TabIndex = 29
        Me.Panel4.Visible = False
        '
        'UltraButton4
        '
        Appearance1.Image = CType(resources.GetObject("Appearance1.Image"), Object)
        Appearance1.ImageHAlign = Infragistics.Win.HAlign.Center
        Me.UltraButton4.Appearance = Appearance1
        Me.UltraButton4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton4.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton4.Location = New System.Drawing.Point(343, 10)
        Me.UltraButton4.Name = "UltraButton4"
        Me.UltraButton4.Size = New System.Drawing.Size(33, 34)
        Me.UltraButton4.TabIndex = 278
        '
        'txtC2
        '
        Me.txtC2.BackColor = System.Drawing.SystemColors.Window
        Me.txtC2.DateButtons.Add(DateButton1)
        Me.txtC2.Location = New System.Drawing.Point(243, 18)
        Me.txtC2.Name = "txtC2"
        Me.txtC2.NonAutoSizeHeight = 21
        Me.txtC2.Size = New System.Drawing.Size(100, 21)
        Me.txtC2.TabIndex = 275
        Me.txtC2.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(191, 21)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(46, 13)
        Me.Label12.TabIndex = 274
        Me.Label12.Text = "To Date"
        '
        'txtC1
        '
        Me.txtC1.BackColor = System.Drawing.SystemColors.Window
        Me.txtC1.DateButtons.Add(DateButton2)
        Me.txtC1.Location = New System.Drawing.Point(85, 18)
        Me.txtC1.Name = "txtC1"
        Me.txtC1.NonAutoSizeHeight = 21
        Me.txtC1.Size = New System.Drawing.Size(100, 21)
        Me.txtC1.TabIndex = 273
        Me.txtC1.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(3, 21)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(56, 13)
        Me.Label13.TabIndex = 0
        Me.Label13.Text = "From Date"
        '
        'frmVarions_rpt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(937, 425)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.UltraGrid1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Name = "frmVarions_rpt"
        Me.Text = "##Report View"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        CType(Me.txtC2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ReportFilterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UsingDateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrintToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ResetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents UltraButton4 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents txtC2 As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtC1 As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label13 As System.Windows.Forms.Label
End Class
