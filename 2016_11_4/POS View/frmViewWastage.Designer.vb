<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmViewWastage
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
        Dim DateButton3 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim DateButton4 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmViewWastage))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.txtTo = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtFrom = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.Label1 = New System.Windows.Forms.Label
        Me.UltraButton3 = New Infragistics.Win.Misc.UltraButton
        Me.UltraGrid3 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.FilterByToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ByDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ResetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Panel1.SuspendLayout()
        CType(Me.txtTo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtFrom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGrid3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.txtTo)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.txtFrom)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.UltraButton3)
        Me.Panel1.Location = New System.Drawing.Point(191, 182)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(382, 61)
        Me.Panel1.TabIndex = 114
        Me.Panel1.Visible = False
        '
        'txtTo
        '
        Me.txtTo.BackColor = System.Drawing.SystemColors.Window
        Me.txtTo.DateButtons.Add(DateButton3)
        Me.txtTo.Location = New System.Drawing.Point(235, 19)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.NonAutoSizeHeight = 21
        Me.txtTo.Size = New System.Drawing.Size(100, 21)
        Me.txtTo.TabIndex = 276
        Me.txtTo.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(183, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 13)
        Me.Label2.TabIndex = 275
        Me.Label2.Text = "To Date"
        '
        'txtFrom
        '
        Me.txtFrom.BackColor = System.Drawing.SystemColors.Window
        Me.txtFrom.DateButtons.Add(DateButton4)
        Me.txtFrom.Location = New System.Drawing.Point(77, 19)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.NonAutoSizeHeight = 21
        Me.txtFrom.Size = New System.Drawing.Size(100, 21)
        Me.txtFrom.TabIndex = 274
        Me.txtFrom.Value = New Date(2017, 9, 4, 0, 0, 0, 0)
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 13)
        Me.Label1.TabIndex = 262
        Me.Label1.Text = "From Date"
        '
        'UltraButton3
        '
        Appearance1.Image = CType(resources.GetObject("Appearance1.Image"), Object)
        Appearance1.ImageHAlign = Infragistics.Win.HAlign.Center
        Me.UltraButton3.Appearance = Appearance1
        Me.UltraButton3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton3.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton3.Location = New System.Drawing.Point(341, 11)
        Me.UltraButton3.Name = "UltraButton3"
        Me.UltraButton3.Size = New System.Drawing.Size(33, 34)
        Me.UltraButton3.TabIndex = 261
        '
        'UltraGrid3
        '
        Me.UltraGrid3.Location = New System.Drawing.Point(15, 48)
        Me.UltraGrid3.Name = "UltraGrid3"
        Me.UltraGrid3.Size = New System.Drawing.Size(745, 364)
        Me.UltraGrid3.TabIndex = 112
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FilterByToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(774, 24)
        Me.MenuStrip1.TabIndex = 115
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FilterByToolStripMenuItem
        '
        Me.FilterByToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ByDateToolStripMenuItem, Me.ResetToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.FilterByToolStripMenuItem.Name = "FilterByToolStripMenuItem"
        Me.FilterByToolStripMenuItem.Size = New System.Drawing.Size(61, 20)
        Me.FilterByToolStripMenuItem.Text = "Filter by"
        '
        'ByDateToolStripMenuItem
        '
        Me.ByDateToolStripMenuItem.Name = "ByDateToolStripMenuItem"
        Me.ByDateToolStripMenuItem.Size = New System.Drawing.Size(114, 22)
        Me.ByDateToolStripMenuItem.Text = "by Date"
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
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(114, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'frmViewWastage
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(774, 428)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.UltraGrid3)
        Me.Name = "frmViewWastage"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "##View Wastage"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.txtTo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtFrom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGrid3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtTo As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtFrom As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents UltraButton3 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGrid3 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FilterByToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ByDateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ResetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
