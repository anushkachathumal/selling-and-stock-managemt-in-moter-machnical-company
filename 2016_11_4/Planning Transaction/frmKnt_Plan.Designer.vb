<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmKnt_Plan
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
        Dim Appearance56 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.grp1 = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraLabel10 = New Infragistics.Win.Misc.UltraLabel
        Me.Panel1 = New System.Windows.Forms.Panel
        CType(Me.grp1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grp1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'grp1
        '
        Me.grp1.BorderStyle = Infragistics.Win.Misc.GroupBoxBorderStyle.Rectangular3D
        Me.grp1.Controls.Add(Me.UltraLabel10)
        Me.grp1.Location = New System.Drawing.Point(10, 0)
        Me.grp1.Name = "grp1"
        Me.grp1.Size = New System.Drawing.Size(795, 71)
        Me.grp1.TabIndex = 183
        '
        'UltraLabel10
        '
        Appearance56.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Appearance56.FontData.BoldAsString = "True"
        Appearance56.FontData.Name = "Times New Roman"
        Appearance56.TextHAlignAsString = "Center"
        Me.UltraLabel10.Appearance = Appearance56
        Me.UltraLabel10.BorderStyleOuter = Infragistics.Win.UIElementBorderStyle.Rounded4Thick
        Me.UltraLabel10.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel10.Location = New System.Drawing.Point(19, 16)
        Me.UltraLabel10.Name = "UltraLabel10"
        Me.UltraLabel10.Size = New System.Drawing.Size(94, 22)
        Me.UltraLabel10.TabIndex = 86
        Me.UltraLabel10.Text = "M/C Group"
        Me.UltraLabel10.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'Panel1
        '
        Me.Panel1.AutoScroll = True
        Me.Panel1.Controls.Add(Me.grp1)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(785, 79)
        Me.Panel1.TabIndex = 184
        '
        'frmKnt_Plan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(809, 262)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "frmKnt_Plan"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Knitting Plan"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.grp1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grp1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grp1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraLabel10 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
End Class
