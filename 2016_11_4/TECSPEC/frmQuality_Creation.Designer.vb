<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmQuality_Creation
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
        Dim Appearance473 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmQuality_Creation))
        Dim Appearance474 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraGrid4 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.lbl1 = New System.Windows.Forms.Label
        Me.UltraButton17 = New Infragistics.Win.Misc.UltraButton
        Me.UltraButton1 = New Infragistics.Win.Misc.UltraButton
        CType(Me.UltraGrid4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UltraGrid4
        '
        Me.UltraGrid4.Location = New System.Drawing.Point(12, 12)
        Me.UltraGrid4.Name = "UltraGrid4"
        Me.UltraGrid4.Size = New System.Drawing.Size(763, 429)
        Me.UltraGrid4.TabIndex = 209
        '
        'lbl1
        '
        Me.lbl1.AutoSize = True
        Me.lbl1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1.ForeColor = System.Drawing.Color.Red
        Me.lbl1.Location = New System.Drawing.Point(12, 453)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(55, 19)
        Me.lbl1.TabIndex = 210
        Me.lbl1.Text = "Label1"
        '
        'UltraButton17
        '
        Appearance473.Image = CType(resources.GetObject("Appearance473.Image"), Object)
        Me.UltraButton17.Appearance = Appearance473
        Me.UltraButton17.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton17.ImageSize = New System.Drawing.Size(22, 22)
        Me.UltraButton17.Location = New System.Drawing.Point(508, 447)
        Me.UltraButton17.Name = "UltraButton17"
        Me.UltraButton17.Size = New System.Drawing.Size(153, 36)
        Me.UltraButton17.TabIndex = 211
        Me.UltraButton17.Text = "&Create Spec"
        '
        'UltraButton1
        '
        Appearance474.Image = CType(resources.GetObject("Appearance474.Image"), Object)
        Me.UltraButton1.Appearance = Appearance474
        Me.UltraButton1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton1.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton1.Location = New System.Drawing.Point(667, 447)
        Me.UltraButton1.Name = "UltraButton1"
        Me.UltraButton1.Size = New System.Drawing.Size(108, 36)
        Me.UltraButton1.TabIndex = 212
        Me.UltraButton1.Text = "&Exit"
        '
        'frmQuality_Creation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(777, 492)
        Me.Controls.Add(Me.UltraButton1)
        Me.Controls.Add(Me.UltraButton17)
        Me.Controls.Add(Me.lbl1)
        Me.Controls.Add(Me.UltraGrid4)
        Me.Name = "frmQuality_Creation"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "##"
        CType(Me.UltraGrid4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents UltraGrid4 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents UltraButton17 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraButton1 As Infragistics.Win.Misc.UltraButton
End Class
