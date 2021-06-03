<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmYrn_Material
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
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance542 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmYrn_Material))
        Dim Appearance543 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance544 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance102 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.OPR4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtVoucher = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel
        Me.txtDis = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraButton1 = New Infragistics.Win.Misc.UltraButton
        Me.UltraButton2 = New Infragistics.Win.Misc.UltraButton
        Me.UltraButton3 = New Infragistics.Win.Misc.UltraButton
        Me.UltraButton4 = New Infragistics.Win.Misc.UltraButton
        CType(Me.OPR4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR4.SuspendLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.txtVoucher, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDis, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OPR4
        '
        Me.OPR4.Controls.Add(Me.UltraGrid1)
        Me.OPR4.Location = New System.Drawing.Point(22, 108)
        Me.OPR4.Name = "OPR4"
        Me.OPR4.Size = New System.Drawing.Size(641, 226)
        Me.OPR4.TabIndex = 122
        '
        'UltraGrid1
        '
        Me.UltraGrid1.Location = New System.Drawing.Point(6, 19)
        Me.UltraGrid1.Name = "UltraGrid1"
        Me.UltraGrid1.Size = New System.Drawing.Size(629, 201)
        Me.UltraGrid1.TabIndex = 1
        '
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.txtVoucher)
        Me.OPR0.Controls.Add(Me.UltraLabel4)
        Me.OPR0.Controls.Add(Me.txtDis)
        Me.OPR0.Controls.Add(Me.UltraLabel2)
        Me.OPR0.Location = New System.Drawing.Point(23, 12)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(640, 92)
        Me.OPR0.TabIndex = 121
        '
        'txtVoucher
        '
        Me.txtVoucher.Location = New System.Drawing.Point(113, 42)
        Me.txtVoucher.MaxLength = 60
        Me.txtVoucher.Name = "txtVoucher"
        Me.txtVoucher.Size = New System.Drawing.Size(519, 21)
        Me.txtVoucher.TabIndex = 57
        '
        'UltraLabel4
        '
        Appearance1.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel4.Appearance = Appearance1
        Me.UltraLabel4.Location = New System.Drawing.Point(15, 42)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel4.TabIndex = 56
        Me.UltraLabel4.Text = "Description"
        Me.UltraLabel4.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtDis
        '
        Me.txtDis.Location = New System.Drawing.Point(113, 15)
        Me.txtDis.MaxLength = 15
        Me.txtDis.Name = "txtDis"
        Me.txtDis.Size = New System.Drawing.Size(127, 21)
        Me.txtDis.TabIndex = 51
        '
        'UltraLabel2
        '
        Appearance8.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel2.Appearance = Appearance8
        Me.UltraLabel2.Location = New System.Drawing.Point(15, 15)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(109, 21)
        Me.UltraLabel2.TabIndex = 36
        Me.UltraLabel2.Text = "Yarn Code"
        Me.UltraLabel2.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'UltraButton1
        '
        Appearance542.Image = CType(resources.GetObject("Appearance542.Image"), Object)
        Me.UltraButton1.Appearance = Appearance542
        Me.UltraButton1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton1.ImageSize = New System.Drawing.Size(22, 22)
        Me.UltraButton1.Location = New System.Drawing.Point(22, 341)
        Me.UltraButton1.Name = "UltraButton1"
        Me.UltraButton1.Size = New System.Drawing.Size(134, 36)
        Me.UltraButton1.TabIndex = 251
        Me.UltraButton1.Text = "&Save"
        '
        'UltraButton2
        '
        Appearance543.Image = CType(resources.GetObject("Appearance543.Image"), Object)
        Me.UltraButton2.Appearance = Appearance543
        Me.UltraButton2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton2.ImageSize = New System.Drawing.Size(32, 32)
        Me.UltraButton2.Location = New System.Drawing.Point(302, 340)
        Me.UltraButton2.Name = "UltraButton2"
        Me.UltraButton2.Size = New System.Drawing.Size(134, 35)
        Me.UltraButton2.TabIndex = 250
        Me.UltraButton2.Text = "&Reset"
        '
        'UltraButton3
        '
        Appearance544.Image = CType(resources.GetObject("Appearance544.Image"), Object)
        Me.UltraButton3.Appearance = Appearance544
        Me.UltraButton3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton3.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton3.Location = New System.Drawing.Point(529, 340)
        Me.UltraButton3.Name = "UltraButton3"
        Me.UltraButton3.Size = New System.Drawing.Size(134, 34)
        Me.UltraButton3.TabIndex = 249
        Me.UltraButton3.Text = "&Exit"
        '
        'UltraButton4
        '
        Appearance102.Image = CType(resources.GetObject("Appearance102.Image"), Object)
        Me.UltraButton4.Appearance = Appearance102
        Me.UltraButton4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton4.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton4.Location = New System.Drawing.Point(162, 341)
        Me.UltraButton4.Name = "UltraButton4"
        Me.UltraButton4.Size = New System.Drawing.Size(134, 36)
        Me.UltraButton4.TabIndex = 252
        Me.UltraButton4.Text = "&Delete"
        '
        'frmYrn_Material
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(781, 443)
        Me.Controls.Add(Me.UltraButton4)
        Me.Controls.Add(Me.UltraButton1)
        Me.Controls.Add(Me.UltraButton2)
        Me.Controls.Add(Me.UltraButton3)
        Me.Controls.Add(Me.OPR4)
        Me.Controls.Add(Me.OPR0)
        Me.Name = "frmYrn_Material"
        Me.Text = "Yarn Material"
        CType(Me.OPR4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR4.ResumeLayout(False)
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.txtVoucher, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDis, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents OPR4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtVoucher As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtDis As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraButton1 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraButton2 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraButton3 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraButton4 As Infragistics.Win.Misc.UltraButton
End Class
