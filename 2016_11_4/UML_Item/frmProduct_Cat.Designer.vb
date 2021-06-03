<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProduct_Cat
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
        Dim Appearance542 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProduct_Cat))
        Dim Appearance102 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance543 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance544 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance31 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.cmdSave = New Infragistics.Win.Misc.UltraButton
        Me.UltraButton4 = New Infragistics.Win.Misc.UltraButton
        Me.UltraButton2 = New Infragistics.Win.Misc.UltraButton
        Me.UltraButton3 = New Infragistics.Win.Misc.UltraButton
        Me.OPR4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtName = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.txtCode = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel
        CType(Me.OPR4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR4.SuspendLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.txtName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCode, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdSave
        '
        Appearance542.Image = CType(resources.GetObject("Appearance542.Image"), Object)
        Me.cmdSave.Appearance = Appearance542
        Me.cmdSave.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ImageSize = New System.Drawing.Size(22, 22)
        Me.cmdSave.Location = New System.Drawing.Point(22, 343)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(134, 49)
        Me.cmdSave.TabIndex = 269
        Me.cmdSave.Text = "&Save"
        '
        'UltraButton4
        '
        Appearance102.Image = CType(resources.GetObject("Appearance102.Image"), Object)
        Me.UltraButton4.Appearance = Appearance102
        Me.UltraButton4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton4.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton4.Location = New System.Drawing.Point(162, 343)
        Me.UltraButton4.Name = "UltraButton4"
        Me.UltraButton4.Size = New System.Drawing.Size(134, 49)
        Me.UltraButton4.TabIndex = 268
        Me.UltraButton4.Text = "&Delete"
        '
        'UltraButton2
        '
        Appearance543.Image = CType(resources.GetObject("Appearance543.Image"), Object)
        Me.UltraButton2.Appearance = Appearance543
        Me.UltraButton2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton2.ImageSize = New System.Drawing.Size(32, 32)
        Me.UltraButton2.Location = New System.Drawing.Point(302, 343)
        Me.UltraButton2.Name = "UltraButton2"
        Me.UltraButton2.Size = New System.Drawing.Size(134, 49)
        Me.UltraButton2.TabIndex = 267
        Me.UltraButton2.Text = "&Reset"
        '
        'UltraButton3
        '
        Appearance544.Image = CType(resources.GetObject("Appearance544.Image"), Object)
        Me.UltraButton3.Appearance = Appearance544
        Me.UltraButton3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton3.ImageSize = New System.Drawing.Size(20, 20)
        Me.UltraButton3.Location = New System.Drawing.Point(529, 343)
        Me.UltraButton3.Name = "UltraButton3"
        Me.UltraButton3.Size = New System.Drawing.Size(134, 49)
        Me.UltraButton3.TabIndex = 266
        Me.UltraButton3.Text = "&Exit"
        '
        'OPR4
        '
        Me.OPR4.Controls.Add(Me.UltraGrid1)
        Me.OPR4.Location = New System.Drawing.Point(22, 107)
        Me.OPR4.Name = "OPR4"
        Me.OPR4.Size = New System.Drawing.Size(641, 226)
        Me.OPR4.TabIndex = 265
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
        Me.OPR0.Controls.Add(Me.txtName)
        Me.OPR0.Controls.Add(Me.UltraLabel2)
        Me.OPR0.Controls.Add(Me.txtCode)
        Me.OPR0.Controls.Add(Me.UltraLabel4)
        Me.OPR0.Location = New System.Drawing.Point(22, 12)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(641, 86)
        Me.OPR0.TabIndex = 264
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(104, 43)
        Me.txtName.MaxLength = 120
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(531, 21)
        Me.txtName.TabIndex = 1
        '
        'UltraLabel2
        '
        Appearance1.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel2.Appearance = Appearance1
        Me.UltraLabel2.Location = New System.Drawing.Point(6, 43)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel2.TabIndex = 115
        Me.UltraLabel2.Text = "Category Name"
        Me.UltraLabel2.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtCode
        '
        Me.txtCode.Location = New System.Drawing.Point(104, 16)
        Me.txtCode.MaxLength = 15
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(101, 21)
        Me.txtCode.TabIndex = 2
        '
        'UltraLabel4
        '
        Appearance31.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel4.Appearance = Appearance31
        Me.UltraLabel4.Location = New System.Drawing.Point(6, 16)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel4.TabIndex = 56
        Me.UltraLabel4.Text = "Cat Code"
        Me.UltraLabel4.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'frmProduct_Cat
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(853, 464)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.UltraButton4)
        Me.Controls.Add(Me.UltraButton2)
        Me.Controls.Add(Me.UltraButton3)
        Me.Controls.Add(Me.OPR4)
        Me.Controls.Add(Me.OPR0)
        Me.Name = "frmProduct_Cat"
        Me.Text = "Product Category"
        CType(Me.OPR4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR4.ResumeLayout(False)
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.txtName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCode, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdSave As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraButton4 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraButton2 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraButton3 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents OPR4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtName As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtCode As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
End Class
