<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDiscount_1
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
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDiscount_1))
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtDiscount = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.txtAmount = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraGrid2 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.cmdReset = New Infragistics.Win.Misc.UltraButton
        Me.cmdAdd = New Infragistics.Win.Misc.UltraButton
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox1.SuspendLayout()
        CType(Me.txtDiscount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAmount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.Controls.Add(Me.txtDiscount)
        Me.UltraGroupBox1.Controls.Add(Me.UltraLabel2)
        Me.UltraGroupBox1.Controls.Add(Me.UltraLabel1)
        Me.UltraGroupBox1.Controls.Add(Me.txtAmount)
        Me.UltraGroupBox1.Location = New System.Drawing.Point(23, 25)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(432, 55)
        Me.UltraGroupBox1.TabIndex = 112
        '
        'txtDiscount
        '
        Appearance3.ForeColor = System.Drawing.Color.Black
        Me.txtDiscount.Appearance = Appearance3
        Me.txtDiscount.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDiscount.Location = New System.Drawing.Point(345, 15)
        Me.txtDiscount.MaxLength = 30
        Me.txtDiscount.Name = "txtDiscount"
        Me.txtDiscount.Size = New System.Drawing.Size(69, 21)
        Me.txtDiscount.TabIndex = 50
        '
        'UltraLabel2
        '
        Appearance18.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel2.Appearance = Appearance18
        Me.UltraLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel2.Location = New System.Drawing.Point(256, 17)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel2.TabIndex = 49
        Me.UltraLabel2.Text = "Discount (%)"
        Me.UltraLabel2.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'UltraLabel1
        '
        Appearance6.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel1.Appearance = Appearance6
        Me.UltraLabel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel1.Location = New System.Drawing.Point(6, 17)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(135, 21)
        Me.UltraLabel1.TabIndex = 48
        Me.UltraLabel1.Text = "Monthly Invoice Amount"
        Me.UltraLabel1.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtAmount
        '
        Appearance21.ForeColor = System.Drawing.Color.Black
        Me.txtAmount.Appearance = Appearance21
        Me.txtAmount.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAmount.Location = New System.Drawing.Point(147, 15)
        Me.txtAmount.MaxLength = 30
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(103, 21)
        Me.txtAmount.TabIndex = 1
        '
        'UltraGrid2
        '
        Me.UltraGrid2.Location = New System.Drawing.Point(23, 86)
        Me.UltraGrid2.Name = "UltraGrid2"
        Me.UltraGrid2.Size = New System.Drawing.Size(432, 218)
        Me.UltraGrid2.TabIndex = 113
        '
        'cmdExit
        '
        Appearance17.Image = CType(resources.GetObject("Appearance17.Image"), Object)
        Me.cmdExit.Appearance = Appearance17
        Me.cmdExit.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdExit.Location = New System.Drawing.Point(257, 310)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(113, 47)
        Me.cmdExit.TabIndex = 115
        Me.cmdExit.Text = "&Exit"
        '
        'cmdReset
        '
        Appearance2.Image = CType(resources.GetObject("Appearance2.Image"), Object)
        Me.cmdReset.Appearance = Appearance2
        Me.cmdReset.ImageSize = New System.Drawing.Size(32, 32)
        Me.cmdReset.Location = New System.Drawing.Point(140, 310)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.Size = New System.Drawing.Size(113, 47)
        Me.cmdReset.TabIndex = 116
        Me.cmdReset.Text = "&Reset"
        '
        'cmdAdd
        '
        Appearance1.Image = CType(resources.GetObject("Appearance1.Image"), Object)
        Appearance1.TextHAlignAsString = "Center"
        Me.cmdAdd.Appearance = Appearance1
        Me.cmdAdd.ImageSize = New System.Drawing.Size(32, 32)
        Me.cmdAdd.Location = New System.Drawing.Point(23, 310)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(113, 47)
        Me.cmdAdd.TabIndex = 114
        Me.cmdAdd.Text = "&Save"
        '
        'frmDiscount_1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(479, 375)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdReset)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.UltraGrid2)
        Me.Controls.Add(Me.UltraGroupBox1)
        Me.MaximizeBox = False
        Me.Name = "frmDiscount_1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "#Advance Discount (Company Discount)"
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox1.ResumeLayout(False)
        Me.UltraGroupBox1.PerformLayout()
        CType(Me.txtDiscount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAmount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGrid2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtDiscount As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtAmount As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraGrid2 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdReset As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdAdd As Infragistics.Win.Misc.UltraButton
End Class
