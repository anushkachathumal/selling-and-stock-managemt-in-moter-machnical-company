<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFile_Upload
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
        Dim Appearance57 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFile_Upload))
        Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance26 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraButton2 = New Infragistics.Win.Misc.UltraButton
        Me.UltraButton6 = New Infragistics.Win.Misc.UltraButton
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.cmdReset = New Infragistics.Win.Misc.UltraButton
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.UltraGroupBox4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.chkYarn_Code = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'UltraButton2
        '
        Appearance57.Image = CType(resources.GetObject("Appearance57.Image"), Object)
        Me.UltraButton2.Appearance = Appearance57
        Me.UltraButton2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.FlatBorderless
        Me.UltraButton2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton2.ImageSize = New System.Drawing.Size(40, 40)
        Me.UltraButton2.Location = New System.Drawing.Point(12, 481)
        Me.UltraButton2.Name = "UltraButton2"
        Me.UltraButton2.Size = New System.Drawing.Size(111, 37)
        Me.UltraButton2.TabIndex = 198
        Me.UltraButton2.Text = "Upload"
        '
        'UltraButton6
        '
        Appearance11.Image = CType(resources.GetObject("Appearance11.Image"), Object)
        Me.UltraButton6.Appearance = Appearance11
        Me.UltraButton6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton6.ImageSize = New System.Drawing.Size(32, 32)
        Me.UltraButton6.Location = New System.Drawing.Point(129, 480)
        Me.UltraButton6.Name = "UltraButton6"
        Me.UltraButton6.Size = New System.Drawing.Size(89, 37)
        Me.UltraButton6.TabIndex = 197
        Me.UltraButton6.Text = "&Save"
        '
        'cmdExit
        '
        Appearance9.Image = CType(resources.GetObject("Appearance9.Image"), Object)
        Me.cmdExit.Appearance = Appearance9
        Me.cmdExit.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdExit.Location = New System.Drawing.Point(319, 480)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(109, 36)
        Me.cmdExit.TabIndex = 195
        Me.cmdExit.Text = "&Exit"
        '
        'cmdReset
        '
        Appearance26.Image = CType(resources.GetObject("Appearance26.Image"), Object)
        Me.cmdReset.Appearance = Appearance26
        Me.cmdReset.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReset.ImageSize = New System.Drawing.Size(32, 32)
        Me.cmdReset.Location = New System.Drawing.Point(224, 480)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.Size = New System.Drawing.Size(89, 36)
        Me.cmdReset.TabIndex = 196
        Me.cmdReset.Text = "&Reset"
        '
        'UltraGrid1
        '
        Me.UltraGrid1.Location = New System.Drawing.Point(12, 65)
        Me.UltraGrid1.Name = "UltraGrid1"
        Me.UltraGrid1.Size = New System.Drawing.Size(808, 409)
        Me.UltraGrid1.TabIndex = 199
        '
        'UltraGroupBox4
        '
        Me.UltraGroupBox4.Controls.Add(Me.chkYarn_Code)
        Me.UltraGroupBox4.Location = New System.Drawing.Point(12, 12)
        Me.UltraGroupBox4.Name = "UltraGroupBox4"
        Me.UltraGroupBox4.Size = New System.Drawing.Size(328, 51)
        Me.UltraGroupBox4.TabIndex = 200
        Me.UltraGroupBox4.Text = " "
        '
        'chkYarn_Code
        '
        Appearance19.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Appearance19.FontData.BoldAsString = "False"
        Me.chkYarn_Code.Appearance = Appearance19
        Me.chkYarn_Code.BackColor = System.Drawing.Color.Transparent
        Me.chkYarn_Code.BackColorInternal = System.Drawing.Color.Transparent
        Me.chkYarn_Code.Location = New System.Drawing.Point(6, 21)
        Me.chkYarn_Code.Name = "chkYarn_Code"
        Me.chkYarn_Code.Size = New System.Drawing.Size(144, 20)
        Me.chkYarn_Code.TabIndex = 94
        Me.chkYarn_Code.Text = "Yarn Stock Code"
        '
        'frmFile_Upload
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(837, 553)
        Me.Controls.Add(Me.UltraGroupBox4)
        Me.Controls.Add(Me.UltraGrid1)
        Me.Controls.Add(Me.UltraButton2)
        Me.Controls.Add(Me.UltraButton6)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdReset)
        Me.Name = "frmFile_Upload"
        Me.Text = "File Upload"
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraButton2 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraButton6 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdReset As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents UltraGroupBox4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents chkYarn_Code As Infragistics.Win.UltraWinEditors.UltraCheckEditor
End Class
