<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMOQ
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
        Dim Appearance32 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMOQ))
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance35 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance37 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraLabel8 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraGroupBox5 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdSave = New Infragistics.Win.Misc.UltraButton
        Me.cmdReset = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraButton7 = New Infragistics.Win.Misc.UltraButton
        Me.lblPro = New Infragistics.Win.Misc.UltraLabel
        Me.pbCount = New Infragistics.Win.UltraWinProgressBar.UltraProgressBar
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox5.SuspendLayout()
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox4.SuspendLayout()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox1.SuspendLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UltraLabel8
        '
        Appearance32.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel8.Appearance = Appearance32
        Me.UltraLabel8.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel8.Location = New System.Drawing.Point(862, 97)
        Me.UltraLabel8.Name = "UltraLabel8"
        Me.UltraLabel8.Size = New System.Drawing.Size(299, 41)
        Me.UltraLabel8.TabIndex = 162
        Me.UltraLabel8.Text = "MOQ"
        Me.UltraLabel8.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'UltraGroupBox5
        '
        Me.UltraGroupBox5.Controls.Add(Me.cmdExit)
        Me.UltraGroupBox5.Location = New System.Drawing.Point(263, 452)
        Me.UltraGroupBox5.Name = "UltraGroupBox5"
        Me.UltraGroupBox5.Size = New System.Drawing.Size(142, 69)
        Me.UltraGroupBox5.TabIndex = 161
        '
        'cmdExit
        '
        Appearance4.Image = CType(resources.GetObject("Appearance4.Image"), Object)
        Me.cmdExit.Appearance = Appearance4
        Me.cmdExit.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdExit.Location = New System.Drawing.Point(6, 13)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(130, 47)
        Me.cmdExit.TabIndex = 3
        Me.cmdExit.Text = "&Exit"
        '
        'UltraGroupBox4
        '
        Me.UltraGroupBox4.Controls.Add(Me.cmdSave)
        Me.UltraGroupBox4.Controls.Add(Me.cmdReset)
        Me.UltraGroupBox4.Location = New System.Drawing.Point(16, 452)
        Me.UltraGroupBox4.Name = "UltraGroupBox4"
        Me.UltraGroupBox4.Size = New System.Drawing.Size(241, 69)
        Me.UltraGroupBox4.TabIndex = 160
        '
        'cmdSave
        '
        Appearance20.Image = CType(resources.GetObject("Appearance20.Image"), Object)
        Me.cmdSave.Appearance = Appearance20
        Me.cmdSave.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ImageSize = New System.Drawing.Size(32, 32)
        Me.cmdSave.Location = New System.Drawing.Point(8, 11)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(108, 49)
        Me.cmdSave.TabIndex = 6
        Me.cmdSave.Text = "&Save"
        '
        'cmdReset
        '
        Appearance16.Image = CType(resources.GetObject("Appearance16.Image"), Object)
        Me.cmdReset.Appearance = Appearance16
        Me.cmdReset.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReset.ImageSize = New System.Drawing.Size(32, 32)
        Me.cmdReset.Location = New System.Drawing.Point(122, 12)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.Size = New System.Drawing.Size(108, 49)
        Me.cmdReset.TabIndex = 5
        Me.cmdReset.Text = "&Reset"
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.Controls.Add(Me.UltraGrid1)
        Me.UltraGroupBox1.Location = New System.Drawing.Point(16, 57)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(555, 340)
        Me.UltraGroupBox1.TabIndex = 159
        Me.UltraGroupBox1.Text = "MOQ"
        '
        'UltraGrid1
        '
        Me.UltraGrid1.Location = New System.Drawing.Point(8, 21)
        Me.UltraGrid1.Name = "UltraGrid1"
        Me.UltraGrid1.Size = New System.Drawing.Size(532, 310)
        Me.UltraGrid1.TabIndex = 1
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(665, 10)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(191, 154)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 164
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(862, 57)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(102, 34)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 163
        Me.PictureBox2.TabStop = False
        '
        'UltraLabel1
        '
        Appearance18.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Appearance18.ForeColor = System.Drawing.Color.Blue
        Me.UltraLabel1.Appearance = Appearance18
        Me.UltraLabel1.Location = New System.Drawing.Point(16, 34)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(129, 19)
        Me.UltraLabel1.TabIndex = 166
        Me.UltraLabel1.Text = "Upload Bulk Data"
        Me.UltraLabel1.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'UltraButton7
        '
        Me.UltraButton7.ButtonStyle = Infragistics.Win.UIElementButtonStyle.FlatBorderless
        Me.UltraButton7.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraButton7.ImageSize = New System.Drawing.Size(40, 40)
        Me.UltraButton7.Location = New System.Drawing.Point(16, 12)
        Me.UltraButton7.Name = "UltraButton7"
        Me.UltraButton7.Size = New System.Drawing.Size(26, 20)
        Me.UltraButton7.TabIndex = 165
        Me.UltraButton7.Text = "..."
        '
        'lblPro
        '
        Appearance35.BackColor = System.Drawing.Color.White
        Appearance35.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.lblPro.Appearance = Appearance35
        Me.lblPro.Location = New System.Drawing.Point(16, 427)
        Me.lblPro.Name = "lblPro"
        Me.lblPro.Size = New System.Drawing.Size(540, 19)
        Me.lblPro.TabIndex = 168
        Me.lblPro.Text = "Progress ...."
        Me.lblPro.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'pbCount
        '
        Appearance37.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Appearance37.FontData.BoldAsString = "True"
        Appearance37.ForeColorDisabled = System.Drawing.Color.Black
        Me.pbCount.Appearance = Appearance37
        Appearance23.FontData.BoldAsString = "True"
        Me.pbCount.FillAppearance = Appearance23
        Me.pbCount.Location = New System.Drawing.Point(16, 403)
        Me.pbCount.Name = "pbCount"
        Me.pbCount.Size = New System.Drawing.Size(540, 21)
        Me.pbCount.TabIndex = 167
        Me.pbCount.Text = "[Formatted]"
        '
        'frmMOQ
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1059, 533)
        Me.Controls.Add(Me.lblPro)
        Me.Controls.Add(Me.pbCount)
        Me.Controls.Add(Me.UltraLabel1)
        Me.Controls.Add(Me.UltraButton7)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.UltraLabel8)
        Me.Controls.Add(Me.UltraGroupBox5)
        Me.Controls.Add(Me.UltraGroupBox4)
        Me.Controls.Add(Me.UltraGroupBox1)
        Me.Name = "frmMOQ"
        Me.Text = "MOQ"
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox5.ResumeLayout(False)
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox4.ResumeLayout(False)
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox1.ResumeLayout(False)
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents UltraLabel8 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraGroupBox5 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdSave As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdReset As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraButton7 As Infragistics.Win.Misc.UltraButton
    Friend WithEvents lblPro As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents pbCount As Infragistics.Win.UltraWinProgressBar.UltraProgressBar
End Class
