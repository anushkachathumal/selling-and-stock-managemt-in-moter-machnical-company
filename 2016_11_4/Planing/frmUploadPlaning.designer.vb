<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUploadPlaning
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
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmUploadPlaning))
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraGroupBox5 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox3 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdDelete = New Infragistics.Win.Misc.UltraButton
        Me.cmdEdit = New Infragistics.Win.Misc.UltraButton
        Me.cmdAdd = New Infragistics.Win.Misc.UltraButton
        Me.OPR1 = New Infragistics.Win.Misc.UltraGroupBox
        Me.lblDis = New Infragistics.Win.Misc.UltraLabel
        Me.pbCount = New Infragistics.Win.UltraWinProgressBar.UltraProgressBar
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox5.SuspendLayout()
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox3.SuspendLayout()
        CType(Me.OPR1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR1.SuspendLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UltraGroupBox5
        '
        Me.UltraGroupBox5.Controls.Add(Me.cmdExit)
        Me.UltraGroupBox5.Location = New System.Drawing.Point(314, 500)
        Me.UltraGroupBox5.Name = "UltraGroupBox5"
        Me.UltraGroupBox5.Size = New System.Drawing.Size(133, 56)
        Me.UltraGroupBox5.TabIndex = 154
        '
        'cmdExit
        '
        Appearance21.Image = CType(resources.GetObject("Appearance21.Image"), Object)
        Me.cmdExit.Appearance = Appearance21
        Me.cmdExit.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdExit.Location = New System.Drawing.Point(6, 10)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(121, 40)
        Me.cmdExit.TabIndex = 3
        Me.cmdExit.Text = "&Exit"
        '
        'UltraGroupBox3
        '
        Me.UltraGroupBox3.Controls.Add(Me.cmdDelete)
        Me.UltraGroupBox3.Controls.Add(Me.cmdEdit)
        Me.UltraGroupBox3.Controls.Add(Me.cmdAdd)
        Me.UltraGroupBox3.Location = New System.Drawing.Point(18, 500)
        Me.UltraGroupBox3.Name = "UltraGroupBox3"
        Me.UltraGroupBox3.Size = New System.Drawing.Size(290, 58)
        Me.UltraGroupBox3.TabIndex = 153
        '
        'cmdDelete
        '
        Appearance19.Image = CType(resources.GetObject("Appearance19.Image"), Object)
        Me.cmdDelete.Appearance = Appearance19
        Me.cmdDelete.Enabled = False
        Me.cmdDelete.Location = New System.Drawing.Point(194, 12)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(85, 40)
        Me.cmdDelete.TabIndex = 2
        Me.cmdDelete.Text = "&Reset"
        '
        'cmdEdit
        '
        Appearance17.Image = CType(resources.GetObject("Appearance17.Image"), Object)
        Me.cmdEdit.Appearance = Appearance17
        Me.cmdEdit.Enabled = False
        Me.cmdEdit.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdEdit.Location = New System.Drawing.Point(100, 12)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(85, 40)
        Me.cmdEdit.TabIndex = 1
        Me.cmdEdit.Text = "&Save"
        '
        'cmdAdd
        '
        Appearance18.Image = CType(resources.GetObject("Appearance18.Image"), Object)
        Appearance18.TextHAlignAsString = "Center"
        Me.cmdAdd.Appearance = Appearance18
        Me.cmdAdd.ImageSize = New System.Drawing.Size(20, 20)
        Me.cmdAdd.Location = New System.Drawing.Point(6, 11)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(85, 41)
        Me.cmdAdd.TabIndex = 0
        Me.cmdAdd.Text = "&Upload"
        '
        'OPR1
        '
        Me.OPR1.Controls.Add(Me.lblDis)
        Me.OPR1.Controls.Add(Me.pbCount)
        Me.OPR1.Enabled = False
        Me.OPR1.Location = New System.Drawing.Point(18, 430)
        Me.OPR1.Name = "OPR1"
        Me.OPR1.Size = New System.Drawing.Size(814, 64)
        Me.OPR1.TabIndex = 152
        '
        'lblDis
        '
        Appearance3.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.lblDis.Appearance = Appearance3
        Me.lblDis.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDis.Location = New System.Drawing.Point(6, 43)
        Me.lblDis.Name = "lblDis"
        Me.lblDis.Size = New System.Drawing.Size(412, 21)
        Me.lblDis.TabIndex = 51
        Me.lblDis.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'pbCount
        '
        Appearance22.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Appearance22.FontData.BoldAsString = "True"
        Appearance22.ForeColorDisabled = System.Drawing.Color.Black
        Appearance22.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent
        Me.pbCount.Appearance = Appearance22
        Appearance23.FontData.BoldAsString = "True"
        Me.pbCount.FillAppearance = Appearance23
        Me.pbCount.Location = New System.Drawing.Point(7, 17)
        Me.pbCount.Maximum = 120
        Me.pbCount.Name = "pbCount"
        Me.pbCount.Size = New System.Drawing.Size(801, 21)
        Me.pbCount.TabIndex = 50
        Me.pbCount.Text = "[Formatted]"
        '
        'UltraGrid1
        '
        Me.UltraGrid1.Location = New System.Drawing.Point(24, 15)
        Me.UltraGrid1.Name = "UltraGrid1"
        Me.UltraGrid1.Size = New System.Drawing.Size(808, 409)
        Me.UltraGrid1.TabIndex = 151
        '
        'frmUploadPlaning
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(867, 559)
        Me.Controls.Add(Me.UltraGroupBox5)
        Me.Controls.Add(Me.UltraGroupBox3)
        Me.Controls.Add(Me.OPR1)
        Me.Controls.Add(Me.UltraGrid1)
        Me.Name = "frmUploadPlaning"
        Me.Text = "Upload Records"
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox5.ResumeLayout(False)
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox3.ResumeLayout(False)
        CType(Me.OPR1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR1.ResumeLayout(False)
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraGroupBox5 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox3 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdDelete As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdEdit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdAdd As Infragistics.Win.Misc.UltraButton
    Friend WithEvents OPR1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents lblDis As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents pbCount As Infragistics.Win.UltraWinProgressBar.UltraProgressBar
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
End Class
