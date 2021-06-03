<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBackup
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBackup))
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance13 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance64 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraGroupBox5 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdReset = New Infragistics.Win.Misc.UltraButton
        Me.cmdSave = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox3 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdEdit = New Infragistics.Win.Misc.UltraButton
        Me.cmdAdd = New Infragistics.Win.Misc.UltraButton
        Me.OPR2 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtTo = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel5 = New Infragistics.Win.Misc.UltraLabel
        Me.txtFrom = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox5.SuspendLayout()
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox4.SuspendLayout()
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox3.SuspendLayout()
        CType(Me.OPR2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR2.SuspendLayout()
        CType(Me.txtTo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtFrom, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UltraGroupBox5
        '
        Me.UltraGroupBox5.Controls.Add(Me.cmdExit)
        Me.UltraGroupBox5.Location = New System.Drawing.Point(424, 106)
        Me.UltraGroupBox5.Name = "UltraGroupBox5"
        Me.UltraGroupBox5.Size = New System.Drawing.Size(181, 50)
        Me.UltraGroupBox5.TabIndex = 106
        '
        'cmdExit
        '
        Appearance21.Image = CType(resources.GetObject("Appearance21.Image"), Object)
        Me.cmdExit.Appearance = Appearance21
        Me.cmdExit.Location = New System.Drawing.Point(6, 10)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(169, 30)
        Me.cmdExit.TabIndex = 3
        Me.cmdExit.Text = "&Exit"
        '
        'UltraGroupBox4
        '
        Me.UltraGroupBox4.Controls.Add(Me.cmdReset)
        Me.UltraGroupBox4.Controls.Add(Me.cmdSave)
        Me.UltraGroupBox4.Location = New System.Drawing.Point(218, 106)
        Me.UltraGroupBox4.Name = "UltraGroupBox4"
        Me.UltraGroupBox4.Size = New System.Drawing.Size(199, 50)
        Me.UltraGroupBox4.TabIndex = 105
        '
        'cmdReset
        '
        Appearance16.Image = CType(resources.GetObject("Appearance16.Image"), Object)
        Me.cmdReset.Appearance = Appearance16
        Me.cmdReset.Location = New System.Drawing.Point(97, 10)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.Size = New System.Drawing.Size(85, 30)
        Me.cmdReset.TabIndex = 5
        Me.cmdReset.Text = "&Reset"
        '
        'cmdSave
        '
        Appearance20.Image = Global.DBLotVbnet.My.Resources.Resources.save_as
        Me.cmdSave.Appearance = Appearance20
        Me.cmdSave.Enabled = False
        Me.cmdSave.Location = New System.Drawing.Point(6, 10)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(85, 30)
        Me.cmdSave.TabIndex = 4
        Me.cmdSave.Text = "&Save"
        '
        'UltraGroupBox3
        '
        Me.UltraGroupBox3.Controls.Add(Me.cmdEdit)
        Me.UltraGroupBox3.Controls.Add(Me.cmdAdd)
        Me.UltraGroupBox3.Location = New System.Drawing.Point(13, 105)
        Me.UltraGroupBox3.Name = "UltraGroupBox3"
        Me.UltraGroupBox3.Size = New System.Drawing.Size(195, 50)
        Me.UltraGroupBox3.TabIndex = 104
        '
        'cmdEdit
        '
        Appearance17.Image = CType(resources.GetObject("Appearance17.Image"), Object)
        Me.cmdEdit.Appearance = Appearance17
        Me.cmdEdit.Enabled = False
        Me.cmdEdit.Location = New System.Drawing.Point(100, 12)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(85, 30)
        Me.cmdEdit.TabIndex = 1
        Me.cmdEdit.Text = "&Edit"
        '
        'cmdAdd
        '
        Appearance18.Image = CType(resources.GetObject("Appearance18.Image"), Object)
        Appearance18.TextHAlignAsString = "Center"
        Me.cmdAdd.Appearance = Appearance18
        Me.cmdAdd.Location = New System.Drawing.Point(6, 11)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(85, 30)
        Me.cmdAdd.TabIndex = 0
        Me.cmdAdd.Text = "&Add"
        '
        'OPR2
        '
        Me.OPR2.Controls.Add(Me.txtTo)
        Me.OPR2.Controls.Add(Me.UltraLabel5)
        Me.OPR2.Controls.Add(Me.txtFrom)
        Me.OPR2.Controls.Add(Me.UltraLabel1)
        Me.OPR2.Controls.Add(Me.UltraLabel4)
        Me.OPR2.Enabled = False
        Me.OPR2.Location = New System.Drawing.Point(12, 10)
        Me.OPR2.Name = "OPR2"
        Me.OPR2.Size = New System.Drawing.Size(593, 82)
        Me.OPR2.TabIndex = 103
        Me.OPR2.Text = " Detailes"
        '
        'txtTo
        '
        Me.txtTo.Location = New System.Drawing.Point(123, 46)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.Size = New System.Drawing.Size(464, 21)
        Me.txtTo.TabIndex = 47
        '
        'UltraLabel5
        '
        Appearance1.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel5.Appearance = Appearance1
        Me.UltraLabel5.Location = New System.Drawing.Point(6, 46)
        Me.UltraLabel5.Name = "UltraLabel5"
        Me.UltraLabel5.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel5.TabIndex = 46
        Me.UltraLabel5.Text = "To Device"
        Me.UltraLabel5.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtFrom
        '
        Me.txtFrom.Location = New System.Drawing.Point(123, 19)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.Size = New System.Drawing.Size(464, 21)
        Me.txtFrom.TabIndex = 45
        '
        'UltraLabel1
        '
        Appearance13.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel1.Appearance = Appearance13
        Me.UltraLabel1.Location = New System.Drawing.Point(6, 21)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel1.TabIndex = 44
        Me.UltraLabel1.Text = "From"
        Me.UltraLabel1.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'UltraLabel4
        '
        Appearance64.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel4.Appearance = Appearance64
        Me.UltraLabel4.Location = New System.Drawing.Point(255, 49)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(92, 21)
        Me.UltraLabel4.TabIndex = 41
        Me.UltraLabel4.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'frmBackup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(660, 194)
        Me.Controls.Add(Me.UltraGroupBox5)
        Me.Controls.Add(Me.UltraGroupBox4)
        Me.Controls.Add(Me.UltraGroupBox3)
        Me.Controls.Add(Me.OPR2)
        Me.Name = "frmBackup"
        Me.Text = "Backup Database"
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox5.ResumeLayout(False)
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox4.ResumeLayout(False)
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox3.ResumeLayout(False)
        CType(Me.OPR2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR2.ResumeLayout(False)
        Me.OPR2.PerformLayout()
        CType(Me.txtTo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtFrom, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraGroupBox5 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdReset As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdSave As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox3 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdEdit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdAdd As Infragistics.Win.Misc.UltraButton
    Friend WithEvents OPR2 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtTo As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel5 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtFrom As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
End Class
