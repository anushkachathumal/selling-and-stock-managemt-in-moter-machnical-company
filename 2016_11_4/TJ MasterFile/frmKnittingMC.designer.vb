<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmKnittingMC
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmKnittingMC))
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.OPR4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraGrid1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.UltraGroupBox5 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdExit = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox4 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdReset = New Infragistics.Win.Misc.UltraButton
        Me.cmdSave = New Infragistics.Win.Misc.UltraButton
        Me.UltraGroupBox3 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmdDelete = New Infragistics.Win.Misc.UltraButton
        Me.cmdEdit = New Infragistics.Win.Misc.UltraButton
        Me.cmdAdd = New Infragistics.Win.Misc.UltraButton
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtVoucher = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel
        Me.txtDis = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        CType(Me.OPR4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR4.SuspendLayout()
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox5.SuspendLayout()
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox4.SuspendLayout()
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox3.SuspendLayout()
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.txtVoucher, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDis, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OPR4
        '
        Me.OPR4.Controls.Add(Me.UltraGrid1)
        Me.OPR4.Location = New System.Drawing.Point(10, 106)
        Me.OPR4.Name = "OPR4"
        Me.OPR4.Size = New System.Drawing.Size(641, 174)
        Me.OPR4.TabIndex = 130
        Me.OPR4.Text = "Downtime Details"
        '
        'UltraGrid1
        '
        Me.UltraGrid1.Location = New System.Drawing.Point(6, 19)
        Me.UltraGrid1.Name = "UltraGrid1"
        Me.UltraGrid1.Size = New System.Drawing.Size(629, 139)
        Me.UltraGrid1.TabIndex = 1
        '
        'UltraGroupBox5
        '
        Me.UltraGroupBox5.Controls.Add(Me.cmdExit)
        Me.UltraGroupBox5.Location = New System.Drawing.Point(506, 285)
        Me.UltraGroupBox5.Name = "UltraGroupBox5"
        Me.UltraGroupBox5.Size = New System.Drawing.Size(145, 50)
        Me.UltraGroupBox5.TabIndex = 129
        '
        'cmdExit
        '
        Appearance21.Image = CType(resources.GetObject("Appearance21.Image"), Object)
        Me.cmdExit.Appearance = Appearance21
        Me.cmdExit.Location = New System.Drawing.Point(6, 10)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(131, 30)
        Me.cmdExit.TabIndex = 3
        Me.cmdExit.Text = "&Exit"
        '
        'UltraGroupBox4
        '
        Me.UltraGroupBox4.Controls.Add(Me.cmdReset)
        Me.UltraGroupBox4.Controls.Add(Me.cmdSave)
        Me.UltraGroupBox4.Location = New System.Drawing.Point(309, 286)
        Me.UltraGroupBox4.Name = "UltraGroupBox4"
        Me.UltraGroupBox4.Size = New System.Drawing.Size(192, 50)
        Me.UltraGroupBox4.TabIndex = 128
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
        Me.UltraGroupBox3.Controls.Add(Me.cmdDelete)
        Me.UltraGroupBox3.Controls.Add(Me.cmdEdit)
        Me.UltraGroupBox3.Controls.Add(Me.cmdAdd)
        Me.UltraGroupBox3.Location = New System.Drawing.Point(11, 285)
        Me.UltraGroupBox3.Name = "UltraGroupBox3"
        Me.UltraGroupBox3.Size = New System.Drawing.Size(290, 50)
        Me.UltraGroupBox3.TabIndex = 127
        '
        'cmdDelete
        '
        Appearance19.Image = CType(resources.GetObject("Appearance19.Image"), Object)
        Me.cmdDelete.Appearance = Appearance19
        Me.cmdDelete.Enabled = False
        Me.cmdDelete.Location = New System.Drawing.Point(194, 12)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(85, 30)
        Me.cmdDelete.TabIndex = 2
        Me.cmdDelete.Text = "&Delete"
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
        Me.cmdAdd.Text = "&New Entry"
        '
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.txtVoucher)
        Me.OPR0.Controls.Add(Me.UltraLabel4)
        Me.OPR0.Controls.Add(Me.txtDis)
        Me.OPR0.Controls.Add(Me.UltraLabel2)
        Me.OPR0.Enabled = False
        Me.OPR0.Location = New System.Drawing.Point(11, 10)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(640, 92)
        Me.OPR0.TabIndex = 126
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
        Me.UltraLabel2.Text = "Ref.No"
        Me.UltraLabel2.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'frmKnittingMC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(684, 401)
        Me.Controls.Add(Me.OPR4)
        Me.Controls.Add(Me.UltraGroupBox5)
        Me.Controls.Add(Me.UltraGroupBox4)
        Me.Controls.Add(Me.UltraGroupBox3)
        Me.Controls.Add(Me.OPR0)
        Me.Name = "frmKnittingMC"
        Me.Text = "Setup New Down Time"
        CType(Me.OPR4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR4.ResumeLayout(False)
        CType(Me.UltraGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox5.ResumeLayout(False)
        CType(Me.UltraGroupBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox4.ResumeLayout(False)
        CType(Me.UltraGroupBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox3.ResumeLayout(False)
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.txtVoucher, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDis, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents OPR4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraGrid1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents UltraGroupBox5 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdExit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox4 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdReset As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdSave As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraGroupBox3 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents cmdDelete As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdEdit As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdAdd As Infragistics.Win.Misc.UltraButton
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtVoucher As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtDis As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
End Class
