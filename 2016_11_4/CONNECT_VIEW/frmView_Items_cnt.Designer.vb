<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmView_Items_cnt
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
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraGrid2 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.OPR0 = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraLabel8 = New Infragistics.Win.Misc.UltraLabel
        Me.txtSearch = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        CType(Me.UltraGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OPR0.SuspendLayout()
        CType(Me.txtSearch, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UltraGrid2
        '
        Me.UltraGrid2.Location = New System.Drawing.Point(15, 70)
        Me.UltraGrid2.Name = "UltraGrid2"
        Me.UltraGrid2.Size = New System.Drawing.Size(798, 386)
        Me.UltraGrid2.TabIndex = 120
        '
        'OPR0
        '
        Me.OPR0.Controls.Add(Me.UltraLabel8)
        Me.OPR0.Controls.Add(Me.txtSearch)
        Me.OPR0.Location = New System.Drawing.Point(15, 12)
        Me.OPR0.Name = "OPR0"
        Me.OPR0.Size = New System.Drawing.Size(798, 52)
        Me.OPR0.TabIndex = 119
        '
        'UltraLabel8
        '
        Appearance7.BackColorAlpha = Infragistics.Win.Alpha.Transparent
        Me.UltraLabel8.Appearance = Appearance7
        Me.UltraLabel8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel8.Location = New System.Drawing.Point(6, 16)
        Me.UltraLabel8.Name = "UltraLabel8"
        Me.UltraLabel8.Size = New System.Drawing.Size(57, 21)
        Me.UltraLabel8.TabIndex = 48
        Me.UltraLabel8.Text = "#Find"
        Me.UltraLabel8.UseOsThemes = Infragistics.Win.DefaultableBoolean.[True]
        '
        'txtSearch
        '
        Appearance8.ForeColor = System.Drawing.Color.Black
        Me.txtSearch.Appearance = Appearance8
        Me.txtSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSearch.Location = New System.Drawing.Point(69, 14)
        Me.txtSearch.MaxLength = 200
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(723, 23)
        Me.txtSearch.TabIndex = 1
        '
        'frmView_Items_cnt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(825, 469)
        Me.Controls.Add(Me.UltraGrid2)
        Me.Controls.Add(Me.OPR0)
        Me.MaximizeBox = False
        Me.Name = "frmView_Items_cnt"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "#View Items"
        CType(Me.UltraGrid2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.OPR0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OPR0.ResumeLayout(False)
        Me.OPR0.PerformLayout()
        CType(Me.txtSearch, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraGrid2 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents OPR0 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraLabel8 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtSearch As Infragistics.Win.UltraWinEditors.UltraTextEditor
End Class
