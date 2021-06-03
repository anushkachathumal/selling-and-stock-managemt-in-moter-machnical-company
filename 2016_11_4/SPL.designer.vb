<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SPL
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SPL))
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraListView1 = New Infragistics.Win.UltraWinListView.UltraListView
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraGroupBox2 = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraLabel3 = New Infragistics.Win.Misc.UltraLabel
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox1.SuspendLayout()
        CType(Me.UltraListView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'UltraGroupBox1
        '
        Appearance1.AlphaLevel = CType(88, Short)
        Appearance1.BackColor = System.Drawing.Color.White
        Me.UltraGroupBox1.Appearance = Appearance1
        Me.UltraGroupBox1.Controls.Add(Me.UltraLabel2)
        Me.UltraGroupBox1.Controls.Add(Me.UltraListView1)
        Me.UltraGroupBox1.Dock = System.Windows.Forms.DockStyle.Right
        Me.UltraGroupBox1.Location = New System.Drawing.Point(237, 0)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(241, 296)
        Me.UltraGroupBox1.TabIndex = 3
        Me.UltraGroupBox1.Text = "Loading Status"
        '
        'UltraLabel2
        '
        Me.UltraLabel2.Location = New System.Drawing.Point(0, 19)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(241, 62)
        Me.UltraLabel2.TabIndex = 6
        Me.UltraLabel2.Text = "Technova Solutions.                                            No 59/5 Miriswatta" & _
            ",                                                                   Avissawella," & _
            "Sri Lanka"
        '
        'UltraListView1
        '
        Appearance14.AlphaLevel = CType(63, Short)
        Me.UltraListView1.Appearance = Appearance14
        Me.UltraListView1.BorderStyle = Infragistics.Win.UIElementBorderStyle.Rounded1Etched
        Me.UltraListView1.Enabled = False
        Appearance3.AlphaLevel = CType(199, Short)
        Me.UltraListView1.ItemSettings.ActiveAppearance = Appearance3
        Appearance4.AlphaLevel = CType(255, Short)
        Me.UltraListView1.ItemSettings.Appearance = Appearance4
        Me.UltraListView1.Location = New System.Drawing.Point(0, 141)
        Me.UltraListView1.Name = "UltraListView1"
        Appearance2.BackColorAlpha = Infragistics.Win.Alpha.UseAlphaLevel
        Appearance2.BorderAlpha = Infragistics.Win.Alpha.UseAlphaLevel
        Me.UltraListView1.ScrollBarLook.Appearance = Appearance2
        Me.UltraListView1.ShowGroups = False
        Me.UltraListView1.Size = New System.Drawing.Size(235, 211)
        Me.UltraListView1.TabIndex = 4
        Me.UltraListView1.TabStop = False
        Me.UltraListView1.View = Infragistics.Win.UltraWinListView.UltraListViewStyle.Details
        Appearance5.BorderAlpha = Infragistics.Win.Alpha.UseAlphaLevel
        Me.UltraListView1.ViewSettingsDetails.ColumnHeaderAppearance = Appearance5
        Me.UltraListView1.ViewSettingsDetails.ColumnHeaderBorderStyle = Infragistics.Win.UIElementBorderStyle.None
        Me.UltraListView1.ViewSettingsDetails.ColumnHeaderImageSize = New System.Drawing.Size(0, 0)
        Me.UltraListView1.ViewSettingsDetails.ColumnsShowSortIndicators = False
        Me.UltraListView1.ViewSettingsDetails.ImageSize = New System.Drawing.Size(0, 0)
        Me.UltraListView1.ViewSettingsDetails.SubItemTipStyle = Infragistics.Win.UltraWinListView.SubItemTipStyle.Hide
        '
        'Timer1
        '
        Me.Timer1.Interval = 5000
        '
        'UltraLabel1
        '
        Appearance6.AlphaLevel = CType(130, Short)
        Me.UltraLabel1.Appearance = Appearance6
        Me.UltraLabel1.Font = New System.Drawing.Font("Arial Black", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel1.Location = New System.Drawing.Point(12, 12)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(219, 127)
        Me.UltraLabel1.TabIndex = 4
        Me.UltraLabel1.Text = "Workshop Management System"
        '
        'UltraGroupBox2
        '
        Appearance7.AlphaLevel = CType(125, Short)
        Appearance7.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.UltraGroupBox2.Appearance = Appearance7
        Me.UltraGroupBox2.BorderStyle = Infragistics.Win.Misc.GroupBoxBorderStyle.Rounded
        Me.UltraGroupBox2.Controls.Add(Me.UltraLabel3)
        Me.UltraGroupBox2.Location = New System.Drawing.Point(-4, 236)
        Me.UltraGroupBox2.Name = "UltraGroupBox2"
        Me.UltraGroupBox2.Size = New System.Drawing.Size(254, 45)
        Me.UltraGroupBox2.TabIndex = 5
        '
        'UltraLabel3
        '
        Me.UltraLabel3.Location = New System.Drawing.Point(16, 10)
        Me.UltraLabel3.Name = "UltraLabel3"
        Me.UltraLabel3.Size = New System.Drawing.Size(234, 30)
        Me.UltraLabel3.TabIndex = 0
        Me.UltraLabel3.Text = "Copyright 2009 @ Technova Solution.          All Rights Reserved. Ver:01.00"
        '
        'SPL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(478, 296)
        Me.Controls.Add(Me.UltraGroupBox2)
        Me.Controls.Add(Me.UltraLabel1)
        Me.Controls.Add(Me.UltraGroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Location = New System.Drawing.Point(75, 50)
        Me.Name = "SPL"
        Me.Opacity = 0.8
        Me.ShowInTaskbar = False
        Me.Text = "Form2"
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox1.ResumeLayout(False)
        CType(Me.UltraListView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents UltraListView1 As Infragistics.Win.UltraWinListView.UltraListView
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraGroupBox2 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraLabel3 As Infragistics.Win.Misc.UltraLabel
End Class
