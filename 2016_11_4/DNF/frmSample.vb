Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Public Class frmSample
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim _EralCode As Integer
    Dim _MCNo As String

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        OPR2.Enabled = True
        OPR1.Enabled = True
     
        txtDate.Text = Today

        ' Call Clear_Text()
        cmdAdd.Enabled = False
        'cboBatch.ToggleDropdown()
        'Call Search_RefNo()

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

    End Sub

    Private Sub frmSample_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtArelWg.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtEralBatchwg.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtSDT.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtDate.Text = Today
        txtLoadDate.Text = Today
        txtUnload_D.Text = Today


    End Sub
End Class