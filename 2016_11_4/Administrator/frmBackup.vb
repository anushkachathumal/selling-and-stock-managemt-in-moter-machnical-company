Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System
Imports System.IO
Public Class frmBackup
    Dim Clicked As String
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Try
            My.Computer.FileSystem.CreateDirectory(ConfigurationManager.AppSettings("Backup") & netCard & CStr(Today.Day) + CStr(Month(Today)) + CStr(Year(Today)))
            Dim di As DirectoryInfo = New DirectoryInfo(ConfigurationManager.AppSettings("Backup") & netCard & CStr(Today.Day) + CStr(Month(Today)) + CStr(Year(Today)))
            If di.Exists = True Then
                ' Indicate that it already exists.
                '  MsgBox("That path exists already.")
            Else
                ' Try to create the directory.
                di.Create()
            End If

            'con = New SqlConnection("Data Source=.\SQLEXPRESS;AttachDbFilename=|DataDirectory|\ClientData.mdf;Integrated Security=True;User Instance=True")
            con = New SqlConnection(ConfigurationManager.AppSettings("CD"))
            'cmd = New SqlCommand("BACKUP DATABASE SingDBStock TO DISK = 'd:DBStock.bak'", con)
            cmd = New SqlCommand("BACKUP DATABASE DBSparepart TO DISK=" & "'" & txtTo.Text & "'", con)
            con.Open()

            cmd.ExecuteNonQuery()

            con.Close()
            MsgBox("System Backup successfully" & txtTo.Text, MsgBoxStyle.Information, "Sign Info .....")

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        '  OPR0.Enabled = True
        ' OPR1.Enabled = True
        OPR2.Enabled = True
        'OPR3.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False
        txtFrom.Text = ConfigurationManager.AppSettings("CD")
        txtTo.Text = ConfigurationManager.AppSettings("Backup") + CStr(netCard) + CStr(Today.Day) + CStr(Month(Today)) + CStr(Year(Today)) + "\DBStock.bak"
        cmdSave.Enabled = True
    End Sub

    Private Sub frmBackup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFrom.ReadOnly = True
        txtTo.ReadOnly = True
       
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR2)
        Clicked = ""
        cmdAdd.Enabled = True
        cmdSave.Enabled = False
        cmdEdit.Enabled = False
        'cmdDelete.Enabled = False
        cmdAdd.Focus()
    End Sub
End Class