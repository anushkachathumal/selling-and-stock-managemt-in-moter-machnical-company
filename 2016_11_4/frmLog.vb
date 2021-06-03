
Imports System.Data.SqlClient
Imports Infragistics.Win.AppStyling
Imports System.Data
Imports System.Configuration
Imports System
Imports System.Collections
Imports DBLotVbnet.common
Imports DBLotVbnet.DBEngin

Public Class frmLog
    Inherits System.Windows.Forms.Form
    Public returnMessage As String
    Private dsUser As DataSet
    Private m_ChildFormNumber As Integer = 0
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
        ' Call Save_Data()
    End Sub

   

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        common.ClearAll(OPR0)
        OPR0.Enabled = True
    End Sub

    Private Sub txtUname_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUname.KeyUp
        If e.KeyCode = 13 Then
            If txtUname.Text <> "" Then
                With txtPW
                    .Focus()
                    .SelectionStart = 0
                    .SelectionLength = Len(.Text)
                End With
            End If
        End If
    End Sub

    Private Sub txtPW_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPW.KeyUp
        If e.KeyCode = 13 Then
            cmdSave.Focus()
        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        ValidateFields()
        'Me.Close()
        ' frmPayment_Cancel.Show()
    End Sub
    Function ValidateFields() As String

        Dim SQL As String
        Dim con = New SqlConnection()
        Dim recRecords As DataSet
        Dim M01 As DataSet
        Dim M02 As DataSet

        Try





            con = DBEngin.GetConnection()
           

            SQL = "SELECT * FROM X02User_Loging WHERE X02User_Name ='" & txtUname.Text & "' AND X02Password='" & Trim(txtPW.Text) & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If dsUser.Tables(0).Rows.Count = 0 Then
                MsgBox("User name and pasword combination not found", MsgBoxStyle.Information, "INFORMATION .......")
                con.ClearAllPools()

                Exit Function
            Else
                strDisname = Trim(dsUser.Tables(0).Rows(0)("X02User_Name"))
                strUGroup = Trim(dsUser.Tables(0).Rows(0)("X02User_Level"))
                con.ClearAllPools()
                MDIMain.Workstation_Load()
                Me.Close()

                End If


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
            End If

        End Try

    End Function

    Private Sub frmLog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Top = 240
        Me.Left = 330

        'frmPayment_Cancel.cmdAdd.Enabled = False
        'frmPayment_Cancel.cmdReset.Enabled = False

        'frmOutstandingcash.cmdAdd.Enabled = False
        'frmOutstandingcash.cmdReset.Enabled = False

        'frmCancelDistributorAdvanve.cmdAdd.Enabled = False
        'frmCancelDistributorAdvanve.cmdReset.Enabled = False
    End Sub

    Private Sub txtPW_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPW.ValueChanged
        cmdSave.Enabled = True
    End Sub
End Class