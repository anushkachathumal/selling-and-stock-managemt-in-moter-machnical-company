
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmTecoProcess
    Dim Clicked As String
    Dim _Status As String
    'Develop by Suranga R Wijesinghe
    'Developing Date - 2011/04/14
    'Time - 10.30 PM -
    

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Me.Close()
    End Sub


    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR5.Enabled = True
        OPR2.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False
        ' txtVoucher.Focus()
        ' cmdSave.Enabled = True

        cboOrder.ToggleDropdown()
    End Sub

    Private Sub frmTecoProcess_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        chk2.Checked = True
        txtFdate.Text = Today
    End Sub

    Function Clear_Text()
        cboOrder.Text = ""
        txtFdate.Text = Today
       
        txtLine.Text = ""
        txtMaterial.Text = ""
        txtMC.Text = ""
        txtQuality.Text = ""
       
        txtY_Code.Text = ""
        txtY_Type.Text = ""
        cmdSave.Enabled = False
        cmdDelete.Enabled = False
    End Function
    Function Load_OrderNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try

            Sql = "select M03OrderNo as [Order No] from M03Knittingorder where M03Status='T'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboOrder
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 170
            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_UNTechoOrderNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try

            Sql = "select M03OrderNo as [Order No] from M03Knittingorder where M03Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboOrder
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 170
            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub chk2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk2.CheckedChanged
        If chk2.Checked = True Then
            chk1.Checked = False
            _Status = "T"
            Call Load_OrderNo()
            Call Clear_Text()
        Else
            chk1.Checked = True
        End If
    End Sub

    Private Sub chk1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chk1.CheckedChanged
        If chk1.Checked = True Then
            chk2.Checked = False
            _Status = "U"
            Call Clear_Text()
            Call Load_UNTechoOrderNo()
        Else
            chk2.Checked = True
        End If
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR2, OPR5)
        Clicked = ""
        cmdAdd.Enabled = True
        cmdSave.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        cmdAdd.Focus()
        chk2.Checked = True
        chk1.Checked = False
    End Sub

    Function Search_Records() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try

            If _Status = "T" Then
                Sql = "select * from M03Knittingorder where M03OrderNo='" & Trim(cboOrder.Text) & "' and M03Status='T'"
            Else
                Sql = "select * from M03Knittingorder where M03OrderNo='" & Trim(cboOrder.Text) & "' and M03Status='A'"
            End If
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then

                
                If _Status = "T" Then
                    cmdDelete.Enabled = True
                    cmdSave.Enabled = False
                Else
                    cmdSave.Enabled = True
                    cmdDelete.Enabled = False
                End If

                txtQuality.Text = M01.Tables(0).Rows(0)("M03Quality")
                txtMaterial.Text = M01.Tables(0).Rows(0)("M03Material")
                txtMC.Text = M01.Tables(0).Rows(0)("M03MCNo")
                txtY_Type.Text = M01.Tables(0).Rows(0)("M03YarnType")
                txtY_Code.Text = M01.Tables(0).Rows(0)("M03Yarnstock")
                txtLine.Text = M01.Tables(0).Rows(0)("M03CuttingLine")

                Search_Records = True
            Else
                Search_Records = False
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboOrder_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboOrder.InitializeLayout

    End Sub

    Private Sub cboOrder_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboOrder.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Records()
            If cmdDelete.Enabled = True Then
                cmdDelete.Focus()
            Else
                cmdSave.Focus()
            End If
        End If
    End Sub

    Private Sub cboOrder_Layout(ByVal sender As Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles cboOrder.Layout

    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Try
            If Search_Records() = True Then
                nvcFieldList1 = "UPDATE M03Knittingorder SET M03Status='A' WHERE M03OrderNo='" & Trim(cboOrder.Text) & "' AND M03Status='T'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'nvcFieldList1 = "DELETE FROM T09Teco WHERE T09OrderNo='" & Trim(cboOrder.Text) & "' AND T09Status='A'"
                'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into T11TechoProcess(T11OrderNo,T11Date,T11User,T11Status)" & _
                                                         " values(" & Trim(cboOrder.Text) & "," & "convert(varchar(50),getdate(),102)" & ",'" & strDisname & "','U')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                MsgBox("Untecho Process sucessfully completed", MsgBoxStyle.Information, "Texturd Jersy ........")
                transaction.Commit()
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                common.ClearAll(OPR2, OPR5)
                Clicked = ""
                cmdAdd.Enabled = True
                cmdSave.Enabled = False
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
                cmdAdd.Focus()
                chk2.Checked = True
                chk1.Checked = False

                Call Load_UNTechoOrderNo()
                Call Load_OrderNo()
            Else
                MsgBox("Wrong Order No Please try again", MsgBoxStyle.Information, "Textured Jersey ......")
            End If


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Try
            If Search_Records() = True Then
                nvcFieldList1 = "UPDATE M03Knittingorder SET M03Status='T' WHERE M03OrderNo='" & Trim(cboOrder.Text) & "' AND M03Status='A'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'nvcFieldList = "Insert Into T09Teco(T09OrderNo,T09Usable,T09Status,T09Count,T09Workstation,T09Date)" & _
                '                                                   " values('" & Trim(cboOrder.Text) & "','" & C01.Tables(0).Rows(0)("T0Rollweight") & "','A','1','" & netCard & "'," & "convert(varchar(50),getdate(),102)" & ")"
                'ExecuteNonQueryText(connection, transaction, nvcFieldList)


                nvcFieldList1 = "Insert Into T11TechoProcess(T11OrderNo,T11Date,T11User,T11Status)" & _
                                                         " values(" & Trim(cboOrder.Text) & "," & "convert(varchar(50),getdate(),102)" & ",'" & strDisname & "','T')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                MsgBox("Techo Process sucessfully completed", MsgBoxStyle.Information, "Texturd Jersy ........")
                transaction.Commit()
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                common.ClearAll(OPR2, OPR5)
                Clicked = ""
                cmdAdd.Enabled = True
                cmdSave.Enabled = False
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
                cmdAdd.Focus()
                chk2.Checked = True
                chk1.Checked = False
            Else
                MsgBox("Wrong Order No Please try again", MsgBoxStyle.Information, "Textured Jersey ......")
            End If


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub
End Class