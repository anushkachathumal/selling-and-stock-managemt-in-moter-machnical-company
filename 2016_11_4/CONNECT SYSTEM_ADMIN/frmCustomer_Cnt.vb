Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.Drawing.Image
Imports System.IO
Public Class frmCustomer_Cnt
    Dim _PrintStatus As String
    Dim _From As String
    Dim _TO As String
    Dim _Sales_ref As String
    Dim _Root As String

    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        OPR0.Visible = True
        ' Call Load_Sales_Ref()
        Call Load_Entry()
        cboType.ToggleDropdown()
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Function Clear_text()
        Me.txtAddress.Text = ""
        Me.txtCR.Text = ""
        Me.cboType.Text = ""
        '  Me.txtShop_Name.Text = ""
        Me.txtTP.Text = ""
        Me.txtEmail.Text = ""
        Me.txtCus_Name.Text = ""
        Me.cboType.Text = ""
        'Me.cboRoot.Text = ""
        'Me.cboSales_Ref.Text = ""
        'Me.txtMobile.Text = ""
        'PictureBox1.Image = Nothing
        'PictureBox2.Image = Nothing
        Me.txtPic1.Text = ""
        Me.txtPic2.Text = ""
    End Function

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Clear_text()
        OPR1.Visible = False
        cboFrom.Text = ""
        cboTo.Text = ""
        OPR0.Visible = False
        Call Load_Grid()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_text()
        Call Load_Entry()
    End Sub

    Function Load_Type()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M11Name as [##] from M11Common where M11Status='TY'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboType
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 285
                '  .Rows.Band.Columns(1).Width = 160


            End With

            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                con.ClearAllPools()
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    
    Private Sub frmCustomer_Cnt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Type()
        '  Call Load_Sales_Ref()
        Call Load_Entry()
        Call Load_Grid()
        txtCode.ReadOnly = True
        txtCode.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCR.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

    End Sub


    Function Load_Entry()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where  P01Code='CU' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01No") >= 1 And M01.Tables(0).Rows(0)("P01No") < 10 Then
                    txtCode.Text = "CU/00" & M01.Tables(0).Rows(0)("P01No")
                ElseIf M01.Tables(0).Rows(0)("P01No") >= 10 And M01.Tables(0).Rows(0)("P01No") < 100 Then
                    txtCode.Text = "CU/0" & M01.Tables(0).Rows(0)("P01No")
                Else
                    txtCode.Text = "CU/" & M01.Tables(0).Rows(0)("P01No")
                End If
            End If

            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()
            End If
        End Try
    End Function

    

    Private Sub txtCus_Name_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCus_Name.KeyUp
        If e.KeyCode = 13 Then
            If Trim(txtCus_Name.Text) <> "" Then
                txtAddress.Focus()
            End If
        End If
    End Sub

    Private Sub txtAddress_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAddress.KeyUp
        If e.KeyCode = Keys.F1 Then
            txtTP.Focus()
        End If
    End Sub

    Private Sub txtNIC_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = 13 Then
            txtTP.Focus()
        End If
    End Sub

    Private Sub txtTP_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTP.KeyUp
        If e.KeyCode = 13 Then
            txtMobile.Focus()
        End If
    End Sub

    Private Sub txtMobile_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMobile.KeyUp
        If e.KeyCode = 13 Then
            txtCR.Focus()
        End If
    End Sub

    Private Sub txtCR_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCR.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtCR.Text) Then
                Value = txtCR.Text
                txtCR.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtCR.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            txtEmail.Focus()
        End If
    End Sub

    Private Sub txtEmail_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEmail.KeyUp
        If e.KeyCode = 13 Then
            cmdAdd.Focus()
        End If
    End Sub


    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If Trim(cboType.Text) <> "" Then
        Else
            MsgBox("Please select the Customer Type", MsgBoxStyle.Information, "Information ........")
            cboType.ToggleDropdown()
            Exit Sub
        End If

        

        If Trim(txtCus_Name.Text) <> "" Then
        Else
            MsgBox("Please enter the Customer Name", MsgBoxStyle.Information, "Information ........")
            txtCus_Name.Focus()
            Exit Sub
        End If

        If txtAddress.Text <> "" Then
        Else
            txtAddress.Text = "-"
        End If

        If txtTP.Text <> "" Then
        Else
            txtTP.Text = "-"
        End If

        If txtMobile.Text <> "" Then
        Else
            txtMobile.Text = "-"
        End If

        If txtEmail.Text <> "" Then
        Else
            txtEmail.Text = "-"
        End If

      
        If txtCR.Text <> "" Then
        Else
            txtCR.Text = "0"
        End If

        If IsNumeric(txtCR.Text) Then
        Else
            MsgBox("Please enter the correct Credit Limit", MsgBoxStyle.Information, "Information .......")
            txtCR.Focus()
            Exit Sub
        End If
        Call Save_Data()
    End Sub

    Function Save_Data()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        SqlConnection.ClearAllPools()
        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim MB51 As DataSet
        Try

            nvcFieldList1 = "SELECT * FROM M06Customer_Master WHERE M06Code='" & Trim(txtCode.Text) & "' "
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                nvcFieldList1 = "update M06Customer_Master set M06Cus_Type='" & Trim(cboType.Text) & "',M06Address='" & Trim(txtAddress.Text) & "',M06Name='" & Trim(txtCus_Name.Text) & "',M06Contact_No='" & Trim(txtTP.Text) & "',M06Mobile_No='" & Trim(txtMobile.Text) & "',M06Email='" & Trim(txtEmail.Text) & "',M06Credit_Limit='" & CDbl(txtCR.Text) & "',M06Status='A' where M06Code='" & Trim(txtCode.Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                               " values('NEW_CUSTOMER','EDIT', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                nvcFieldList1 = "update P01Parameter set P01No=P01No+ " & 1 & " where P01Code='CU' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into M06Customer_Master(M06Code,M06Name,M06Address,M06Contact_No,M06Mobile_No,M06Email,M06Cus_Type,M06Credit_Limit,M06Status)" & _
                                                                  " values('" & Trim(txtCode.Text) & "','" & Trim(txtCus_Name.Text) & "', '" & Trim(txtAddress.Text) & "','" & Trim(txtTP.Text) & "','" & Trim(txtMobile.Text) & "','" & Trim(txtEmail.Text) & "','" & Trim(cboType.Text) & "','" & CDbl(txtCR.Text) & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                                " values('NEW_CUSTOMER','SAVE', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            End If
            MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()
            connection.ClearAllPools()
            connection.Close()
            '  Call Update_Image()
            'Dim ms As New MemoryStream
            ''  ms.Dispose()
            'PictureBox2.Image.Save(ms, PictureBox2.Image.RawFormat)
            'Dim command As New SqlCommand("UPDATE S01Stock SET S01Img=@Img WHERE  S01id=" & Trim(M01.Tables(0).Rows(0)("S01id")) & " and S01Status='A'  and S01Tr_Type='ISU'", connection)
            'command.Parameters.Add("@img", SqlDbType.Image).Value = ms.ToArray()
            'connection.Open()
            'If command.ExecuteNonQuery() = 1 Then

            '    '  MsgBox("Records update Successfully", MsgBoxStyle.Information, "Information .......")

            'Else

            '    MsgBox("Records update unsuccessfully", MsgBoxStyle.Information, "Information .......")

            'End If
            'connection.Close()
            'ms.Dispose()


            ' Me.txtShop_Name.Text = ""
            Me.txtCus_Name.Text = ""
            Me.txtAddress.Text = ""
            Me.txtMobile.Text = ""
            Me.txtTP.Text = ""
            Me.txtCR.Text = ""
            Me.txtEmail.Text = ""
            Me.cboType.Text = ""
            cboType.ToggleDropdown()
            Call Load_Grid()
            Call Load_Entry()
            frmJob_Card_Uniq.Load_Customer_name()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.ClearAllPools()
                connection.Close()
            End If
        End Try
    End Function
    
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Call Clear_text()
        Call Load_Entry()
        OPR0.Visible = False
    End Sub


    Function Load_Grid_Root_Change()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  M02Root_Name as [Root Name],M05Cus_Code as [Customer Code],M05Shop_Name as [Shop Name]  from M05New_Customer inner join M02New_Root on M05Root_Code=M02Root_Code inner join M04Create_Sale_Ref on M04Rer_No=M05Ref_Code where M05Status='A' and M04Ref_Name='" & cboFrom.Text & "' order by M05ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 160
            UltraGrid2.Rows.Band.Columns(1).Width = 110
            UltraGrid2.Rows.Band.Columns(2).Width = 260

            ' UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            con.ClearAllPools()
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  M06Cus_Type as [##],M06Code as [Customer Code],M06Name as [Customer Name],M06Contact_No as [Contact No]  from M06Customer_Master where M06Status='A' order by M06ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            UltraGrid1.Rows.Band.Columns(0).Width = 110
            UltraGrid1.Rows.Band.Columns(1).Width = 110
            UltraGrid1.Rows.Band.Columns(2).Width = 310
            UltraGrid1.Rows.Band.Columns(3).Width = 90
            'UltraGrid1.Rows.Band.Columns(4).Width = 260
            'UltraGrid1.Rows.Band.Columns(5).Width = 120
            ' UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            con.ClearAllPools()
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_deactive()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  M06Cus_Type as [##],M06Code as [Customer Code],M06Name as [Customer Name],M06Contact_No as [Contact No]  from M06Customer_Master where M06Status='I' order by M06ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            UltraGrid1.Rows.Band.Columns(0).Width = 110
            UltraGrid1.Rows.Band.Columns(1).Width = 110
            UltraGrid1.Rows.Band.Columns(2).Width = 310
            UltraGrid1.Rows.Band.Columns(3).Width = 90
            ' UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            con.ClearAllPools()
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function


    Private Sub UltraGrid1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid1.DoubleClickRow
        On Error Resume Next
        Dim _Row As Integer

        _Row = UltraGrid1.ActiveRow.Index
        OPR0.Visible = True
        txtCode.Text = Trim(UltraGrid1.Rows(_Row).Cells(1).Text)
        Call Search_Records()
        cboType.ToggleDropdown()
    End Sub

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double

        Try
            Sql = "select  *  from M06Customer_Master where M06Code='" & Trim(txtCode.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboType.Text = Trim(M01.Tables(0).Rows(0)("M06Cus_Type"))
               
                txtCus_Name.Text = Trim(M01.Tables(0).Rows(0)("M06Name"))
                txtAddress.Text = Trim(M01.Tables(0).Rows(0)("M06Address"))
                txtTP.Text = Trim(M01.Tables(0).Rows(0)("M06Contact_No"))
                txtMobile.Text = Trim(M01.Tables(0).Rows(0)("M06Mobile_No"))
                txtEmail.Text = Trim(M01.Tables(0).Rows(0)("M06Email"))
                Value = Trim(M01.Tables(0).Rows(0)("M06Credit_Limit"))
                txtCR.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtCR.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                
            End If

            con.ClearAllPools()
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()
            End If
        End Try
    End Function

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Dim A As String

        A = MsgBox("Are you sure you want to deactive this customer", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Deactive Customer .......")
        If A = vbYes Then
            Call Deactive_Customer()
        End If
    End Sub

    Function Deactive_Customer()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        SqlConnection.ClearAllPools()
        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim MB51 As DataSet
        Try

            nvcFieldList1 = "SELECT * FROM M06Customer_Master WHERE M06Code='" & Trim(txtCode.Text) & "' "
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                nvcFieldList1 = "update M06Customer_Master set M06Status='I' where M06Code='" & Trim(txtCode.Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                               " values('NEW_CUSTOMER','DELETE', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


            End If
            MsgBox("Record deactivate successfully", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()
            connection.ClearAllPools()
            connection.Close()

            Me.cboType.Text = ""
            Me.txtCus_Name.Text = ""
            Me.txtAddress.Text = ""
            Me.txtMobile.Text = ""
            Me.txtTP.Text = ""
            Me.txtCR.Text = ""
            Me.cboType.Text = ""
           
            Me.txtEmail.Text = ""
            cboType.ToggleDropdown()
            Call Load_Grid()
            Call Load_Entry()
            frmJob_Card_Uniq.Load_Customer_name()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.ClearAllPools()
                connection.Close()
            End If
        End Try
    End Function

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        OPR1.Visible = False
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        OPR1.Visible = True
        cboFrom.Text = ""
        cboTo.Text = ""
        Call Load_Grid_Root_Change()
        cboFrom.ToggleDropdown()
    End Sub

    Private Sub cboFrom_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFrom.AfterCloseUp
        Call Load_Grid_Root_Change()
    End Sub

    Private Sub cboFrom_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFrom.TextChanged
        Call Load_Grid_Root_Change()
    End Sub

    Private Sub DeactiveCustomersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeactiveCustomersToolStripMenuItem.Click
        Call Load_Grid_deactive()
    End Sub

    Private Sub cboType_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboType.KeyUp
        If e.KeyCode = 13 Then
            txtCus_Name.Focus()
        End If
    End Sub

 
End Class