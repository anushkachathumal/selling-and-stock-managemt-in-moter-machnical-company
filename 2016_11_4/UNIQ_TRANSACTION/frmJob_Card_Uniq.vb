Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Imports System.IO
Public Class frmJob_Card_Uniq
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _Supplier As String
    Dim _Location As Integer
    Dim _LogStaus As Boolean
    Dim _UserLevel As String
    Dim _CusNo As String

    Private Sub frmJob_Card_Uniq_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmCustomer_Cnt.Close()
        frmvew_Job.Close()
        frmView_Customer.Close()
        frmView_Vehicle_History.Close()
    End Sub

    Private Sub frmJob_Card_Uniq_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim A As String

        Call Load_EntryNo()
        txtEntry.ReadOnly = True
        txtEntry.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDate.Text = Today
        Call Load_DEPARTMENT()
        Call Load_Brand()
        Call Load_Type()
        Call Load_VNO()
        txtMtr.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_CUS_TYPE()
        Call Load_Customer_name()
        A = ConfigurationManager.AppSettings("IMAGE") + "\images.jpg"
        PictureBox1.Image = Image.FromFile(A)
        txtPic1.Text = A
    End Sub

    Function Load_Brand()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select M02Description as [##] from M02Barnd_Name WHERE M02Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboBrand
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 159
                '  .Rows.Band.Columns(1).Width = 160


            End With

            con.ClearAllPools()
            con.CLOSE()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Function Load_CUS_TYPE()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select  upper(M11Name) as [##] from M11Common WHERE M11Status='TY' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCus_Type
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 124
                '  .Rows.Band.Columns(1).Width = 160


            End With

            con.ClearAllPools()
            con.CLOSE()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Function Load_Customer_name()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select  (M06Name) as [##] from M06Customer_Master WHERE M06Status='A' order by M06ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCus_Name
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 384
                '  .Rows.Band.Columns(1).Width = 160


            End With

            con.ClearAllPools()
            con.CLOSE()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function


    Function Load_Type()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select M03Description as [##] from M03Vehicle_Type  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cbov_Type
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 148
                '  .Rows.Band.Columns(1).Width = 160


            End With

            con.ClearAllPools()
            con.CLOSE()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Function Load_DEPARTMENT()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M08Description as [##] from M08Department where M08Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboDepartment
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 212
                ' .Rows.Band.Columns(1).Width = 360

            End With
            con.ClearAllPools()
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Function Load_VNO()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M07V_No as [##],M06Name as [Customer Name] from M07Vehicle_Master inner join M06Customer_Master on M06Code=M07Cus_Code where M07Status='A'  order by M07ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboV_no
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 159
                .Rows.Band.Columns(1).Width = 210

            End With
            con.ClearAllPools()
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Function SEARCH_RECORDS()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select * from T05Job_Card  where T05Status='A' AND T05Job_No='" & Trim(txtEntry.Text) & "' order by T05Id"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtRef_no.Text = Trim(M01.Tables(0).Rows(0)("T05Ref_No"))
                txtDate.Text = Trim(M01.Tables(0).Rows(0)("T05Date"))
                cboDepartment.Text = Trim(M01.Tables(0).Rows(0)("T05Department"))
                cboV_no.Text = Trim(M01.Tables(0).Rows(0)("T05Vehi_No"))
                Call Search_Vehicle_No()
                txtRemark.Text = Trim(M01.Tables(0).Rows(0)("T05Remark"))
                txtMtr.Text = Trim(M01.Tables(0).Rows(0)("T05Mtr"))
                cmdEdit.Enabled = True
                cmdDelete.Enabled = True
                Dim arrayImage() As Byte = CType(M01.Tables(0).Rows(0)("T05Img"), Byte())
                Dim ms As New MemoryStream(arrayImage)
                PictureBox1.Image = Image.FromStream(ms)
            End If
            con.ClearAllPools()
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function
    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Load_EntryNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where  P01Code='JB'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01No") >= 1 And M01.Tables(0).Rows(0)("P01No") < 10 Then
                    txtEntry.Text = "JOB-00" & M01.Tables(0).Rows(0)("P01No")
                ElseIf M01.Tables(0).Rows(0)("P01No") >= 10 And M01.Tables(0).Rows(0)("P01No") < 100 Then
                    txtEntry.Text = "JOB-0" & M01.Tables(0).Rows(0)("P01No")
                Else
                    txtEntry.Text = "JOB-" & M01.Tables(0).Rows(0)("P01No")
                End If
            End If

            con.ClearAllPools()
            con.CLOSE()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Private Sub txtRef_no_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRef_no.KeyUp
        If e.KeyCode = 13 Then
            cboDepartment.ToggleDropdown()
        End If
    End Sub

    Private Sub cboDepartment_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboDepartment.KeyUp
        If e.KeyCode = 13 Then
            cboV_no.ToggleDropdown()
        End If
    End Sub

    
    Function Search_Vehicle_No() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Search_Vehicle_No = False
            Sql = "select * from M07Vehicle_Master inner join M06Customer_Master on M06Code=M07Cus_Code where M07Status='A'  and M07V_No='" & Trim(cboV_no.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Vehicle_No = True
                _CusNo = Trim(M01.Tables(0).Rows(0)("M06Code"))
                cboBrand.Text = Trim(M01.Tables(0).Rows(0)("M07Brand_Name"))
                cbov_Type.Text = Trim(M01.Tables(0).Rows(0)("M07Type"))
                txtTp.Text = Trim(M01.Tables(0).Rows(0)("M06Mobile_No"))
                cboCus_Name.Text = Trim(M01.Tables(0).Rows(0)("M06Name"))
                txtAddress.Text = Trim(M01.Tables(0).Rows(0)("M06Address"))
                cboCus_Type.Text = Trim(M01.Tables(0).Rows(0)("M06Cus_Type"))
            End If
            con.ClearAllPools()
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Private Sub cboV_no_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboV_no.AfterCloseUp
        Call Search_Vehicle_No()
    End Sub

    Private Sub cboV_no_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboV_no.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Vehicle_No()
            txtMtr.Focus()
        End If
    End Sub

    Private Sub cboBrand_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboBrand.KeyUp
        If e.KeyCode = 13 Then
            cbov_Type.ToggleDropdown()
        End If
    End Sub

    Private Sub txtMtr_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMtr.KeyUp
        If e.KeyCode = 13 Then
            cboBrand.ToggleDropdown()
        End If
    End Sub

    Private Sub cbov_Type_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cbov_Type.KeyUp
        If e.KeyCode = 13 Then
            cboCus_Name.ToggleDropdown()
        End If
    End Sub

    Function search_Customer() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select * from M06Customer_Master  where M06Status='A'  and M06Name='" & Trim(cboCus_Name.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _CusNo = Trim(M01.Tables(0).Rows(0)("M06Code"))
                txtTp.Text = Trim(M01.Tables(0).Rows(0)("M06Mobile_No"))
                txtAddress.Text = Trim(M01.Tables(0).Rows(0)("M06Address"))
                cboCus_Type.Text = Trim(M01.Tables(0).Rows(0)("M06Cus_Type"))
                search_Customer = True
                Exit Function
            Else
               
            End If
            con.ClearAllPools()
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Private Sub txtTp_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTp.KeyUp
        If e.KeyCode = 13 Then
            If search_Customer() = True Then
                txtRemark.Focus()
            Else
                cboCus_Type.ToggleDropdown()
            End If
        End If
    End Sub

    Private Sub txtTp_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTp.ValueChanged
        'Call search_Customer()
    End Sub

    Private Sub cboCus_Type_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCus_Type.KeyUp
        If e.KeyCode = 13 Then
            txtRemark.Focus()
        End If
    End Sub

    Private Sub cboCus_Name_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCus_Name.AfterCloseUp
        Call search_Customer()
    End Sub

    Private Sub cboCus_Name_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCus_Name.KeyUp
        If e.KeyCode = 13 Then
            txtAddress.Focus()
        End If
    End Sub

    Private Sub txtAddress_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAddress.KeyUp
        If e.KeyCode = Keys.Escape Then
            txtTp.Focus()
        End If
    End Sub

    Private Sub txtRemark_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemark.KeyUp
        If e.KeyCode = Keys.Escape Then
            cmdAdd.Focus()
        End If
    End Sub

    Private Sub cmdpic_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdpic_1.Click
        On Error Resume Next
        OpenFileDialog1.Filter = "Image Files|*.jpg;*.gif;*.png;*.bmp"
        OpenFileDialog1.ShowDialog()
        PictureBox1.Image = Image.FromFile(OpenFileDialog1.FileName)
        txtPic1.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If Trim(cboDepartment.Text) <> "" Then
        Else
            MsgBox("Please select the department", MsgBoxStyle.Information, "Information .........")
            cboDepartment.ToggleDropdown()
            Exit Sub
        End If

        If txtMtr.Text <> "" Then
        Else
            txtMtr.Text = "0"
        End If

        If IsNumeric(txtMtr.Text) Then
        Else
            MsgBox("Please enter the meter reading", MsgBoxStyle.Information, "Information .........")
            txtMtr.Focus()
            Exit Sub
        End If

        If Trim(cboV_no.Text) <> "" Then
        Else
            MsgBox("Please enter the Vehicle No", MsgBoxStyle.Information, "Information ...........")
            cboV_no.ToggleDropdown()
            Exit Sub
        End If

        If Trim(txtTp.Text) <> "" Then
        Else
            txtTp.Text = "-"
        End If

        If Trim(cboCus_Type.Text) <> "" Then
        Else
            MsgBox("Please select the customer type", MsgBoxStyle.Information, "Information .........")
            cboCus_Type.ToggleDropdown()
            Exit Sub
        End If

        If Trim(cboCus_Name.Text) <> "" Then
        Else
            MsgBox("Please select the customer name", MsgBoxStyle.Information, "Information ..........")
            cboCus_Name.ToggleDropdown()
            Exit Sub
        End If

        If txtAddress.Text <> "" Then
        Else
            txtAddress.Text = "-"
        End If

        If Trim(txtRemark.Text) <> "" Then
        Else
            txtRemark.Text = "-"
        End If

        If txtRef_no.Text <> "" Then
        Else
            txtRef_no.Text = "-"
        End If

        If Trim(cboBrand.Text) <> "" Then
        Else
            MsgBox("Please enter the Brand Name", MsgBoxStyle.Information, "Information ...........")
            cboBrand.ToggleDropdown()
            Exit Sub
        End If

        If Trim(cbov_Type.Text) <> "" Then
        Else
            MsgBox("Please enter the Vehicle Type", MsgBoxStyle.Information, "Information ...........")
            cboBrand.ToggleDropdown()
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
        SqlClient.SqlConnection.ClearAllPools()
        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim MB51 As DataSet
        Dim _GetDate As DateTime
        Dim _Get_Time As DateTime
        Dim A As String
        Dim B As New ReportDocument

        Dim M01 As DataSet
        Dim M02 As DataSet
        Try
            nvcFieldList1 = "SELECT * FROM T05Job_Card WHERE T05Job_No='" & Trim(txtEntry.Text) & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                MsgBox("This Job No alrady exsist", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                Exit Function

            End If
            If Search_Vehicle_No() = True Then

            Else
                If search_Customer() = True Then

                Else
                    nvcFieldList1 = "select * from P01Parameter where  P01Code='CU' "
                    M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M01) Then
                        If M01.Tables(0).Rows(0)("P01No") >= 1 And M01.Tables(0).Rows(0)("P01No") < 10 Then
                            _CusNo = "CU/00" & M01.Tables(0).Rows(0)("P01No")
                        ElseIf M01.Tables(0).Rows(0)("P01No") >= 10 And M01.Tables(0).Rows(0)("P01No") < 100 Then
                            _CusNo = "CU/0" & M01.Tables(0).Rows(0)("P01No")
                        Else
                            _CusNo = "CU/" & M01.Tables(0).Rows(0)("P01No")
                        End If
                    End If
                    '=============================================================
                    nvcFieldList1 = "update P01Parameter set P01No=P01No+ " & 1 & " where P01Code='CU' "
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    'SAVE CUSTOMER
                    nvcFieldList1 = "Insert Into M06Customer_Master(M06Code,M06Name,M06Address,M06Contact_No,M06Mobile_No,M06Email,M06Cus_Type,M06Credit_Limit,M06Status)" & _
                                                             " values('" & _CusNo & "','" & UCase(Trim(cboCus_Name.Text)) & "', '" & Trim(txtAddress.Text) & "','-','" & Trim(txtTp.Text) & "','-','" & Trim(cboCus_Type.Text) & "','0','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                '==================================================================
                'SAVE VEHICLE DETAILES
                nvcFieldList1 = "Insert Into M07Vehicle_Master(M07V_No,M07Cus_Code,M07Type,M07Brand_Name,M07Status)" & _
                                                         " values('" & UCase(Trim(cboV_no.Text)) & "','" & _CusNo & "', '" & Trim(cbov_Type.Text) & "','" & Trim(cboBrand.Text) & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            End If
            '==========================================================================
            Call Load_EntryNo()
            nvcFieldList1 = "update P01Parameter set P01No=P01No+ " & 1 & " where P01Code='JB' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            'SAVE T05Job_Card
            _GetDate = Month(txtDate.Text) & "/" & Microsoft.VisualBasic.Day(txtDate.Text) & "/" & Year(txtDate.Text)

            _Get_Time = Month(txtDate.Text) & "/" & Microsoft.VisualBasic.Day(txtDate.Text) & "/" & Year(txtDate.Text) & " " & Hour(Now) & ":" & Minute(Now)

            nvcFieldList1 = "Insert Into T05Job_Card(T05Job_No,T05Ref_No,T05Date,T05Time,T05Department,T05Vehi_No,T05Cus_No,T05Mtr,T05Remark,T05Status,T05Inv_No)" & _
                                                     " values('" & Trim(txtEntry.Text) & "','" & Trim(txtRef_no.Text) & "', '" & _GetDate & "','" & _Get_Time & "','" & Trim(cboDepartment.Text) & "','" & UCase(Trim(cboV_no.Text)) & "','" & _CusNo & "','" & txtMtr.Text & "','" & Trim(txtRemark.Text) & "','A','-')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "Insert Into tmpTransaction_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                      " values('JOB CARD','SAVE', '" & _Get_Time & "','" & strDisname & "','" & Trim(txtEntry.Text) & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            ' Call Update_Image()
            transaction.Commit()
            connection.ClearAllPools()
            connection.Close()
            Call Update_Image()
            A = MsgBox("Are you sure you want to print Job Card", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Print Job Card ..........")
            If A = vbYes Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\JobCard.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T05Job_Card.T05Job_No} ='" & Trim(txtEntry.Text) & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If

            Call Load_Customer_name()
            Call Load_EntryNo()
            Call Load_VNO()
            Call Clear_text()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.ClearAllPools()
                connection.Close()
            End If
        End Try
    End Function

    Function Update_Image()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim IP As String
        Dim _STName As String
        Dim _PIC_Path As String
        Dim connection As New SqlConnection(ConfigurationManager.AppSettings("CD"))
        '  Dim command As New SqlCommand("insert into M31Vehicle_Master(M31Vehicle_No,M31BRAND,M31Pic,m31Engin_No,M31Chasis_no,M31Fuel,M31Type,M31Next_Lis,M31Next_Insu,M31Pic_Path,M31Status,M31Capacity) values(@name,@desc,@img,@ENG_NO,@M31Chasis_no,@M31Fuel,@M31Type,@M31Next_Lis,@M31Next_Insu,@M31Pic_Path,@M31Status,@M31Capacity)", connection)

        Try

            'MsgBox(Trim(txtEntry.Text))
            IP = ""
            Sql = "SELECT * FROM T05Job_Card WHERE T05Job_No='" & Trim(txtEntry.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)

            If isValidDataset(M01) Then

                Dim ms As New MemoryStream
                Dim ms1 As New MemoryStream
                '  ms.Dispose()
                PictureBox1.Image.Save(ms, PictureBox1.Image.RawFormat)
                ' PictureBox2.Image.Save(ms1, PictureBox2.Image.RawFormat)
                Dim command As New SqlCommand("UPDATE T05Job_Card SET T05Img=@Img WHERE  T05Job_No='" & Trim(txtEntry.Text) & "'", connection)
                command.Parameters.Add("@img", SqlDbType.Image).Value = ms.ToArray()
                ' command.Parameters.Add("@img1", SqlDbType.Image).Value = ms1.ToArray()
                connection.Open()
                If command.ExecuteNonQuery() = 1 Then
                    ' MsgBox("test1", MsgBoxStyle.Information, "Information .......")
                    '  MsgBox("Records update Successfully", MsgBoxStyle.Information, "Information .......")

                Else

                    '  MsgBox("test", MsgBoxStyle.Information, "Information .......")

                End If
                connection.ClearAllPools()
                connection.Close()
                ms.Dispose()

            End If
            con.ClearAllPools()
            con.CLOSE()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Function Clear_text()
        Dim A As String
        Try
            Call Load_EntryNo()
            Me.txtRemark.Text = ""
            Me.txtMtr.Text = ""
            Me.txtAddress.Text = ""
            Me.cboCus_Name.Text = ""
            Me.cboCus_Type.Text = ""
            Me.cboV_no.Text = ""
            Me.cbov_Type.Text = ""
            Me.cboBrand.Text = ""
            Me.txtRef_no.Text = ""
            Me.cboDepartment.Text = ""
            Me.cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            Me.txtTp.Text = ""
            A = ConfigurationManager.AppSettings("IMAGE") + "\images.jpg"
            PictureBox1.Image = Image.FromFile(A)
            txtPic1.Text = A
            txtRef_no.Focus()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
   

    Private Sub DeactivateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeactivateToolStripMenuItem.Click
        frmvew_Job.Close()
        frmvew_Job.Show()
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        If Trim(cboDepartment.Text) <> "" Then
        Else
            MsgBox("Please select the department", MsgBoxStyle.Information, "Information .........")
            cboDepartment.ToggleDropdown()
            Exit Sub
        End If

        If txtMtr.Text <> "" Then
        Else
            txtMtr.Text = "0"
        End If

        If IsNumeric(txtMtr.Text) Then
        Else
            MsgBox("Please enter the meter reading", MsgBoxStyle.Information, "Information .........")
            txtMtr.Focus()
            Exit Sub
        End If

        If Search_Vehicle_No() = True Then
        Else
            MsgBox("Please select the correct vehicle no", MsgBoxStyle.Information, "Information .........")
            cboV_no.ToggleDropdown()
            Exit Sub
        End If
        If txtTp.Text <> "" Then
        Else
            txtTp.Text = "-"
        End If

        If Trim(cboCus_Type.Text) <> "" Then
        Else
            MsgBox("Please select the customer type", MsgBoxStyle.Information, "Information .........")
            cboCus_Type.ToggleDropdown()
            Exit Sub
        End If

        If Trim(cboCus_Name.Text) <> "" Then
        Else
            MsgBox("Please select the customer name", MsgBoxStyle.Information, "Information ..........")
            cboCus_Name.ToggleDropdown()
            Exit Sub
        End If

        If txtAddress.Text <> "" Then
        Else
            txtAddress.Text = "-"
        End If

        If Trim(txtRemark.Text) <> "" Then
        Else
            txtRemark.Text = "-"
        End If

        If txtRef_no.Text <> "" Then
        Else
            txtRef_no.Text = "-"
        End If

        If Trim(cboBrand.Text) <> "" Then
        Else
            MsgBox("Please enter the Brand Name", MsgBoxStyle.Information, "Information ...........")
            cboBrand.ToggleDropdown()
            Exit Sub
        End If

        If Trim(cbov_Type.Text) <> "" Then
        Else
            MsgBox("Please enter the Vehicle Type", MsgBoxStyle.Information, "Information ...........")
            cboBrand.ToggleDropdown()
            Exit Sub
        End If

        Call EDIT_DATA()
    End Sub

    Function EDIT_DATA()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim MB51 As DataSet
        Dim _GetDate As DateTime
        Dim _Get_Time As DateTime
        Dim A As String

        Dim M01 As DataSet
        Dim M02 As DataSet
        Try
            nvcFieldList1 = "SELECT * FROM T05Job_Card WHERE T05Job_No='" & Trim(txtEntry.Text) & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                If Trim(M01.Tables(0).Rows(0)("T05Status")) = "CLOSE" Then
                    MsgBox("This Job no alrady close", MsgBoxStyle.Information, "Information ........")
                    connection.Close()
                    Exit Function
                Else
                    _GetDate = Month(txtDate.Text) & "/" & Microsoft.VisualBasic.Day(txtDate.Text) & "/" & Year(txtDate.Text)

                    _Get_Time = Month(txtDate.Text) & "/" & Microsoft.VisualBasic.Day(txtDate.Text) & "/" & Year(txtDate.Text) & " " & Hour(Now) & ":" & Minute(Now)

                    nvcFieldList1 = "UPDATE T05Job_Card SET T05Ref_No='" & Trim(txtRef_no.Text) & "',T05Date='" & _GetDate & "',T05Time='" & _Get_Time & "',T05Department='" & Trim(cboDepartment.Text) & "',T05Vehi_No='" & Trim(cboV_no.Text) & "',T05Cus_No='" & _CusNo & "',T05Mtr='" & txtMtr.Text & "',T05Remark='" & Trim(txtRemark.Text) & "',T05Status='A' WHERE T05Job_No='" & Trim(txtEntry.Text) & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "Insert Into tmpTransaction_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                         " values('JOB CARD','EDIT', '" & _Get_Time & "','" & strDisname & "','" & Trim(txtEntry.Text) & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ......")
                End If
            End If

            ' Call Update_Image()
            transaction.Commit()
            connection.ClearAllPools()
            connection.Close()
            Call Update_Image()

            Call Load_Customer_name()
            Call Load_EntryNo()
            Call Load_VNO()
            Call Clear_text()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.ClearAllPools()
                connection.Close()
            End If
        End Try
    End Function

    Private Sub CustomersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomersToolStripMenuItem.Click
        strWindowName = Me.Name
        frmView_Customer.Close()
        frmView_Customer.Show()
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
        Dim MB51 As DataSet
        Dim _GetDate As DateTime
        Dim _Get_Time As DateTime
        Dim A As String

        Dim M01 As DataSet
        Dim M02 As DataSet
        Try
            A = MsgBox("Are you sure you want to cancel this Job Card", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Cancel Job Card ............")
            If A = vbYes Then
                nvcFieldList1 = "SELECT * FROM T05Job_Card WHERE T05Job_No='" & Trim(txtEntry.Text) & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    If Trim(M01.Tables(0).Rows(0)("T05Status")) = "CLOSE" Then
                        MsgBox("This Job no alrady close", MsgBoxStyle.Information, "Information ........")
                        connection.Close()
                        Exit Sub
                    Else
                        _GetDate = Month(txtDate.Text) & "/" & Microsoft.VisualBasic.Day(txtDate.Text) & "/" & Year(txtDate.Text)

                        _Get_Time = Month(txtDate.Text) & "/" & Microsoft.VisualBasic.Day(txtDate.Text) & "/" & Year(txtDate.Text) & " " & Hour(Now) & ":" & Minute(Now)

                        nvcFieldList1 = "UPDATE T05Job_Card SET T05Status='CANCEL' WHERE T05Job_No='" & Trim(txtEntry.Text) & "'"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                        nvcFieldList1 = "Insert Into tmpTransaction_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                             " values('JOB CARD','CANCEL', '" & _Get_Time & "','" & strDisname & "','" & Trim(txtEntry.Text) & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                        MsgBox("Job card cancel successfully", MsgBoxStyle.Information, "Information ......")
                    End If
                End If
            End If
            ' Call Update_Image()
            transaction.Commit()
            connection.ClearAllPools()
            connection.Close()
            'Call Update_Image()

            Call Load_Customer_name()
            Call Load_EntryNo()
            Call Load_VNO()
            Call Clear_text()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.ClearAllPools()
                connection.Close()
            End If
        End Try
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_text()
    End Sub


    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim A As String
        Dim B As New ReportDocument
        Try
            A = ConfigurationManager.AppSettings("ReportPath") + "\JobCard.rpt"
            B.Load(A.ToString)
            B.SetDatabaseLogon("sa", "sainfinity")
            'B.SetParameterValue("To", _To)
            'B.SetParameterValue("From", _From)
            '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            frmReport.CrystalReportViewer1.SelectionFormula = "{T05Job_Card.T05Job_No} ='" & Trim(txtEntry.Text) & "' "
            frmReport.Refresh()
            ' frmReport.CrystalReportViewer1.PrintReport()
            ' B.PrintToPrinter(1, True, 0, 0)
            frmReport.MdiParent = MDIMain
            frmReport.Show()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                ' con.CLOSE()
            End If

        End Try
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        frmCustomer_Cnt.Close()
        frmCustomer_Cnt.Show()
    End Sub

   
    Private Sub VehicleRepairHistoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VehicleRepairHistoryToolStripMenuItem.Click
        frmView_Vehicle_History.Close()
        frmView_Vehicle_History.Show()
    End Sub
End Class