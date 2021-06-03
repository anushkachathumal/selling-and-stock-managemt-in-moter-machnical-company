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
Public Class frmSupplier_Cnt
    Dim _PrintStatus As String
    Dim _From As String
    Dim _TO As String
    Private M01 As DataSet
    Dim _Sales_ref As String
    Dim _Root As String

    Private Sub frmSupplier_Cnt_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmView_Supplier.Close()
        SqlConnection.ClearAllPools()
    End Sub

    Private Sub frmSupplier_Cnt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtCode.ReadOnly = True
        txtCode.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        Call Load_cOUNTRY()
        Call Load_Entry()
    End Sub

    Function Load_cOUNTRY()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        ' Dim M01 As DataSet
        con = DBEngin.GetConnection()
        Try
            Sql = "select M20Name as [##] from M20Country "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCountry
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 285
                '  .Rows.Band.Columns(1).Width = 160


            End With
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
            End If
        End Try
    End Function

  

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Function Load_Entry()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()


        Try
            Sql = "select * from P01Parameter where  P01Code='SP' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01No") >= 1 And M01.Tables(0).Rows(0)("P01No") < 10 Then
                    txtCode.Text = "SP/00" & M01.Tables(0).Rows(0)("P01No")
                ElseIf M01.Tables(0).Rows(0)("P01No") >= 10 And M01.Tables(0).Rows(0)("P01No") < 100 Then
                    txtCode.Text = "SP/0" & M01.Tables(0).Rows(0)("P01No")
                Else
                    txtCode.Text = "SP/" & M01.Tables(0).Rows(0)("P01No")
                End If
            End If


            con.ClearAllPools()
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
            End If
        End Try
    End Function

    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        Call Load_Entry()
        cboCountry.ToggleDropdown()

    End Sub

   
    Private Sub cboCountry_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCountry.KeyUp
        If e.KeyCode = 13 Then
            txtSupplier.Focus()
        End If
    End Sub

    Private Sub txtSupplier_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSupplier.KeyUp
        If e.KeyCode = 13 Then
            If Trim(txtSupplier.Text) <> "" Then
                txtContact_person.Focus()
            End If
        End If
    End Sub

    Private Sub txtContact_person_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtContact_person.KeyUp
        If e.KeyCode = 13 Then
            txtAddress.Focus()
        End If
    End Sub

    Private Sub txtAddress_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAddress.KeyUp
        If e.KeyCode = Keys.F1 Then
            txtVAT.Focus()
        End If
    End Sub

    Private Sub txtVAT_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVAT.KeyUp
        If e.KeyCode = 13 Then
            txtTP.Focus()
        End If
    End Sub

    Private Sub txtTP_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTP.KeyUp
        If e.KeyCode = 13 Then
            txtEmail.Focus()
        End If
    End Sub

    Private Sub txtEmail_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEmail.KeyUp
        If e.KeyCode = 13 Then
            cmdAdd.Focus()
        End If
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If cboCountry.Text <> "" Then
        Else
            MsgBox("Please select the Country Name", MsgBoxStyle.Information, "Information ........")
            cboCountry.ToggleDropdown()
            Exit Sub
        End If

        If Trim(txtSupplier.Text) <> "" Then
        Else
            MsgBox("Please enter the Supplier Name", MsgBoxStyle.Information, "Information ........")
            Exit Sub
        End If

        If txtContact_person.Text <> "" Then
        Else
            txtContact_person.Text = "-"
        End If

        If txtTP.Text <> "" Then
        Else
            txtTP.Text = "-"
        End If

        If txtAddress.Text <> "" Then
        Else
            txtAddress.Text = "-"

        End If

        If txtVAT.Text <> "" Then
        Else
            txtVAT.Text = "-"
        End If

        'If txtPic1.Text <> "" Then
        'Else
        '    MsgBox("Please select the Document 01 Image", MsgBoxStyle.Information, "Information .........")
        '    Exit Sub
        'End If

        'If txtPic2.Text <> "" Then
        'Else
        '    MsgBox("Please select the Document 02 Image", MsgBoxStyle.Information, "Information .........")
        '    Exit Sub
        'End If

        Call Save_Data()
    End Sub

    Function Save_Data()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        'connection.ClearAllPools()
        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim MB51 As DataSet
        Try

            nvcFieldList1 = "SELECT * FROM M04Supplier WHERE M04Code='" & Trim(txtCode.Text) & "' "
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                nvcFieldList1 = "update M04Supplier set M04Country='" & Trim(cboCountry.Text) & "',M04Name='" & Trim(txtSupplier.Text) & "',M05Con_Person='" & Trim(txtContact_person.Text) & "',M04Address='" & Trim(txtAddress.Text) & "',M04VAT_No='" & Trim(txtVAT.Text) & "',M04Contact_No='" & Trim(txtTP.Text) & "',M04email='" & Trim(txtEmail.Text) & "',M07Status='A' where M04Code='" & Trim(txtCode.Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                               " values('NEW_SUPPLIER','EDIT', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                nvcFieldList1 = "update P01Parameter set P01No=P01No+ " & 1 & " where P01Code='SP' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into M04Supplier(M04Code,M04Country,M04Name,M05Con_Person,M04Address,M04VAT_No,M04Contact_No,M04email,M04Status)" & _
                                                                  " values('" & Trim(txtCode.Text) & "','" & Trim(cboCountry.Text) & "', '" & Trim(txtSupplier.Text) & "','" & Trim(txtContact_person.Text) & "','" & Trim(txtAddress.Text) & "','" & Trim(txtVAT.Text) & "','" & Trim(txtTP.Text) & "','" & Trim(txtEmail.Text) & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                                " values('NEW_SUPPLIER','SAVE', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            End If
            MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()
            connection.ClearAllPools()
            connection.Close()

            Call Clear_Text()
            frmGRN_uniq.Load_Supplier()
            cboCountry.ToggleDropdown()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.ClearAllPools()
            End If
        End Try
    End Function

    Function Deactivate_Data()
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

            nvcFieldList1 = "SELECT * FROM M04Supplier WHERE M04Code='" & Trim(txtCode.Text) & "' "
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                nvcFieldList1 = "update M04Supplier set M04Status='I' where M04Code='" & Trim(txtCode.Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                               " values('NEW_SUPPLIER','DEACTIVE', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
           

            End If
            MsgBox("Supplier Deactivated successfully", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()
            connection.ClearAllPools()
            connection.Close()

            Call Clear_Text()
            frmGRN_uniq.Load_Supplier()
            cboCountry.ToggleDropdown()

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
        ' Dim M01 As DataSet
        Dim IP As String
        Dim _STName As String
        Dim _PIC_Path As String
        Dim connection As New SqlConnection(ConfigurationManager.AppSettings("CD"))
        '  Dim command As New SqlCommand("insert into M31Vehicle_Master(M31Vehicle_No,M31BRAND,M31Pic,m31Engin_No,M31Chasis_no,M31Fuel,M31Type,M31Next_Lis,M31Next_Insu,M31Pic_Path,M31Status,M31Capacity) values(@name,@desc,@img,@ENG_NO,@M31Chasis_no,@M31Fuel,@M31Type,@M31Next_Lis,@M31Next_Insu,@M31Pic_Path,@M31Status,@M31Capacity)", connection)

        Try

            'MsgBox(Trim(txtEntry.Text))
            IP = ""
            Sql = "SELECT * FROM M07Supplier WHERE M07Sup_Code='" & Trim(txtCode.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)

            If isValidDataset(M01) Then

                Dim ms As New MemoryStream
                Dim ms1 As New MemoryStream
                '  ms.Dispose()
                PictureBox1.Image.Save(ms, PictureBox1.Image.RawFormat)
                PictureBox2.Image.Save(ms1, PictureBox2.Image.RawFormat)
                Dim command As New SqlCommand("UPDATE M07Supplier SET M07Img_1=@Img,M07Img_2=@Img1 WHERE  M07Sup_Code='" & Trim(txtCode.Text) & "'", connection)
                command.Parameters.Add("@img", SqlDbType.Image).Value = ms.ToArray()
                command.Parameters.Add("@img1", SqlDbType.Image).Value = ms1.ToArray()
                connection.Open()
                If command.ExecuteNonQuery() = 1 Then
                    ' MsgBox("test1", MsgBoxStyle.Information, "Information .......")
                    '  MsgBox("Records update Successfully", MsgBoxStyle.Information, "Information .......")

                Else

                    '  MsgBox("test", MsgBoxStyle.Information, "Information .......")

                End If
                connection.ClearAllPools()
                ms.Dispose()

            End If
            con.ClearAllPools()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function
    Function Clear_Text()
        Call Load_Entry()
        Me.cboCountry.Text = ""
        Me.txtVAT.Text = ""
        Me.txtEmail.Text = ""
        Me.txtPic1.Text = ""
        Me.txtPic2.Text = ""
        Me.txtSupplier.Text = ""
        Me.txtTP.Text = ""
        Me.txtEmail.Text = ""
        Me.txtAddress.Text = ""
        Me.txtContact_person.Text = ""
        PictureBox1.Image = Nothing
        PictureBox2.Image = Nothing

    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_Text()
    End Sub

    Private Sub cmdpic_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdpic_1.Click
        On Error Resume Next
        OpenFileDialog1.Filter = "Image Files|*.jpg;*.gif;*.png;*.bmp"
        OpenFileDialog1.ShowDialog()
        PictureBox1.Image = Image.FromFile(OpenFileDialog1.FileName)
        txtPic1.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub cmdPic_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPic_2.Click
        On Error Resume Next
        OpenFileDialog1.Filter = "Image Files|*.jpg;*.gif;*.png;*.bmp"
        OpenFileDialog1.ShowDialog()
        PictureBox2.Image = Image.FromFile(OpenFileDialog1.FileName)
        txtPic2.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        frmView_Supplier.Close()
        frmView_Supplier.Show()
    End Sub

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        '  Dim M01 As DataSet
        Dim Value As Double

        Try
            Sql = "select  *  from M04Supplier where M04Code='" & Trim(txtCode.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboCountry.Text = Trim(M01.Tables(0).Rows(0)("M04Country"))
                txtSupplier.Text = Trim(M01.Tables(0).Rows(0)("M04Name"))
                txtAddress.Text = Trim(M01.Tables(0).Rows(0)("M04Address"))
                txtContact_person.Text = Trim(M01.Tables(0).Rows(0)("M05Con_Person"))
                txtVAT.Text = Trim(M01.Tables(0).Rows(0)("M04VAT_No"))
                txtTP.Text = Trim(M01.Tables(0).Rows(0)("M04Contact_No"))
                txtEmail.Text = Trim(M01.Tables(0).Rows(0)("M04email"))
                ' txtTP.Text = Trim(M01.Tables(0).Rows(0)("M07TP"))
           
                'txtPic1.Text = "-"
                'txtPic2.Text = "-"
                'Dim arrayImage() As Byte = CType(M01.Tables(0).Rows(0)("M07Img_1"), Byte())
                'Dim ms As New MemoryStream(arrayImage)
                'PictureBox1.Image = Image.FromStream(ms)
                'Dim arrayImage1() As Byte = CType(M01.Tables(0).Rows(0)("M07Img_2"), Byte())
                'Dim ms1 As New MemoryStream(arrayImage)
                'PictureBox2.Image = Image.FromStream(ms1)
            End If


            con.ClearAllPools()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
            End If
        End Try
    End Function


    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Dim A As String

        A = MsgBox("Are you sure you want to deactivate this suppliere", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Deactivate ..........")
        If A = vbYes Then
            Call Deactivate_Data()
        End If

    End Sub

    Private Sub txtSupplier_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSupplier.ValueChanged

    End Sub
End Class