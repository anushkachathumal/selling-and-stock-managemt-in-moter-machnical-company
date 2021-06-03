Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
'Imports CrystalDecisions.CrystalReports.Engine
Public Class frmPurchasingSpairpart
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim dblInsuaranceCommision As Double
    Dim c_dataCustomer As DataTable
    Dim strPrice As Double
    Dim strTicket_price As Double
    Dim strSupplierscode As String
    Const MAX_SERIALS = 156000

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub


    Function Load_PO()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select M03OPNO as [P/O No],M03Yarn as [Yarn],M03DelivaryDate as [Delivary Date],M03Qty as [Quantity] from M03Purchase_Order where M03Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboPO.DataSource = M01
                cboPO.Rows.Band.Columns(0).Width = 125
                cboPO.Rows.Band.Columns(1).Width = 270
                cboPO.Rows.Band.Columns(2).Width = 170
                cboPO.Rows.Band.Columns(3).Width = 130
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_PO() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Search_PO = False
            Sql = "select M03Yarn,M03DelivaryDate,M03Qty,M01Name,M03StockCode from M03Purchase_Order inner join M01Supplier on M03SuppCode=M01Code where M03OPNO='" & Trim(cboPO.Text) & "' and M03Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql) '
            If isValidDataset(M01) Then
                With M01
                    txtCompany.Text = .Tables(0).Rows(0)("M01Name")
                    txtOQty.Text = .Tables(0).Rows(0)("M03Qty")
                    txtDdate.Text = .Tables(0).Rows(0)("M03DelivaryDate")
                    txtYarn.Text = .Tables(0).Rows(0)("M03Yarn")
                    txtStockCode.Text = .Tables(0).Rows(0)("M03StockCode")
                    Search_PO = True

                End With

                Sql = "select sum(T01Qty)as Qty from T01Transaction where T01OrderNo='" & Trim(cboPO.Text) & "' and T01Status='A' group by T01OrderNo"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql) '
                If isValidDataset(M01) Then
                    txtGRN.Text = M01.Tables(0).Rows(0)("Qty")
                Else
                    txtGRN.Text = "0"
                End If
                Search_PO = True
                cmdSave.Enabled = True
            Else
                Search_PO = False
                txtCompany.Text = ""
                txtOQty.Text = ""
                txtGRN.Text = ""
                txtDdate.Text = ""
                txtYarn.Text = ""
                cmdSave.Enabled = False
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Find_Purchaseorder()
        Dim strFileName As String
        strFileName = ConfigurationManager.AppSettings("FilePath") + "\PO.txt"
        Dim CurrGameWinningSerials(0 To MAX_SERIALS) As Long
        Dim fileHndl As Long
        Dim lLineNo As Long

        Dim strOrder, strQuality, strMaterial, _
      strSuppcode, strSuppName, strCutting_line, strPOQty, strDiscription, strDdate As String
        Dim strOrderqty As Double
        Dim strLineItem As String
        ' Dim strFileName As String '= _
        Dim strRolls As Double

        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim Sql As String
        Dim M01 As DataSet

        Dim M03Knittingorder As DataSet
        Dim ncQryType As String
        Dim nvcVccode As String
        Dim linesList As New List(Of String)(IO.File.ReadAllLines(strFileName))

        Try
            fileHndl = FreeFile()


            ' strFileName = Dir(strFileName)

            'UPGRADE_WARNING: Couldn't resolve default property of object fileHndl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object strValidSerialFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileOpen(fileHndl, strFileName, OpenMode.Input)
            lLineNo = 0
            Dim strRow As String

            Do Until EOF(fileHndl)

                '  Line Input #fileHndl, strRow
                'UPGRADE_WARNING: Couldn't resolve default property of object fileHndl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strRow = LineInput(fileHndl)


                If Trim(strRow) <> "" Then

                    If InStr(1, strRow, vbTab) > 0 Then
                        '  CurrGameWinningSerials(lLineNo) = Trim(Split(strRow, vbTab)(0))
                        strSuppcode = (Trim(Split(strRow, vbTab)(0)))
                        strSuppName = (Trim(Split(strRow, vbTab)(1)))
                        strOrder = (Trim(Split(strRow, vbTab)(2)))
                        strLineItem = (Trim(Split(strRow, vbTab)(3)))
                        strMaterial = (Trim(Split(strRow, vbTab)(4)))
                        strDiscription = (Trim(Split(strRow, vbTab)(5)))
                        strPOQty = (Trim(Split(strRow, vbTab)(6)))
                        ' strDiscription = (Trim(Split(strRow, vbTab)(15)))
                        strDdate = (Trim(Split(strRow, vbTab)(7)))
                        'str30class = (Trim(Split(strRow, vbTab)(9)))
                        'strRoot = (Trim(Split(strRow, vbTab)(10)))
                        'strMC = (Trim(Split(strRow, vbTab)(11)))
                        'strCutting_line = (Trim(Split(strRow, vbTab)(12)))
                        'strRolls = (Trim(Split(strRow, vbTab)(13)))
                        'strStatus = (Trim(Split(strRow, vbTab)(14)))


                        nvcFieldList1 = "select * from M01Supplier where M01Code='" & strSuppcode & "' "
                        M03Knittingorder = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                        If isValidDataset(M03Knittingorder) Then
                            'nvcFieldList1 = "UPDATE M03Knittingorder set M03Orderqty=" & strOrderqty & ",M03Yarnstock='" & strYarncode & "',M03YarnType='" & strYarntype & "',M03IType='" & strI_Type & "',M03LineItem='" & strLineItem & "',M0330Class='" & str30class & "',M03Root='" & strRoot & "',M03MCNo='" & strMC & "',M03CuttingLine='" & strCutting_line & "',M03NoofRoll='" & strRolls & "' where M03OrderNo='" & strOrder & "' and M03Quality='" & strQuality & "' and M03Material='" & strMaterial & "'"
                            'ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        Else
                            ''ncQryType = "ADD"
                            'nvcFieldList1 = "(M03OrderNo," & "M03Quality," & "M03Material," & "M03Orderqty," & "M03Yarnstock," & "M03YarnType," & "M03IType," & "M03LineItem," & "M0330Class," & "M03Root," & "M03MCNo," & "M03CuttingLine," & "M03NoofRoll," & "M03Status) " & "values('" & Trim(strOrder) & "','" & strQuality & "','" & strMaterial & "'," & strOrderqty & ",'" & strYarncode & "','" & strYarntype & "','" & strI_Type & "','" & strLineItem & "','" & str30class & "','" & strRoot & "','" & strMC & "','" & strCutting_line & "'," & strRolls & ",'A')"
                            'up_GetSetM03Knittingorder(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)

                            nvcFieldList1 = "Insert Into M01Supplier(M01Code,M01Name,M01Status)" & _
                                                      " values('" & strSuppcode & "', '" & Trim(strSuppName) & "','A')"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        End If

                        Sql = "select * from M03Purchase_Order where M03OPNO='" & strOrder & "'"
                        M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                        If isValidDataset(M01) Then
                            nvcFieldList1 = "Insert Into M03Purchase_Order(M03OPNO,M03SuppCode,M03Yarn,M03PODate,M03DelivaryDate,M03Qty,M03StockCode,M03Status)" & _
                                                          " values('" & Trim(strOrder) & "', '" & Trim(strSuppcode) & "','" & strMaterial & "','" & Today & "','" & strDdate & "'," & strPOQty & ",'" & strLineItem & "','A')"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        End If

                        linesList.RemoveAt(0)
                        ''  MsgBox(linesList.ToArray().ToString)
                        'IO.File.WriteAllLines(strFileName, linesList.ToArray())



                    Else
                        Err.Raise(vbObjectError + 18001, "GenerateInstantFile(str,str,str)", "Invalid Record At Line " & CStr(lLineNo))
                    End If

                End If

                lLineNo = lLineNo + 1

            Loop

            transaction.Commit()
            FileClose()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub frmPurchasingSpairpart_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Try
            txtDate.Text = Today
            txtInvoice.ReadOnly = True
            txtStockCode.ReadOnly = True

            'txtTotal.ReadOnly = True
            'txtTotal.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            'txtNqty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            'txtAmount.ReadOnly = True
            'txtAmount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtCompany.ReadOnly = True
            txtOQty.ReadOnly = True
            txtDdate.ReadOnly = True
            txtGRN.ReadOnly = True

            txtOQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtGRN.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            txtCompany.ReadOnly = True
            ' Call Loadgride()

            Call Load_PO()



        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim recParameter As DataSet
        Call Load_PO()
        Clicked = "ADD"
        OPR0.Enabled = True
        OPR1.Enabled = True
        OPR2.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False
        txtRef.Focus()
        txtDate.Text = Today

        Try
            Sql = "select * from P01parameter where P01code='GRN'"
            recParameter = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(recParameter) Then
                txtInvoice.Text = recParameter.Tables(0).Rows(0)("P01LastNo")
            End If

            cboPO.ToggleDropdown()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Function Loadgride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer = CustomerDataClass.MakeDataTableSpairpart
        'UltraGrid1.DataSource = c_dataCustomer
        'With UltraGrid1
        '.DisplayLayout.Bands(0).Columns(0).Width = 120
        '.DisplayLayout.Bands(0).Columns(1).AutoEdit = False
        '.DisplayLayout.Bands(0).Columns(1).Width = 230
        '.DisplayLayout.Bands(0).Columns(1).AutoEdit = False
        '.DisplayLayout.Bands(0).Columns(2).Width = 90
        '.DisplayLayout.Bands(0).Columns(2).AutoEdit = False
        '.DisplayLayout.Bands(0).Columns(3).Width = 90
        '.DisplayLayout.Bands(0).Columns(3).AutoEdit = False
        '.DisplayLayout.Bands(0).Columns(4).Width = 90
        '.DisplayLayout.Bands(0).Columns(4).AutoEdit = False
        '.DisplayLayout.Bands(0).Columns(5).Width = 90
        '.DisplayLayout.Bands(0).Columns(5).AutoEdit = False
        '.DisplayLayout.Bands(0).Columns(6).Width = 110
        '.DisplayLayout.Bands(0).Columns(6).AutoEdit = False
        'End With
    End Function

    Private Sub txtRef_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRef.KeyUp
        If e.KeyCode = 13 Then
            txtQty.Focus()
        End If
    End Sub

    Private Sub txtCustomer_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = 13 Then
            'cboItem.ToggleDropdown()
        End If
    End Sub
    Private Sub txtRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVNo.KeyUp
        If e.KeyCode = 13 Then
            txtContainer.Focus()
        End If
    End Sub

    Private Sub txtDrivers_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDrivers.KeyUp
        If e.KeyCode = 13 Then
            cmdSave.Focus()
        End If
    End Sub

    Private Sub txtDiscount_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDrivers.ValueChanged

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim ncQryType As String
        Dim nvcFieldList As String
        Dim nvcWhereClause As String
        Dim nvcVccode As String
        Dim i As Integer
        Dim _Roll As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim M01 As DataSet

        Try

            If Search_PO() = True Then
                nvcFieldList = "Update P01parameter set P01LastNo = " & Val(txtInvoice.Text) + 1 & " where P01code = 'GRN'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)

                nvcFieldList = "select * from T01Transaction where T01OrderNo='" & Trim(cboPO.Text) & "' and T01Status='A'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)

                If M01.Tables(0).Rows.Count = 0 Then
                    _Roll = "001"
                ElseIf M01.Tables(0).Rows.Count >= 9 Then
                    _Roll = "00" & M01.Tables(0).Rows.Count
                Else
                    _Roll = "0" & M01.Tables(0).Rows.Count
                End If


                nvcFieldList = "Insert Into T01Transaction(T01RefNo,T01OrderNo,T01ComInvoice,T01Date,T01Qty,T01Vehicle,T01Driver,T01No,T01User,T01Status,T01Container)" & _
                                                                                     " values(" & Trim(txtInvoice.Text) & ",'" & cboPO.Text & "','" & txtRef.Text & "'," & "convert(varchar(50),getdate(),102)" & ",'" & txtQty.Text & "','" & txtVNo.Text & "','" & txtDrivers.Text & "','" & _Roll & "','" & strDisname & "','A','" & Trim(txtContainer.Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)



                MsgBox("Records update sucessfully", MsgBoxStyle.Information, "Informaion ..........")
                transaction.Commit()
                common.ClearAll(OPR0, OPR1, OPR2)
                Clicked = ""
                cmdAdd.Enabled = True
                cmdSave.Enabled = False
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False

                cmdAdd.Focus()
                Call Loadgride()
            Else
                MsgBox("Please enter the correct Purchase order", MsgBoxStyle.Information, "Textured Jersey ........")
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

   
    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0, OPR2, OPR1)
        Clicked = ""
        cmdAdd.Enabled = True
        cmdSave.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        ' OPR4.Enabled = True
        cmdAdd.Focus()
        'txtAmount.Text = ""
        Loadgride()
    End Sub

  

    Private Sub txtCustomer_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs)

    End Sub


    Function SerachLocation() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        'SerachLocation = False
        'Sql = "select * from M08Location where M08LocationCode='" & Trim(cboLocation.Text) & "'"
        'dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
        'If isValidDataset(dsUser) Then
        '    SerachLocation = True
        '    'txtLocation.Text = dsUser.Tables(0).Rows(0)("M08Name")
        '    ' txtCurrentqty.Text = dsUser.Tables(0).Rows(0)("M05NewBalance")
        'Else
        '    SerachLocation = False
        '    'txtLocation.Text = ""
        '    'txtCurrentqty.Text = ""
        '    'txtNewQty.Text = ""
        'End If
    End Function

    Private Sub cboLocation_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs)
        SerachLocation()
    End Sub

    Private Sub cboLocation_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs)

    End Sub

    Private Sub cboPO_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPO.AfterCloseUp
        Call Search_PO()
    End Sub

    Private Sub cboPO_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboPO.InitializeLayout

    End Sub

    Private Sub cboPO_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboPO.KeyUp
        If e.KeyCode = 13 Then
            Call Search_PO()
            txtRef.Focus()
        End If
    End Sub

    Private Sub cboPO_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPO.TextChanged
        Call Search_PO()
    End Sub

    Private Sub txtRef_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRef.ValueChanged

    End Sub

    Private Sub txtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQty.KeyUp
        If e.KeyCode = 13 Then
            If txtQty.Text <> "" Then
                txtVNo.Focus()
            Else

            End If
        End If
    End Sub

    Private Sub txtQty_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtQty.ValueChanged
        If IsNumeric(txtQty.Text) Then
        Else
            txtQty.Text = ""
        End If
    End Sub

    Private Sub txtVNo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVNo.ValueChanged

    End Sub

    Private Sub UltraLabel10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraLabel10.Click

    End Sub

    Private Sub txtContainer_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtContainer.KeyUp
        If e.KeyCode = 13 Then
            txtDrivers.Focus()
        End If
    End Sub
End Class