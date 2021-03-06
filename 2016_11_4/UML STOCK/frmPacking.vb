Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmPacking
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _Itemcode As String

    Function Load_Gride_Item()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_StockIN
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 320
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 190
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False

            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            ' .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            '.DisplayLayout.Bands(0).Columns(1).CellAppearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            '.DisplayLayout.Bands(0).Columns(0).CellAppearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            ' .DisplayLayout.Bands(0).Columns(1).
            ' .DisplayLayout.Bands(0).Header.Height = 60

            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function


    Function Load_Date()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _QTY As Integer
        Try
            'Product Item 
            '---------------------------------------------->> OB (OPERNING BALANCE)
            '---------------------------------------------->> DR (SALES TRANSACTION)
            '---------------------------------------------->> RN (RETURN) REMARK USERBALE
            '---------------------------------------------->> PK (PACKING)
            '---------------------------------------------->> SI (STOCK IN)
            _QTY = 0
            i = 0

            Sql = "SELECT * FROM View_Production_Items WHERE M14Status='A' AND CATEGORY='PS'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows


                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Item Code") = M01.Tables(0).Rows(i)("M14Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M14Item_Name")
                ' newRow("Current Stock") = _QTY
                c_dataCustomer1.Rows.Add(newRow)




                i = i + 1
            Next

            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M14Item_Name as [##] from View_Production_Items where M14Status='A' and Category='PS' order by M14Item_Code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItem
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 220
                ' .Rows.Band.Columns(1).Width = 180


            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Parameter()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01PARAMETER where P01CODE='PK'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtBatch.Text = M01.Tables(0).Rows(0)("P01NO")
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Me.Close()
    End Sub



    Function Search_Itemcode() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _qty As Integer
        Dim _stockIn As Integer
        Try
            Sql = "select * from View_Production_Items where M14Status='A' and M14Item_Name='" & cboItem.Text & "'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _qty = 0
                _stockIn = 0
                _Itemcode = Trim(M01.Tables(0).Rows(0)("M14Item_code"))
                Search_Itemcode = True

                If Trim(M01.Tables(0).Rows(0)("category")) = "PS" Then

                    'PACKING
                    If strUGroup = "ADMIN" Or strUGroup = "MD" Or strUGroup = "ADD_ACCO" Then
                        Sql = "SELECT SUM(S02Qty) AS QTY FROM S02Set_Stock WHERE S02Status in ('A','HD') AND S02Tr_Type='PK' AND S02Pr_Code='" & Trim(M01.Tables(0).Rows(0)("M14Item_Code")) & "' AND S02Product_Status='GOOD' GROUP BY S02Pr_Code"
                    Else
                        Sql = "SELECT SUM(S02Qty) AS QTY FROM S02Set_Stock WHERE S02Status in ('A') AND S02Tr_Type='PK' AND S02Pr_Code='" & Trim(M01.Tables(0).Rows(0)("M14Item_Code")) & "' AND S02Product_Status='GOOD' GROUP BY S02Pr_Code"
                    End If
                    M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M02) Then
                        _qty = M02.Tables(0).Rows(0)("QTY")
                        _stockIn = _stockIn + M02.Tables(0).Rows(0)("QTY")

                    End If
                  
                    'RETURN
                    Sql = "SELECT SUM(S02Qty) AS QTY FROM S02Set_Stock WHERE S02Status in ('A','HD') AND S02Tr_Type='RN' AND S02Pr_Code='" & Trim(M01.Tables(0).Rows(0)("M14Item_Code")) & "' AND S02Product_Status='GOOD' GROUP BY S02Pr_Code"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M02) Then
                        _qty = _qty + M02.Tables(0).Rows(0)("QTY")
                        _stockIn = _stockIn + M02.Tables(0).Rows(0)("QTY")
                    End If

                    txtStockIN.Text = _stockIn

                    _stockIn = 0
                    'SALES
                    If strUGroup = "ADMIN" Or strUGroup = "MD" Or strUGroup = "ADD_ACCO" Then
                        Sql = "SELECT SUM(S02Qty) AS QTY FROM S02Set_Stock WHERE S02Status in ('A','HD') AND S02Tr_Type='DR' AND S02Pr_Code='" & Trim(M01.Tables(0).Rows(0)("M14Item_Code")) & "' AND S02Product_Status='GOOD' GROUP BY S02Pr_Code"
                    Else
                        Sql = "SELECT SUM(S02Qty) AS QTY FROM S02Set_Stock WHERE S02Status in ('A') AND S02Tr_Type='DR' AND S02Pr_Code='" & Trim(M01.Tables(0).Rows(0)("M14Item_Code")) & "' AND S02Product_Status='GOOD' GROUP BY S02Pr_Code"

                    End If
                    M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M02) Then
                        _stockIn = M02.Tables(0).Rows(0)("QTY")
                    End If
                    'GATE PASS

                    If strUGroup = "ADMIN" Or strUGroup = "MD" Or strUGroup = "ADD_ACCO" Then
                        Sql = "SELECT SUM(S02Qty) AS QTY FROM S02Set_Stock WHERE S02Status in ('A','HD') AND S02Tr_Type='GP' AND S02Pr_Code='" & Trim(M01.Tables(0).Rows(0)("M14Item_Code")) & "' AND S02Product_Status='GOOD' GROUP BY S02Pr_Code"
                    Else
                        Sql = "SELECT SUM(S02Qty) AS QTY FROM S02Set_Stock WHERE S02Status in ('A') AND S02Tr_Type='GP' AND S02Pr_Code='" & Trim(M01.Tables(0).Rows(0)("M14Item_Code")) & "' AND S02Product_Status='GOOD' GROUP BY S02Pr_Code"

                    End If
                    M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M02) Then
                        _stockIn = _stockIn + M02.Tables(0).Rows(0)("QTY")
                    End If

                    'UNPACKING
                    If strUGroup = "ADMIN" Or strUGroup = "MD" Or strUGroup = "ADD_ACCO" Then
                        Sql = "SELECT SUM(S02Qty) AS QTY FROM S02Set_Stock WHERE S02Status in ('A','HD') AND S02Tr_Type='UPK' AND S02Pr_Code='" & Trim(M01.Tables(0).Rows(0)("M14Item_Code")) & "'  GROUP BY S02Pr_Code"
                    Else
                        Sql = "SELECT SUM(S02Qty) AS QTY FROM S02Set_Stock WHERE S02Status in ('A') AND S02Tr_Type='UPK' AND S02Pr_Code='" & Trim(M01.Tables(0).Rows(0)("M14Item_Code")) & "'  GROUP BY S02Pr_Code"

                    End If
                    M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M02) Then
                        _stockIn = _stockIn + M02.Tables(0).Rows(0)("QTY")
                    End If

                    txtStockOut.Text = _stockIn
                    ' txtCurrent.Text = _qty
                Else

                End If

            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub cboItem_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboItem.AfterCloseUp
        Call Search_Itemcode()
    End Sub


    Private Sub cboItem_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboItem.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Itemcode()
            txtCurrent.Focus()
        End If
    End Sub

    Private Sub txtCurrent_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCurrent.KeyUp
        If e.KeyCode = 13 Then
            txtRemark.Focus()
        End If
    End Sub

    Private Sub txtRemark_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemark.KeyUp
        Dim I As Integer
        Try
            If e.KeyCode = 13 Then
                I = 0
                If IsNumeric(txtCurrent.Text) Then
                Else
                    MsgBox("Please enter the correct qty", MsgBoxStyle.Information, "Information .......")
                    txtCurrent.Focus()
                    Exit Sub
                End If

                If Search_Itemcode() = True Then
                Else
                    MsgBox("Please enter the correct Item ", MsgBoxStyle.Information, "Information ......")
                    cboItem.ToggleDropdown()
                    Exit Sub
                End If

                If txtCurrent.Text <> "" Then
                Else
                    txtCurrent.Text = "0"
                End If
                For Each uRow As UltraGridRow In UltraGrid1.Rows
                    If Trim(UltraGrid1.Rows(I).Cells(0).Text) = _Itemcode Then
                        UltraGrid1.Rows(I).Cells(2).Value = txtCurrent.Text
                        If Trim(txtRemark.Text) <> "" Then
                            UltraGrid1.Rows(I).Cells(3).Value = txtRemark.Text
                        Else
                            UltraGrid1.Rows(I).Cells(3).Value = "-"
                        End If
                        Exit For
                    End If
                    I = I + 1
                Next

                txtCurrent.Text = ""
                txtStockIN.Text = ""
                txtStockOut.Text = ""
                cboItem.Text = ""
                cboItem.ToggleDropdown()
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'con.close()
            End If
        End Try
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
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
        Dim _count As Integer
        Dim i As Integer
        Dim A As String
        Dim X As Integer
        Dim T01 As DataSet

        Try
            nvcFieldList1 = "UPDATE P01PARAMETER SET P01NO=P01NO + " & 1 & " WHERE P01CODE='PK'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            _count = 0
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                If IsNumeric(UltraGrid1.Rows(i).Cells(2).Text) Then
                    _count = _count + 1

                    nvcFieldList1 = "SELECT * FROM View_Production_Items WHERE M14Item_Code='" & UltraGrid1.Rows(i).Cells(0).Value & "' AND m14status='A' and Category='PS'"
                    MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(MB51) Then
                        nvcFieldList1 = "Insert Into S02Set_Stock(S02Tr_Type,S02Date,S02Pr_Code,S02Qty,S02Location,S02Status,S02User,S02Product_Status,S02Ref_No)" & _
                                                            " values('PK', '" & txtDate.Text & "','" & UltraGrid1.Rows(i).Cells(0).Text & "','" & UltraGrid1.Rows(i).Cells(2).Text & "','MS','A','" & strDisname & "','GOOD','" & txtBatch.Text & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                        nvcFieldList1 = "Insert Into T04Packing(T04Ref_No,T04Date,T04Item_Code,T04Qty,T04FG_Qty,T04Remark,T04User)" & _
                                                          " values('" & txtBatch.Text & "', '" & txtDate.Text & "','" & UltraGrid1.Rows(i).Cells(0).Text & "','" & UltraGrid1.Rows(i).Cells(2).Text & "','" & UltraGrid1.Rows(i).Cells(2).Text & "','" & UltraGrid1.Rows(i).Cells(3).Text & "','" & strDisname & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)



                        X = 0
                        nvcFieldList1 = "SELECT * FROM M16Item_for_Set WHERE M16Product_Code='" & UltraGrid1.Rows(i).Cells(0).Value & "' AND M16Status='A'"
                        T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                        For Each DTRow1 As DataRow In T01.Tables(0).Rows
                            nvcFieldList1 = "Insert Into S01Product_Stock(S01Tr_Type,S01Date,S01Item_Code,S01Qty,S01Location,S01Status,S01User,S01Product_Status,S01Ref_No)" & _
                                                           " values('PK', '" & txtDate.Text & "','" & T01.Tables(0).Rows(X)("M16Item_Code") & "','" & -(T01.Tables(0).Rows(X)("M16Qty") * UltraGrid1.Rows(i).Cells(2).Value) & "','MS','A','" & strDisname & "','GOOD','" & txtBatch.Text & "')"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                            X = X + 1
                        Next
                    End If
                End If
                i = i + 1
            Next

            MsgBox(_count & " Items successfully updated", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()
            connection.Close()
            Call Clear_Text()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub

    Function Clear_Text()
        Me.cboItem.Text = ""
        Me.txtCurrent.Text = ""
        Me.txtStockIN.Text = ""
        Me.txtStockOut.Text = ""
        Me.txtRemark.Text = ""
        Call Load_Parameter()
        Call Load_Gride_Item()
        Call Load_Date()
    End Function

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Clear_Text()
    End Sub

    Private Sub frmPacking_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride_Item()
        Call Load_Date()
        Call Load_Combo()
        Call Load_Parameter()

        txtDate.Text = Today
        txtCurrent.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtStockIN.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtStockOut.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtStockIN.ReadOnly = True
        txtStockOut.ReadOnly = True
        txtBatch.ReadOnly = True
        txtBatch.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
    End Sub

    Private Sub txtCurrent_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCurrent.ValueChanged

    End Sub
End Class