Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_Distributors

Public Class frmDNFProduction
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim T As Boolean

    Function Load_DyeRecepy()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim _Fromdate As Date
        Dim _Todate As Date

        _Fromdate = txtFromDate.Text & " " & "7:30AM"
        _Todate = CDate(txtFromDate.Text).AddDays(+1)

        Try
            SQL = "select sum(M04Batchwt) as M04Batchwt from M04Lot inner join M08Sub_Shade on M08Code=M04Shade_Code inner join T03Machine on M04Machine_No=T03Code  where M04ETime between '" & _Fromdate & "' and '" & _Todate & "' and T03Type in ('02','01') and M08CL='colour' group by M08CL"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                txtD1.Text = T01.Tables(0).Rows(0)("M04Batchwt")
            End If

            SQL = "select sum(M04Batchwt) as M04Batchwt from M04Lot inner join M08Sub_Shade on M08Code=m04shade_Type inner join T03Machine on M04Machine_No=T03Code  where M04ETime between '" & _Fromdate & "' and '" & _Todate & "' and T03Type in ('02','01') and M08CL='White' group by M08CL"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                txtD2.Text = T01.Tables(0).Rows(0)("M04Batchwt")
            End If

            SQL = "select sum(M04Batchwt) as M04Batchwt from M04Lot inner join M08Sub_Shade on M08Code=m04shade_Type inner join T03Machine on M04Machine_No=T03Code  where M04ETime between '" & _Fromdate & "' and '" & _Todate & "' and T03Type in ('02','01') and M08CL='MARL & Yarn dye' group by M08CL"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                txtD3.Text = T01.Tables(0).Rows(0)("M04Batchwt")
            End If
            '---------------------------------------------------------------------------
            'REPROCESSR
            Dim I As Integer

            I = 0
            txtR1.Text = ""
            SQL = "select sum(M04Batchwt) as M04Batchwt from M04Lot inner join T03Machine on M04Machine_No=T03Code WHERE M04PROGRAMETYPE IN ('W','O') AND  M04ETime between '" & _Fromdate & "' and '" & _Todate & "' and T03Type in ('02','01') group by M04PROGRAMETYPE"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                txtR1.Text = Val(txtR1.Text) + T01.Tables(0).Rows(I)("M04Batchwt")
                I = I + 1
            Next

            I = 0
            txtR2.Text = ""
            SQL = "select sum(M04Batchwt) as M04Batchwt from M04Lot inner join T03Machine on M04Machine_No=T03Code WHERE M04PROGRAMETYPE IN ('S') AND  M04ETime between '" & _Fromdate & "' and '" & _Todate & "' and T03Type in ('02','01') group by M04PROGRAMETYPE"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                txtR2.Text = Val(txtR2.Text) + T01.Tables(0).Rows(I)("M04Batchwt")
                I = I + 1
            Next

            I = 0
            txtR3.Text = ""
            SQL = "select sum(M04Batchwt) as M04Batchwt from M04Lot inner join M08Sub_Shade on M08Code=m04shade_Type inner join T03Machine on M04Machine_No=T03Code WHERE M04PROGRAMETYPE IN ('R') AND  M04ETime between '" & _Fromdate & "' and '" & _Todate & "' and T03Type in ('02','01') and M08CL='Colour' group by M04PROGRAMETYPE"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                txtR3.Text = Val(txtR3.Text) + T01.Tables(0).Rows(I)("M04Batchwt")
                I = I + 1
            Next

            I = 0
            txtR4.Text = ""
            SQL = "select sum(M04Batchwt) as M04Batchwt from M04Lot inner join M08Sub_Shade on M08Code=m04shade_Type inner join T03Machine on M04Machine_No=T03Code WHERE M04PROGRAMETYPE IN ('R') AND  M04ETime between '" & _Fromdate & "' and '" & _Todate & "' and T03Type in ('02','01') and M08CL='White' group by M04PROGRAMETYPE"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                txtR4.Text = Val(txtR4.Text) + T01.Tables(0).Rows(I)("M04Batchwt")
                I = I + 1
            Next

           
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click

        Call Search_RefDoc()

        Clicked = "ADD"
        OPR0.Enabled = True
        OPR2.Enabled = True
        OPR1.Enabled = True
        OPR3.Enabled = True
        OPR4.Enabled = True
        OPR5.Enabled = True
        OPR6.Enabled = True
        OPR7.Enabled = True
        OPR8.Enabled = True
        OPR9.Enabled = True
        OPR10.Enabled = True

        'OPR9.Enabled = True
        txtFromDate.Text = Today
        cmdAdd.Enabled = False
        cmdSave.Enabled = True
        txtP1.Focus()

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub frmDNFProduction_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Search_RefDoc()
        Call Load_DyeRecepy()
        txtRef.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRef.ReadOnly = True

        txtP1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtP2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtP3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtP4.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtP5.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtP6.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtD1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtD2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtD3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtY1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtY2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtY3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtO.Appearance.TextHAlign = Infragistics.Win.HAlign.Center


        txtR1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtR2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtR3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtR4.Appearance.TextHAlign = Infragistics.Win.HAlign.Center


        txtD1.ReadOnly = True
        txtD2.ReadOnly = True
        txtD3.ReadOnly = True

        txtT1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtT2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtT3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtT4.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtT5.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtT2.ReadOnly = True
        txtT4.ReadOnly = True
        txtT5.ReadOnly = True

        txtI1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtI2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtI3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtI4.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtI2.ReadOnly = True

        txtF1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtF2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtF3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtF4.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtF5.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtF6.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtF7.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtN1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtN2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtN3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtN4.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtN5.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtN6.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtF7.Appearance.TextHAlign = Infragistics.Win.HAlign.Center


        txtS1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtS2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtS3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtS4.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtS5.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txts6.Appearance.TextHAlign = Infragistics.Win.HAlign.Center


        Call Search_Records()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim nvcFieldList1 As String
        Dim Sql As String
        Dim X1 As Integer
        '  _TotalHR = "00:00"
        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True


        Dim M04Lot As DataSet
        Dim nvcVccode As String

        Dim ncQryType As String
        Dim M01 As DataSet
        Dim T01 As DataSet
        Dim hh1 As Integer
        Dim mm1 As Integer
        Dim _TimeDifferance1 As Date
        Dim n_year As Integer

        Dim vMax As Integer
        ncQryType = "ADD"


        Try
            'Validate OPR0 group box controls
            If IsNumeric(txtP1.Text) Then
            Else
                MsgBox("Please enter the correct Knitting Plan Kgs", MsgBoxStyle.Information, "Information .......")

                txtP1.Focus()
                txtP1.SelectionStart = 0
                txtP1.SelectionLength = Len(txtP1.Text)
                Exit Sub
            End If

            If IsNumeric(txtP2.Text) Then
            Else
                MsgBox("Please enter the correct Single Jersey Qty ", MsgBoxStyle.Information, "Information .......")

                txtP2.Focus()
                txtP2.SelectionStart = 0
                txtP2.SelectionLength = Len(txtP2.Text)
                Exit Sub
            End If


            If IsNumeric(txtP3.Text) Then
            Else
                MsgBox("Please enter the correct Rib/Interlock Qty", MsgBoxStyle.Information, "Information .......")

                txtP3.Focus()
                txtP3.SelectionStart = 0
                txtP3.SelectionLength = Len(txtP3.Text)
                Exit Sub
            End If

            If IsNumeric(txtP4.Text) Then
            Else
                MsgBox("Please enter the correct Knitted Qty TJL", MsgBoxStyle.Information, "Information .......")

                txtP4.Focus()
                txtP4.SelectionStart = 0
                txtP4.SelectionLength = Len(txtP4.Text)
                Exit Sub
            End If

            If IsNumeric(txtP5.Text) Then
            Else
                MsgBox("Please enter the correct Knitted Qty OCL", MsgBoxStyle.Information, "Information .......")

                txtP5.Focus()
                txtP5.SelectionStart = 0
                txtP5.SelectionLength = Len(txtP5.Text)
                Exit Sub
            End If

            If IsNumeric(txtP6.Text) Then
            Else
                MsgBox("Please enter the correct Knitting Downgrades", MsgBoxStyle.Information, "Information .......")

                txtP6.Focus()
                txtP6.SelectionStart = 0
                txtP6.SelectionLength = Len(txtP6.Text)
                Exit Sub
            End If

            Sql = "select * from PRODUCTION_FIGURES where EDate='" & txtFromDate.Text & "'"
            T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            If isValidDataset(T01) Then
                nvcFieldList1 = "Update PRODUCTION_FIGURES set Kni_PlnQty='" & txtP1.Text & "',S_JerseyQty='" & txtP2.Text & "',Rib_Qty='" & txtP3.Text & "',Knitted_Qty_Tj='" & txtP4.Text & "',Knitted_Qty_OCI='" & txtP5.Text & "',Knt_Downgrade='" & txtP6.Text & "' where EDate='" & txtFromDate.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                vMax = Get_highestVouNumber()
                nvcFieldList1 = "(RefNo," & "EDate," & "Kni_PlnQty," & "S_JerseyQty," & "Rib_Qty," & "Knitted_Qty_Tj," & "Knitted_Qty_OCI," & "Knt_Downgrade) " & "values(" & txtRef.Text & ",'" & txtFromDate.Text & "','" & txtP1.Text & "','" & txtP2.Text & "','" & txtP3.Text & "','" & txtP4.Text & "','" & txtP5.Text & "','" & txtP6.Text & "')"
                up_GetSetPRODUCTION_FIGURES(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)

            End If
            '-------------------------------------------------------------------------------------------
            'Validate OPR1 group box controls
            If IsNumeric(txtD1.Text) Then
            Else
                MsgBox("Please enter the correct Dyeing Plan Qty Colour", MsgBoxStyle.Information, "Information .......")

                txtD1.Focus()
                txtD1.SelectionStart = 0
                txtD1.SelectionLength = Len(txtD1.Text)
                Exit Sub
            End If

            If IsNumeric(txtD2.Text) Then
            Else
                MsgBox("Please enter the correct Dyeing Plan Qty White", MsgBoxStyle.Information, "Information .......")

                txtD2.Focus()
                txtD2.SelectionStart = 0
                txtD2.SelectionLength = Len(txtD2.Text)
                Exit Sub
            End If

            If IsNumeric(txtD3.Text) Then
            Else
                MsgBox("Please enter the correct Dyeing Plan Qty Marl", MsgBoxStyle.Information, "Information .......")

                txtD3.Focus()
                txtD3.SelectionStart = 0
                txtD3.SelectionLength = Len(txtD3.Text)
                Exit Sub
            End If

            Sql = "select * from PRODUCTION_FIGURES where EDate='" & txtFromDate.Text & "'"
            T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            If isValidDataset(T01) Then
                nvcFieldList1 = "Update PRODUCTION_FIGURES set Dye_PQty_Colour='" & txtD1.Text & "',Dye_PQty_White='" & txtD2.Text & "',Dye_PQty_Mal='" & txtD3.Text & "' where EDate='" & txtFromDate.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                vMax = Get_highestVouNumber()
                nvcFieldList1 = "(RefNo," & "EDate," & "Dye_PQty_Colour," & "Dye_PQty_White," & "Dye_PQty_Mal) " & "values(" & txtRef.Text & ",'" & txtFromDate.Text & "','" & txtD1.Text & "','" & txtD2.Text & "','" & txtD3.Text & "')"
                up_GetSetPRODUCTION_FIGURES(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)

            End If
            '-------------------------------------------------------------------------------------
            'YARN DYE

            If IsNumeric(txtY1.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Yarn Dyeing Plan Kgs", _
                                            "Information ...", _
                                             MessageBoxButtons.OK, _
                                            MessageBoxIcon.Information, _
                                             MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    txtY1.Focus()
                    txtY1.SelectionStart = 0
                    txtY1.SelectionLength = Len(txtY1.Text)
                    Exit Sub
                End If
            End If


            If IsNumeric(txtY2.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Yarn Dyeing Plan TJL Kgs", _
                                            "Information ...", _
                                             MessageBoxButtons.OK, _
                                            MessageBoxIcon.Information, _
                                             MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    txtY2.Focus()
                    txtY2.SelectionStart = 0
                    txtY2.SelectionLength = Len(txtY2.Text)
                    Exit Sub
                End If
            End If


            If IsNumeric(txtY3.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Yarn Dyeing Plan Kgs", _
                                            "Information ...", _
                                             MessageBoxButtons.OK, _
                                            MessageBoxIcon.Information, _
                                             MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    txtY3.Focus()
                    txtY3.SelectionStart = 0
                    txtY3.SelectionLength = Len(txtY3.Text)
                    Exit Sub
                End If
            End If

            Sql = "select * from PRODUCTION_FIGURES where EDate='" & txtFromDate.Text & "'"
            T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            If isValidDataset(T01) Then
                nvcFieldList1 = "Update PRODUCTION_FIGURES set Yarn_DyePQty='" & txtY1.Text & "',Yarn_DyeTjQty='" & txtY2.Text & "',Yarn_DyeComQty='" & txtY3.Text & "' where EDate='" & txtFromDate.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                vMax = Get_highestVouNumber()
                nvcFieldList1 = "(RefNo," & "EDate," & "Yarn_DyePQty," & "Yarn_DyeTjQty," & "Yarn_DyeComQty) " & "values(" & txtRef.Text & ",'" & txtFromDate.Text & "','" & txtY1.Text & "','" & txtY2.Text & "','" & txtY3.Text & "')"
                up_GetSetPRODUCTION_FIGURES(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)

            End If
            '----------------------------------------------------------------------------------------
            'OFFSHADE
            If IsNumeric(txtO.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct offshade Qty", _
                                            "Information ...", _
                                             MessageBoxButtons.OK, _
                                            MessageBoxIcon.Information, _
                                             MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    txtO.Focus()
                    txtO.SelectionStart = 0
                    txtO.SelectionLength = Len(txtO.Text)
                    Exit Sub
                End If
            End If

            Sql = "select * from PRODUCTION_FIGURES where EDate='" & txtFromDate.Text & "'"
            T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            If isValidDataset(T01) Then
                nvcFieldList1 = "Update PRODUCTION_FIGURES set Off_ShadeQty='" & txtO.Text & "' where EDate='" & txtFromDate.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                vMax = Get_highestVouNumber()
                nvcFieldList1 = "(RefNo," & "EDate," & "Off_ShadeQty) " & "values(" & txtRef.Text & ",'" & txtFromDate.Text & "','" & txtO.Text & "')"
                up_GetSetPRODUCTION_FIGURES(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)

            End If
            '---------------------------------------------------------------------------------------------
            'Reprocess



            Sql = "select * from PRODUCTION_FIGURES where EDate='" & txtFromDate.Text & "'"
            T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            If isValidDataset(T01) Then
                nvcFieldList1 = "Update PRODUCTION_FIGURES set R1_ReprocessOther='" & txtR1.Text & "',R2_Stripped='" & txtR2.Text & "',R3_RedyeColour='" & txtR3.Text & "',R4_RedyeWhite='" & txtR4.Text & "' where EDate='" & txtFromDate.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                vMax = Get_highestVouNumber()
                nvcFieldList1 = "(RefNo," & "EDate," & "R1_ReprocessOther," & "R2_Stripped," & "R3_RedyeColour," & "R4_RedyeWhite) " & "values(" & txtRef.Text & ",'" & txtFromDate.Text & "','" & txtR1.Text & "','" & txtR2.Text & "','" & txtR3.Text & "','" & txtR4.Text & "')"
                up_GetSetPRODUCTION_FIGURES(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)

            End If
            '----------------------------------------------------------------------------------------------
            If IsNumeric(txtT1.Text) Then
                txtT2.Text = (Val(txtT1.Text) / Val(txtT3.Text))
                txtT4.Text = (Val(txtT2.Text) / 340)
                txtT5.Text = Val(txtR3.Text) + Val(txtR4.Text) + Val(txtT1.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Reprocess Indirect Kgs", _
                                            "Information ...", _
                                             MessageBoxButtons.OK, _
                                            MessageBoxIcon.Information, _
                                             MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    txtT1.Focus()
                    txtT1.SelectionStart = 0
                    txtT1.SelectionLength = Len(txtT1.Text)
                    Exit Sub
                End If
            End If

            T = False
            If IsNumeric(txtT3.Text) Then
                txtT2.Text = (Val(txtT1.Text) / Val(txtT3.Text))
                txtT4.Text = (Val(txtT2.Text) / 340)
                txtT5.Text = Val(txtR3.Text) + Val(txtR4.Text) + Val(txtT1.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Total Reprocess Indirect Hrs", _
                                            "Information ...", _
                                             MessageBoxButtons.OK, _
                                            MessageBoxIcon.Information, _
                                             MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    txtT3.Focus()
                    txtT3.SelectionStart = 0
                    txtT3.SelectionLength = Len(txtT3.Text)
                    Exit Sub
                End If
            End If


            Sql = "select * from PRODUCTION_FIGURES where EDate='" & txtFromDate.Text & "'"
            T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            If isValidDataset(T01) Then
                nvcFieldList1 = "Update PRODUCTION_FIGURES set Reprocess_IndQTY='" & txtT1.Text & "',Tot_Reprocess_Ind='" & txtT2.Text & "' where EDate='" & txtFromDate.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                vMax = Get_highestVouNumber()
                nvcFieldList1 = "(RefNo," & "EDate," & "Reprocess_IndQTY," & "Tot_Reprocess_Ind) " & "values(" & txtRef.Text & ",'" & txtFromDate.Text & "','" & txtT1.Text & "','" & txtT2.Text & "')"
                up_GetSetPRODUCTION_FIGURES(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)

            End If
            '----------------------------------------------------------------------------------------------
            T = False
            If IsNumeric(txtI1.Text) Then

            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Inspected Qty Mtr", _
                                            "Information ...", _
                                             MessageBoxButtons.OK, _
                                            MessageBoxIcon.Information, _
                                             MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    txtI1.Focus()
                    txtI1.SelectionStart = 0
                    txtI1.SelectionLength = Len(txtI1.Text)
                    Exit Sub
                End If
            End If

            T = False
            If IsNumeric(txtI3.Text) Then
                txtI2.Text = Val(txtI3.Text) + Val(txtI4.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct W/H Booked Mtr TJL", _
                                            "Information ...", _
                                             MessageBoxButtons.OK, _
                                            MessageBoxIcon.Information, _
                                             MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    txtI3.Focus()
                    txtI3.SelectionStart = 0
                    txtI3.SelectionLength = Len(txtI3.Text)
                    Exit Sub
                End If
            End If


            T = False
            If IsNumeric(txtI4.Text) Then
                txtI2.Text = Val(txtI3.Text) + Val(txtI4.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct W/H Booked Mtr PTL", _
                                            "Information ...", _
                                             MessageBoxButtons.OK, _
                                            MessageBoxIcon.Information, _
                                             MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    txtI4.Focus()
                    txtI4.SelectionStart = 0
                    txtI4.SelectionLength = Len(txtI4.Text)
                    Exit Sub
                End If
            End If

            Sql = "select * from PRODUCTION_FIGURES where EDate='" & txtFromDate.Text & "'"
            T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            If isValidDataset(T01) Then
                nvcFieldList1 = "Update PRODUCTION_FIGURES set Inspected_QtyMtr='" & txtI1.Text & "',Warehouse_BLKMtrTJ='" & txtI3.Text & "',Warehouse_BLKMtrPTL='" & txtI3.Text & "' where EDate='" & txtFromDate.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                vMax = Get_highestVouNumber()
                nvcFieldList1 = "(RefNo," & "EDate," & "Inspected_QtyMtr," & "Warehouse_BLKMtrTJ," & "Warehouse_BLKMtrPTL) " & "values(" & txtRef.Text & ",'" & txtFromDate.Text & "','" & txtI1.Text & "','" & txtI3.Text & "','" & txtI4.Text & "')"
                up_GetSetPRODUCTION_FIGURES(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)

            End If

            '-----------------------------------------------------------------------------------------------
            'FINAL FINISH
            T = False
            If IsNumeric(txtF1.Text) Then
                lblReDyeTot.Text = Val(txtF1.Text) + Val(txtF2.Text) + Val(txtF3.Text) + Val(lblReDye.Text) + Val(txtF6.Text) + Val(txtF7.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Final Finish Mtr", _
                                            "Information ...", _
                                             MessageBoxButtons.OK, _
                                            MessageBoxIcon.Information, _
                                             MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    txtF1.Focus()
                    txtF1.SelectionStart = 0
                    txtF1.SelectionLength = Len(txtF1.Text)
                    Exit Sub
                End If
            End If

            T = False
            If IsNumeric(txtF2.Text) Then
                lblReDyeTot.Text = Val(txtF1.Text) + Val(txtF2.Text) + Val(txtF3.Text) + Val(lblReDye.Text) + Val(txtF6.Text) + Val(txtF7.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Prepare for Print", _
                                            "Information ...", _
                                             MessageBoxButtons.OK, _
                                            MessageBoxIcon.Information, _
                                             MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    txtF2.Focus()
                    txtF2.SelectionStart = 0
                    txtF2.SelectionLength = Len(txtF2.Text)
                    Exit Sub
                End If
            End If

            T = False
            If IsNumeric(txtF3.Text) Then
                lblReDyeTot.Text = Val(txtF1.Text) + Val(txtF2.Text) + Val(txtF3.Text) + Val(lblReDye.Text) + Val(txtF6.Text) + Val(txtF7.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Daily Preset Qty Mtr", _
                                            "Information ...", _
                                             MessageBoxButtons.OK, _
                                            MessageBoxIcon.Information, _
                                             MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    txtF3.Focus()
                    txtF3.SelectionStart = 0
                    txtF3.SelectionLength = Len(txtF3.Text)
                    Exit Sub
                End If
            End If

            T = False
            If IsNumeric(txtF4.Text) Then
                lblReDye.Text = Val(txtF4.Text) + Val(txtF5.Text)
                lblReDyeTot.Text = Val(txtF1.Text) + Val(txtF2.Text) + Val(txtF3.Text) + Val(lblReDye.Text) + Val(txtF6.Text) + Val(txtF7.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Refinish Due to Offshade", _
                                            "Information ...", _
                                             MessageBoxButtons.OK, _
                                            MessageBoxIcon.Information, _
                                             MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    txtF4.Focus()
                    txtF4.SelectionStart = 0
                    txtF4.SelectionLength = Len(txtF4.Text)
                    Exit Sub
                End If
            End If

            T = False
            If IsNumeric(txtF5.Text) Then
                lblReDye.Text = Val(txtF4.Text) + Val(txtF5.Text)
                lblReDyeTot.Text = Val(txtF1.Text) + Val(txtF2.Text) + Val(txtF3.Text) + Val(lblReDye.Text) + Val(txtF6.Text) + Val(txtF7.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Refinish Due to Other", _
                                            "Information ...", _
                                             MessageBoxButtons.OK, _
                                            MessageBoxIcon.Information, _
                                             MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    txtF5.Focus()
                    txtF5.SelectionStart = 0
                    txtF5.SelectionLength = Len(txtF5.Text)
                    Exit Sub
                End If
            End If

            T = False
            If IsNumeric(txtF6.Text) Then
                lblReDyeTot.Text = Val(txtF1.Text) + Val(txtF2.Text) + Val(txtF3.Text) + Val(lblReDye.Text) + Val(txtF6.Text) + Val(txtF7.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Refinish Due to Finish", _
                                            "Information ...", _
                                             MessageBoxButtons.OK, _
                                            MessageBoxIcon.Information, _
                                             MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    txtF6.Focus()
                    txtF6.SelectionStart = 0
                    txtF6.SelectionLength = Len(txtF6.Text)
                    Exit Sub
                End If
            End If

            T = False
            If IsNumeric(txtF7.Text) Then
                lblReDyeTot.Text = Val(txtF1.Text) + Val(txtF2.Text) + Val(txtF3.Text) + Val(lblReDye.Text) + Val(txtF6.Text) + Val(txtF7.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Refinish Due to Other", _
                                            "Information ...", _
                                             MessageBoxButtons.OK, _
                                            MessageBoxIcon.Information, _
                                             MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    txtF7.Focus()
                    txtF7.SelectionStart = 0
                    txtF7.SelectionLength = Len(txtF7.Text)
                    Exit Sub
                End If
            End If

            Sql = "select * from PRODUCTION_FIGURES where EDate='" & txtFromDate.Text & "'"
            T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            If isValidDataset(T01) Then
                nvcFieldList1 = "Update PRODUCTION_FIGURES set F_Finishmtr='" & txtF1.Text & "',Dailly_PreQty='" & txtF3.Text & "',Prepare_Print='" & txtF2.Text & "',RF_DueOff='" & txtF6.Text & "',RF_DueOthr='" & txtF7.Text & "',Due_Offshaid='" & txtF4.Text & "',Due_Other='" & txtF5.Text & "' where EDate='" & txtFromDate.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                vMax = Get_highestVouNumber()
                nvcFieldList1 = "(RefNo," & "EDate," & "F_Finishmtr," & "Dailly_PreQty," & "Prepare_Print," & "Due_Offshaid," & "Due_Other," & "RF_DueOff," & "RF_DueOthr) " & "values(" & txtRef.Text & ",'" & txtFromDate.Text & "','" & txtF1.Text & "','" & txtF3.Text & "','" & txtF2.Text & "','" & txtF4.Text & "','" & txtF5.Text & "','" & txtF6.Text & "','" & txtF7.Text & "')"
                up_GetSetPRODUCTION_FIGURES(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)

            End If

            '----------------------------------------------------------------------------------------
            'NON CONFERMATION
            T = False
            If Trim(txtN1.Text) <> "" Then
                If IsNumeric(txtN1.Text) Then
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Non Conformance Kgs", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtN1
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If


            T = False
            If Trim(txtN2.Text) <> "" Then
                If IsNumeric(txtN2.Text) Then
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct AW 1st Bulk Approval Kgs", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtN2
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If

            T = False
            If Trim(txtN3.Text) <> "" Then
                If IsNumeric(txtN3.Text) Then
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Block Stock D & Kgs", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtN3
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If

            T = False
            If Trim(txtN4.Text) <> "" Then
                If IsNumeric(txtN4.Text) Then
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Sample Production", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtN4
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If


            T = False
            If Trim(txtN5.Text) <> "" Then
                If IsNumeric(txtN5.Text) Then
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Block Stock Quality Kgs", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtN5
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If

            T = False
            If Trim(txtN6.Text) <> "" Then
                If IsNumeric(txtN6.Text) Then
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Dye Plan Qty", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtN6
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If

            Sql = "select * from PRODUCTION_FIGURES where EDate='" & txtFromDate.Text & "'"
            T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            If isValidDataset(T01) Then
                nvcFieldList1 = "Update PRODUCTION_FIGURES set Non_Conformance_KG='" & txtN1.Text & "',AW_Appr='" & txtN2.Text & "',BLK_StokD='" & txtN3.Text & "',BLK_Stock_Quality='" & txtN5.Text & "',Dye_PlnQty='" & txtN6.Text & "',Sample_Pro='" & txtN4.Text & "' where EDate='" & txtFromDate.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                vMax = Get_highestVouNumber()
                nvcFieldList1 = "(RefNo," & "EDate," & "Non_Conformance_KG," & "AW_Appr," & "BLK_StokD," & "BLK_Stock_Quality," & "Dye_PlnQty," & "Sample_Pro) " & "values(" & txtRef.Text & ",'" & txtFromDate.Text & "','" & txtN1.Text & "','" & txtN2.Text & "','" & txtN3.Text & "','" & txtN5.Text & "','" & txtN6.Text & "','" & txtN4.Text & "')"
                up_GetSetPRODUCTION_FIGURES(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)

            End If

            '-----------------------------------------------------------------------------------------
            'YARN STOCK
            T = False
            If Trim(txtS1.Text) <> "" Then
                If IsNumeric(txtS1.Text) Then
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Yarn Stock  Kgs", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtS1
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If

            T = False
            If Trim(txtS2.Text) <> "" Then
                If IsNumeric(txtS2.Text) Then
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Greige Stock Kgs", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtS2
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If

            T = False
            If Trim(txtS3.Text) <> "" Then
                If IsNumeric(txtS3.Text) Then
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Warehouse Stock Mts", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtS3
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If


            T = False
            If Trim(txtS4.Text) <> "" Then
                If IsNumeric(txtS4.Text) Then
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Delivered Qty Mts", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtS4
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If

            T = False
            If Trim(txtS5.Text) <> "" Then
                If IsNumeric(txtS5.Text) Then
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Late Delivery Kgs", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtS5
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If

            Sql = "select * from PRODUCTION_FIGURES where EDate='" & txtFromDate.Text & "'"
            T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            If isValidDataset(T01) Then
                nvcFieldList1 = "Update PRODUCTION_FIGURES set Yarn_Stock='" & txtS1.Text & "',Grage_Stock='" & txtS2.Text & "',WH_Stock='" & txtS3.Text & "',Div_Qty='" & txtS4.Text & "',Last_Dil_Qty='" & txtS5.Text & "' where EDate='" & txtFromDate.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                vMax = Get_highestVouNumber()
                nvcFieldList1 = "(RefNo," & "EDate," & "Yarn_Stock," & "Grage_Stock," & "WH_Stock," & "Div_Qty," & "Last_Dil_Qty) " & "values(" & txtRef.Text & ",'" & txtFromDate.Text & "','" & txtS1.Text & "','" & txtS2.Text & "','" & txtS3.Text & "','" & txtS4.Text & "','" & txtS5.Text & "')"
                up_GetSetPRODUCTION_FIGURES(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)

            End If



            MsgBox("Record updateing successfully", MsgBoxStyle.Information, "Information ........")
            transaction.Commit()


            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Function Get_highestVouNumber() As String
        Dim con = New SqlConnection()
        Dim vMax As String

        '=======================================================================
        Try
            con = DBEngin.GetConnection()
            dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_GetSetParameter", New SqlParameter("@cQryType", "UPD"), New SqlParameter("@vcCode", "PF"))
            If common.isValidDataset(dsUser) Then
                For Each DTRow As DataRow In dsUser.Tables(0).Rows
                    vMax = dsUser.Tables(0).Rows(0)("P01LastNo")
                    Return vMax
                Next
            Else
                MessageBox.Show("Record Not Found", "Textured Jersey", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
            '===================================================================
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
        '=========================================================================
        ' "asdasd"
    End Function

    Function Search_RefDoc()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet

        Try
            SQL = "select * from P01Parameter where P01code='PF'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                txtRef.Text = T01.Tables(0).Rows(0)("P01lastno")
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_Records()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet

        Try
            SQL = "select * from PRODUCTION_FIGURES where EDate='" & txtFromDate.Text & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                txtRef.Text = T01.Tables(0).Rows(0)("RefNo")
                If IsDBNull(T01.Tables(0).Rows(0)("Kni_PlnQty")) Then
                Else
                    txtP1.Text = CInt(T01.Tables(0).Rows(0)("Kni_PlnQty"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("S_JerseyQty")) Then
                Else
                    txtP2.Text = CInt(T01.Tables(0).Rows(0)("S_JerseyQty"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Rib_Qty")) Then
                Else
                    txtP3.Text = CInt(T01.Tables(0).Rows(0)("Rib_Qty"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Knitted_Qty_Tj")) Then
                Else
                    txtP4.Text = CInt(T01.Tables(0).Rows(0)("Knitted_Qty_Tj"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Knitted_Qty_OCI")) Then
                Else
                    txtP5.Text = CInt(T01.Tables(0).Rows(0)("Knitted_Qty_OCI"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Knt_Downgrade")) Then
                Else
                    txtP6.Text = CInt(T01.Tables(0).Rows(0)("Knt_Downgrade"))
                End If
                'txtD1.Text = CInt(T01.Tables(0).Rows(0)("Dye_PQty_Colour"))
                'txtD2.Text = CInt(T01.Tables(0).Rows(0)("Dye_PQty_White"))
                'txtD3.Text = CInt(T01.Tables(0).Rows(0)("Dye_PQty_Mal"))
                If IsDBNull(T01.Tables(0).Rows(0)("Yarn_DyePQty")) Then
                Else
                    txtY1.Text = CInt(T01.Tables(0).Rows(0)("Yarn_DyePQty"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Yarn_DyeTjQty")) Then
                Else
                    txtY2.Text = CInt(T01.Tables(0).Rows(0)("Yarn_DyeTjQty"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Yarn_DyeComQty")) Then
                Else
                    txtY3.Text = CInt(T01.Tables(0).Rows(0)("Yarn_DyeComQty"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Off_ShadeQty")) Then
                Else
                    txtO.Text = CInt(T01.Tables(0).Rows(0)("Off_ShadeQty"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("R1_ReprocessOther")) Then
                Else
                    txtR1.Text = CInt(T01.Tables(0).Rows(0)("R1_ReprocessOther"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("R2_Stripped")) Then
                Else
                    txtR2.Text = CInt(T01.Tables(0).Rows(0)("R2_Stripped"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("R3_RedyeColour")) Then
                Else
                    txtR3.Text = CInt(T01.Tables(0).Rows(0)("R3_RedyeColour"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("R4_RedyeWhite")) Then
                Else
                    txtR4.Text = CInt(T01.Tables(0).Rows(0)("R4_RedyeWhite"))
                End If

                If IsDBNull(T01.Tables(0).Rows(0)("Reprocess_IndQTY")) Then
                Else
                    txtT1.Text = CInt(T01.Tables(0).Rows(0)("Reprocess_IndQTY"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Tot_Reprocess_Ind")) Then
                Else
                    txtT3.Text = CInt(T01.Tables(0).Rows(0)("Tot_Reprocess_Ind"))
                End If

                txtT2.Text = (Val(txtT1.Text) / Val(txtT3.Text))
                txtT4.Text = (Val(txtT2.Text) / 340)
                txtT5.Text = Val(txtR3.Text) + Val(txtR4.Text) + Val(txtT1.Text)

                If IsDBNull(T01.Tables(0).Rows(0)("Inspected_QtyMtr")) Then
                Else
                    txtI1.Text = CInt(T01.Tables(0).Rows(0)("Inspected_QtyMtr"))
                End If

                If IsDBNull(T01.Tables(0).Rows(0)("Warehouse_BLKMtrTJ")) Then
                Else
                    txtI3.Text = CInt(T01.Tables(0).Rows(0)("Warehouse_BLKMtrTJ"))
                End If

                If IsDBNull(T01.Tables(0).Rows(0)("Warehouse_BLKMtrPTL")) Then
                Else
                    txtI4.Text = CInt(T01.Tables(0).Rows(0)("Warehouse_BLKMtrPTL"))
                End If

                If IsDBNull(T01.Tables(0).Rows(0)("F_Finishmtr")) Then
                Else
                    txtF1.Text = CInt(T01.Tables(0).Rows(0)("F_Finishmtr"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Dailly_PreQty")) Then
                Else
                    txtF3.Text = CInt(T01.Tables(0).Rows(0)("Dailly_PreQty"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Prepare_Print")) Then
                Else
                    txtF2.Text = CInt(T01.Tables(0).Rows(0)("Prepare_Print"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Due_Offshaid")) Then
                Else
                    txtF4.Text = CInt(T01.Tables(0).Rows(0)("Due_Offshaid"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Due_Other")) Then
                Else
                    txtF5.Text = CInt(T01.Tables(0).Rows(0)("Due_Other"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("RF_DueOff")) Then
                Else
                    txtF6.Text = CInt(T01.Tables(0).Rows(0)("RF_DueOff"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("RF_DueOthr")) Then
                Else
                    txtF7.Text = CInt(T01.Tables(0).Rows(0)("RF_DueOthr"))
                End If


                If IsDBNull(T01.Tables(0).Rows(0)("Non_Conformance_KG")) Then
                Else
                    txtN1.Text = CInt(T01.Tables(0).Rows(0)("Non_Conformance_KG"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("AW_Appr")) Then
                Else
                    txtN2.Text = CInt(T01.Tables(0).Rows(0)("AW_Appr"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("BLK_StokD")) Then
                Else
                    txtN3.Text = CInt(T01.Tables(0).Rows(0)("BLK_StokD"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("BLK_Stock_Quality")) Then
                Else
                    txtN5.Text = CInt(T01.Tables(0).Rows(0)("BLK_Stock_Quality"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Dye_PlnQty")) Then
                Else
                    txtN6.Text = CInt(T01.Tables(0).Rows(0)("Dye_PlnQty"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Sample_Pro")) Then
                Else
                    txtN4.Text = CInt(T01.Tables(0).Rows(0)("Sample_Pro"))
                End If



                If IsDBNull(T01.Tables(0).Rows(0)("Yarn_Stock")) Then
                Else
                    txtS1.Text = CInt(T01.Tables(0).Rows(0)("Yarn_Stock"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Grage_Stock")) Then
                Else
                    txtS2.Text = CInt(T01.Tables(0).Rows(0)("Grage_Stock"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("WH_Stock")) Then
                Else
                    txtS3.Text = CInt(T01.Tables(0).Rows(0)("WH_Stock"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Div_Qty")) Then
                Else
                    txtS4.Text = CInt(T01.Tables(0).Rows(0)("Div_Qty"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Last_Dil_Qty")) Then
                Else
                    txtS5.Text = CInt(T01.Tables(0).Rows(0)("Last_Dil_Qty"))
                End If


            End If



            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub txtP1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtP1.KeyUp
        If e.KeyCode = 13 Then
            txtP2.Focus()
        End If
    End Sub

    Private Sub txtP1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtP1.ValueChanged

    End Sub

    Private Sub txtP2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtP2.KeyUp
        If e.KeyCode = 13 Then
            txtP3.Focus()
        End If
    End Sub

    Private Sub txtP2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtP2.ValueChanged

    End Sub

    Private Sub txtP3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtP3.KeyUp
        If e.KeyCode = 13 Then
            txtP4.Focus()
        End If
    End Sub

    Private Sub txtP3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtP3.ValueChanged

    End Sub

    Private Sub txtP4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtP4.KeyUp
        If e.KeyCode = 13 Then
            txtP5.Focus()
        End If
    End Sub

    Private Sub txtP4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtP4.ValueChanged

    End Sub

    Private Sub txtP5_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtP5.KeyUp
        If e.KeyCode = 13 Then
            txtP6.Focus()
        End If
    End Sub

    Private Sub txtP5_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtP5.ValueChanged

    End Sub

    Private Sub txtFromDate_BeforeDropDown(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFromDate.BeforeDropDown

    End Sub

    Private Sub txtFromDate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFromDate.TextChanged
        Call Search_Records()
        Call Load_DyeRecepy()
    End Sub

    Private Sub txtP6_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtP6.KeyUp
        If e.KeyCode = 13 Then
            txtY1.Focus()
        End If
    End Sub

    Private Sub txtP6_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtP6.ValueChanged

    End Sub

    Private Sub txtD2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtD2.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtD3.Focus()
            End If
        End If
    End Sub

    Private Sub txtD2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtD2.ValueChanged
        T = False
        If txtD1.Text <> "" Then
            If IsNumeric(txtD1.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Colour Qty", _
"Information ...", _
MessageBoxButtons.OK, _
MessageBoxIcon.Information, _
MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtD1
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                End If
                Exit Sub
                End If
        End If

        If txtD2.Text <> "" Then
            If IsNumeric(txtD2.Text) Then
            Else

                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct White Qty", _
"Information ...", _
MessageBoxButtons.OK, _
MessageBoxIcon.Information, _
MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtD2
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
                End If
        End If

        If txtD3.Text <> "" Then
            If IsNumeric(txtD3.Text) Then
            Else

                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Marl Qty", _
"Information ...", _
MessageBoxButtons.OK, _
MessageBoxIcon.Information, _
MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtD3
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
                End If
        End If

        lblTotD.Text = Val(txtD1.Text) + Val(txtD2.Text) + Val(txtD3.Text)
    End Sub

    Private Sub txtD1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtD1.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtD2.Focus()
            End If
        End If
    End Sub

    Private Sub txtD1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtD1.ValueChanged
        Dim A As DialogResult
        T = False
        If txtD1.Text <> "" Then
            If IsNumeric(txtD1.Text) Then
            Else

                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Colour Qty", _
"Information ...", _
MessageBoxButtons.OK, _
MessageBoxIcon.Information, _
MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtD1
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If

        If txtD2.Text <> "" Then
            If IsNumeric(txtD2.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct White Qty", _
 "Information ...", _
MessageBoxButtons.OK, _
MessageBoxIcon.Information, _
MessageBoxDefaultButton.Button2)
                If result3 = DialogResult.OK Then
                    With txtD2
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If

        If txtD3.Text <> "" Then
            If IsNumeric(txtD3.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Marl Qty", _
         "Information ...", _
     MessageBoxButtons.OK, _
     MessageBoxIcon.Information, _
     MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    With txtD3
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                End If
                Exit Sub
                End If
        End If

        lblTotD.Text = Val(txtD1.Text) + Val(txtD2.Text) + Val(txtD3.Text)
    End Sub

    Private Sub txtD3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtD3.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtY1.Focus()
            End If

        End If
    End Sub

    Private Sub txtD3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtD3.ValueChanged
        T = False
        If txtD1.Text <> "" Then
            If IsNumeric(txtD1.Text) Then
            Else

                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Colour Qty", _
"Information ...", _
MessageBoxButtons.OK, _
MessageBoxIcon.Information, _
MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtD1
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If

        If txtD2.Text <> "" Then
            If IsNumeric(txtD2.Text) Then
            Else

                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct White Qty", _
"Information ...", _
MessageBoxButtons.OK, _
MessageBoxIcon.Information, _
MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtD2
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If


        If txtD3.Text <> "" Then
            If IsNumeric(txtD3.Text) Then
            Else

                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Marl Qty", _
"Information ...", _
MessageBoxButtons.OK, _
MessageBoxIcon.Information, _
MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtD3
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If

        lblTotD.Text = Val(txtD1.Text) + Val(txtD2.Text) + Val(txtD3.Text)
    End Sub

    Private Sub txtY1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtY1.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtY2.Focus()
            End If
        End If
    End Sub

    Private Sub txtY1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtY1.ValueChanged
        T = False
        If Trim(txtY1.Text) <> "" Then
            If IsNumeric(txtY1.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Details", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtY1
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtY2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtY2.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtY3.Focus()

            End If
        End If
    End Sub

    Private Sub txtY2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtY2.Validated

    End Sub

    Private Sub txtY2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtY2.ValueChanged
        T = False
        If Trim(txtY2.Text) <> "" Then
            If IsNumeric(txtY2.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Details", _
    "Information ...", _
    MessageBoxButtons.OK, _
    MessageBoxIcon.Information, _
    MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtY2
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtY3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtY3.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtO.Focus()
            End If

        End If
    End Sub

    Private Sub txtY3_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtY3.ValueChanged
        T = False
        If Trim(txtY3.Text) <> "" Then
            If IsNumeric(txtY3.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Details", _
    "Information ...", _
    MessageBoxButtons.OK, _
    MessageBoxIcon.Information, _
    MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtY3
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtO_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtO.KeyUp
        If e.KeyCode = 13 Then
            txtT1.Focus()
        End If
    End Sub

    Private Sub txtO_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtO.ValueChanged
        T = False
        If IsNumeric(txtO.Text) Then
        Else
            Dim result3 As DialogResult = MessageBox.Show("Please enter the correct offshade Qty", _
                                        "Information ...", _
                                         MessageBoxButtons.OK, _
                                        MessageBoxIcon.Information, _
                                         MessageBoxDefaultButton.Button2)
            If result3 = Windows.Forms.DialogResult.OK Then
                T = True
                txtO.Focus()
                txtO.SelectionStart = 0
                txtO.SelectionLength = Len(txtO.Text)
                Exit Sub
            End If
        End If
    End Sub

   
  

    Private Sub txtT1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtT1.KeyUp
        If e.KeyCode = 13 Then
            txtT3.Focus()
        End If
    End Sub

    Private Sub txtT1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtT1.ValueChanged
        T = False
        If IsNumeric(txtT1.Text) Then
            txtT2.Text = (Val(txtT1.Text) / Val(txtT3.Text))
            txtT4.Text = (Val(txtT2.Text) / 340)
            txtT5.Text = Val(txtR3.Text) + Val(txtR4.Text) + Val(txtT1.Text)
        Else
            Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Reprocess Indirect Kgs", _
                                        "Information ...", _
                                         MessageBoxButtons.OK, _
                                        MessageBoxIcon.Information, _
                                         MessageBoxDefaultButton.Button2)
            If result3 = Windows.Forms.DialogResult.OK Then
                T = True
                txtT1.Focus()
                txtT1.SelectionStart = 0
                txtT1.SelectionLength = Len(txtT1.Text)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtT3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtT3.KeyUp
        If e.KeyCode = 13 Then
            txtI1.Focus()
        End If
    End Sub

    Private Sub txtT3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtT3.ValueChanged
        T = False
        If IsNumeric(txtT3.Text) Then
            txtT2.Text = (Val(txtT1.Text) / Val(txtT3.Text))
            txtT4.Text = (Val(txtT2.Text) / 340)
            txtT5.Text = Val(txtR3.Text) + Val(txtR4.Text) + Val(txtT1.Text)
        Else
            Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Total Reprocess Indirect Hrs", _
                                        "Information ...", _
                                         MessageBoxButtons.OK, _
                                        MessageBoxIcon.Information, _
                                         MessageBoxDefaultButton.Button2)
            If result3 = Windows.Forms.DialogResult.OK Then
                T = True
                txtT3.Focus()
                txtT3.SelectionStart = 0
                txtT3.SelectionLength = Len(txtT3.Text)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtT4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtT4.ValueChanged

    End Sub

    Private Sub txtI1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtI1.KeyUp
        If e.KeyCode = 13 Then
            txtI3.Focus()
        End If
    End Sub

    Private Sub txtI1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtI1.ValueChanged
        T = False
        If IsNumeric(txtI1.Text) Then

        Else
            Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Inspected Qty Mtr", _
                                        "Information ...", _
                                         MessageBoxButtons.OK, _
                                        MessageBoxIcon.Information, _
                                         MessageBoxDefaultButton.Button2)
            If result3 = Windows.Forms.DialogResult.OK Then
                T = True
                txtI1.Focus()
                txtI1.SelectionStart = 0
                txtI1.SelectionLength = Len(txtI1.Text)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtI3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtI3.KeyUp
        If e.KeyCode = 13 Then
            txtI4.Focus()
        End If
    End Sub

    Private Sub txtI3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtI3.ValueChanged
        T = False
        If IsNumeric(txtI3.Text) Then
            txtI2.Text = Val(txtI3.Text) + Val(txtI4.Text)
        Else
            Dim result3 As DialogResult = MessageBox.Show("Please enter the correct W/H Booked Mtr TJL", _
                                        "Information ...", _
                                         MessageBoxButtons.OK, _
                                        MessageBoxIcon.Information, _
                                         MessageBoxDefaultButton.Button2)
            If result3 = Windows.Forms.DialogResult.OK Then
                T = True
                txtI3.Focus()
                txtI3.SelectionStart = 0
                txtI3.SelectionLength = Len(txtI3.Text)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtI4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtI4.KeyUp
        If e.KeyCode = 13 Then
            txtF1.Focus()
        End If
    End Sub

    Private Sub txtI4_SystemColorsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtI4.SystemColorsChanged

    End Sub

    Private Sub txtI4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtI4.ValueChanged
        T = False
        If IsNumeric(txtI4.Text) Then
            txtI2.Text = Val(txtI3.Text) + Val(txtI4.Text)
        Else
            Dim result3 As DialogResult = MessageBox.Show("Please enter the correct W/H Booked Mtr PTL", _
                                        "Information ...", _
                                         MessageBoxButtons.OK, _
                                        MessageBoxIcon.Information, _
                                         MessageBoxDefaultButton.Button2)
            If result3 = Windows.Forms.DialogResult.OK Then
                T = True
                txtI4.Focus()
                txtI4.SelectionStart = 0
                txtI4.SelectionLength = Len(txtI4.Text)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtF1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtF1.KeyUp
        If e.KeyCode = 13 Then
            txtF2.Focus()

        End If
    End Sub

    Private Sub txtF1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF1.ValueChanged
        T = False
        If IsNumeric(txtF1.Text) Then
            lblReDyeTot.Text = Val(txtF1.Text) + Val(txtF2.Text) + Val(txtF3.Text) + Val(lblReDye.Text) + Val(txtF6.Text) + Val(txtF7.Text)
        Else
            Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Final Finish Mtr", _
                                        "Information ...", _
                                         MessageBoxButtons.OK, _
                                        MessageBoxIcon.Information, _
                                         MessageBoxDefaultButton.Button2)
            If result3 = Windows.Forms.DialogResult.OK Then
                T = True
                txtF1.Focus()
                txtF1.SelectionStart = 0
                txtF1.SelectionLength = Len(txtF1.Text)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtF2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtF2.KeyUp
        If e.KeyCode = 13 Then
            txtF3.Focus()
        End If
    End Sub

    Private Sub txtF2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF2.ValueChanged
        T = False
        If IsNumeric(txtF2.Text) Then
            lblReDyeTot.Text = Val(txtF1.Text) + Val(txtF2.Text) + Val(txtF3.Text) + Val(lblReDye.Text) + Val(txtF6.Text) + Val(txtF7.Text)
        Else
            Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Prepare for Print", _
                                        "Information ...", _
                                         MessageBoxButtons.OK, _
                                        MessageBoxIcon.Information, _
                                         MessageBoxDefaultButton.Button2)
            If result3 = Windows.Forms.DialogResult.OK Then
                T = True
                txtF2.Focus()
                txtF2.SelectionStart = 0
                txtF2.SelectionLength = Len(txtF2.Text)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtF3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtF3.KeyUp
        If e.KeyCode = 13 Then
            txtF4.Focus()

        End If
    End Sub

    Private Sub txtF3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF3.ValueChanged
        T = False
        If IsNumeric(txtF3.Text) Then
            lblReDyeTot.Text = Val(txtF1.Text) + Val(txtF2.Text) + Val(txtF3.Text) + Val(lblReDye.Text) + Val(txtF6.Text) + Val(txtF7.Text)
        Else
            Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Daily Preset Qty Mtr", _
                                        "Information ...", _
                                         MessageBoxButtons.OK, _
                                        MessageBoxIcon.Information, _
                                         MessageBoxDefaultButton.Button2)
            If result3 = Windows.Forms.DialogResult.OK Then
                T = True
                txtF3.Focus()
                txtF3.SelectionStart = 0
                txtF3.SelectionLength = Len(txtF3.Text)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtF4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtF4.KeyUp
        If e.KeyCode = 13 Then
            txtF5.Focus()
        End If
    End Sub

    Private Sub txtF4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF4.ValueChanged
        T = False
        If IsNumeric(txtF4.Text) Then
            lblReDye.Text = Val(txtF4.Text) + Val(txtF5.Text)
            lblReDyeTot.Text = Val(txtF1.Text) + Val(txtF2.Text) + Val(txtF3.Text) + Val(lblReDye.Text) + Val(txtF6.Text) + Val(txtF7.Text)
        Else
            Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Refinish Due to Offshade", _
                                        "Information ...", _
                                         MessageBoxButtons.OK, _
                                        MessageBoxIcon.Information, _
                                         MessageBoxDefaultButton.Button2)
            If result3 = Windows.Forms.DialogResult.OK Then
                T = True
                txtF4.Focus()
                txtF4.SelectionStart = 0
                txtF4.SelectionLength = Len(txtF4.Text)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtF5_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtF5.KeyUp
        If e.KeyCode = 13 Then
            txtF6.Focus()
        End If
    End Sub

    Private Sub txtF5_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF5.ValueChanged
        T = False
        If IsNumeric(txtF5.Text) Then
            lblReDye.Text = Val(txtF4.Text) + Val(txtF5.Text)
            lblReDyeTot.Text = Val(txtF1.Text) + Val(txtF2.Text) + Val(txtF3.Text) + Val(lblReDye.Text) + Val(txtF6.Text) + Val(txtF7.Text)
        Else
            Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Refinish Due to Other", _
                                        "Information ...", _
                                         MessageBoxButtons.OK, _
                                        MessageBoxIcon.Information, _
                                         MessageBoxDefaultButton.Button2)
            If result3 = Windows.Forms.DialogResult.OK Then
                T = True
                txtF5.Focus()
                txtF5.SelectionStart = 0
                txtF5.SelectionLength = Len(txtF5.Text)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtF6_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtF6.KeyUp
        If e.KeyCode = 13 Then
            txtF7.Focus()

        End If
    End Sub

    Private Sub txtF6_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF6.ValueChanged
        T = False
        If IsNumeric(txtF6.Text) Then
            lblReDyeTot.Text = Val(txtF1.Text) + Val(txtF2.Text) + Val(txtF3.Text) + Val(lblReDye.Text) + Val(txtF6.Text) + Val(txtF7.Text)
        Else
            Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Refinish Due to Finish", _
                                        "Information ...", _
                                         MessageBoxButtons.OK, _
                                        MessageBoxIcon.Information, _
                                         MessageBoxDefaultButton.Button2)
            If result3 = Windows.Forms.DialogResult.OK Then
                T = True
                txtF6.Focus()
                txtF6.SelectionStart = 0
                txtF6.SelectionLength = Len(txtF6.Text)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtF7_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtF7.KeyUp
        If e.KeyCode = 13 Then
            txtN1.Focus()

        End If
    End Sub

    Private Sub txtF7_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF7.ValueChanged
        T = False
        If IsNumeric(txtF7.Text) Then
            lblReDyeTot.Text = Val(txtF1.Text) + Val(txtF2.Text) + Val(txtF3.Text) + Val(lblReDye.Text) + Val(txtF6.Text) + Val(txtF7.Text)
        Else
            Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Refinish Due to Other", _
                                        "Information ...", _
                                         MessageBoxButtons.OK, _
                                        MessageBoxIcon.Information, _
                                         MessageBoxDefaultButton.Button2)
            If result3 = Windows.Forms.DialogResult.OK Then
                T = True
                txtF7.Focus()
                txtF7.SelectionStart = 0
                txtF7.SelectionLength = Len(txtF7.Text)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtN1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtN1.KeyUp
        If e.KeyCode = 13 Then
            txtN2.Focus()
        End If
    End Sub

    Private Sub txtN1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtN1.ValueChanged
        T = False
        If Trim(txtN1.Text) <> "" Then
            If IsNumeric(txtN1.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Non Conformance Kgs", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtN1
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtN2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtN2.KeyUp
        If e.KeyCode = 13 Then
            txtN3.Focus()
        End If
    End Sub

    Private Sub txtN2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtN2.ValueChanged
        T = False
        If Trim(txtN2.Text) <> "" Then
            If IsNumeric(txtN2.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct AW 1st Bulk Approval Kgs", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtN2
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtN3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtN3.KeyUp
        If e.KeyCode = 13 Then
            txtN4.Focus()

        End If
    End Sub

    Private Sub txtN3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtN3.ValueChanged
        T = False
        If Trim(txtN3.Text) <> "" Then
            If IsNumeric(txtN3.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Block Stock D & Kgs", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtN3
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtN4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtN4.KeyUp
        If e.KeyCode = 13 Then
            txtN5.Focus()
        End If
    End Sub

    Private Sub txtN4_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtN4.ValueChanged
        T = False
        If Trim(txtN4.Text) <> "" Then
            If IsNumeric(txtN4.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Sample Production", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtN4
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtN5_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtN5.KeyUp
        If e.KeyCode = 13 Then
            txtN6.Focus()
        End If
    End Sub

    Private Sub txtN5_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtN5.ValueChanged
        T = False
        If Trim(txtN5.Text) <> "" Then
            If IsNumeric(txtN5.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Block Stock Quality Kgs", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtN5
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtN6_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtN6.KeyUp
        If e.KeyCode = 13 Then

            If T = False Then
                txtS1.Focus()
            End If
        End If
    End Sub

    Private Sub txtN6_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtN6.ValueChanged
        T = False
        If Trim(txtN6.Text) <> "" Then
            If IsNumeric(txtN6.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Dye Plan Qty", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtN6
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtS1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtS1.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtS2.Focus()


            End If
        End If
    End Sub

    Private Sub txtS1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtS1.ValueChanged
        T = False
        If Trim(txtS1.Text) <> "" Then
            If IsNumeric(txtS1.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Yarn Stock  Kgs", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtS1
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtS2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtS2.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtS3.Focus()
            End If
        End If
    End Sub

    Private Sub txtS2_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtS2.ValueChanged
        T = False
        If Trim(txtS2.Text) <> "" Then
            If IsNumeric(txtS2.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Greige Stock Kgs", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtS2
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtS3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtS3.KeyUp
        If e.KeyCode = 13 Then
            txtS4.Focus()

        End If
    End Sub

    Private Sub txtS3_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtS3.ValueChanged
        T = False
        If Trim(txtS3.Text) <> "" Then
            If IsNumeric(txtS3.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Warehouse Stock Mts", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtS3
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtS4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtS4.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtS5.Focus()

            End If
        End If
    End Sub

    Private Sub txtS4_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtS4.ValueChanged
        T = False
        If Trim(txtS4.Text) <> "" Then
            If IsNumeric(txtS4.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Delivered Qty Mts", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtS4
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtS5_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtS5.KeyUp
        If e.KeyCode = 13 Then
            cmdSave.Focus()

        End If
    End Sub

    Private Sub txtS5_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtS5.ValueChanged
        T = False
        If Trim(txtS5.Text) <> "" Then
            If IsNumeric(txtS5.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Late Delivery Kgs", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtS5
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0, OPR1, OPR2, OPR3, OPR4, OPR5, OPR6, OPR7, OPR8, OPR9, OPR10)
        Clicked = ""
        cmdAdd.Enabled = True
        cmdAdd.Focus()
    End Sub
End Class