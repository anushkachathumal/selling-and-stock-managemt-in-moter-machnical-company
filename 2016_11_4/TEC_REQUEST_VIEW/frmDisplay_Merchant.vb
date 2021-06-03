Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmDisplay_Merchant
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String

    Private Sub frmDisplay_Merchant_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmView_Marchant.MdiParent = MDIMain
        ' m_ChildFormNumber += 1
        ' frmWinner.Text = "WNF"
        frmView_Marchant.Show()
    End Sub

    Private Sub frmDisplay_Merchant_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboCustomer.ReadOnly = True
        cboEnd_User.ReadOnly = True
        cboHanger.ReadOnly = True
        cboTest_Standard.ReadOnly = True

        txtTarger_Price.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTarget_Weight.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTarget_Width.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtOrder_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtYardage.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtRequierd_Date.ReadOnly = True
        txtReuest_Date.ReadOnly = True
        txtFabric_Des.ReadOnly = True
        txtOrder_Qty.ReadOnly = True
        txtTarger_Price.ReadOnly = True
        txtTarget_Weight.ReadOnly = True
        txtTarget_Width.ReadOnly = True
        txtOther_Requierment.ReadOnly = True
        cboCustomer.ReadOnly = True
        cboEnd_User.ReadOnly = True
        cboHanger.ReadOnly = True
        cboTest_Standard.ReadOnly = True
        cboQuality.ReadOnly = True
        txtColour.ReadOnly = True
        txtColour_no.ReadOnly = True
        txtCompositin.ReadOnly = True
        txtCuttable_Width.ReadOnly = True
        txtCuttable_Width.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtYardage.ReadOnly = True
        txtEnd_End.ReadOnly = True
        txtEnd_End.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRequierd_Date.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtReuest_Date.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        cboHanger.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        Call Load_Data()
    End Sub

    Function Load_Data()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim VALUE As Double

        'Load sales order to cboSO combobox

        Try
            Sql = "select * from T01_TEC_Development_Request inner join M01_TEC_Customer on T01Customer_Ref=M01Cus_RefNo where T01Req_No=" & strFab_Req_No & " and T01Status='DR' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With M01
                    txtReuest_Date.Text = Month(.Tables(0).Rows(0)("T01Req_Date")) & "/" & Microsoft.VisualBasic.Day(.Tables(0).Rows(0)("T01Req_Date")) & "/" & Year(.Tables(0).Rows(0)("T01Req_Date"))
                    txtRequierd_Date.Text = Month(.Tables(0).Rows(0)("T01Requied_Date")) & "/" & Microsoft.VisualBasic.Day(.Tables(0).Rows(0)("T01Requied_Date")) & "/" & Year(.Tables(0).Rows(0)("T01Requied_Date"))
                    cboCustomer.Text = .Tables(0).Rows(0)("M01Cus_Name")
                    txtCustomer_Referance.Text = .Tables(0).Rows(0)("T01Customer_Ref")
                    cboQuality.Text = .Tables(0).Rows(0)("T01Ref_Quality")
                    VALUE = .Tables(0).Rows(0)("T01Order_Qty")
                    txtOrder_Qty.Text = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtOrder_Qty.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))

                    VALUE = .Tables(0).Rows(0)("T01Target_Price")
                    txtTarger_Price.Text = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtTarger_Price.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))

                    cboEnd_User.Text = .Tables(0).Rows(0)("T01End_User")
                    txtFabric_Des.Text = .Tables(0).Rows(0)("T01Fabric_Description")
                    txtCompositin.Text = .Tables(0).Rows(0)("T01Composition")
                    VALUE = .Tables(0).Rows(0)("T01Target_Width")
                    txtTarget_Width.Text = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtTarget_Width.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))

                    VALUE = .Tables(0).Rows(0)("T01Target_Weight")
                    txtTarget_Weight.Text = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtTarget_Weight.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))

                    VALUE = .Tables(0).Rows(0)("T01End_End")
                    txtEnd_End.Text = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtEnd_End.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))

                    VALUE = .Tables(0).Rows(0)("T01Cutterble_Width")
                    txtCuttable_Width.Text = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtCuttable_Width.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))

                    txtColour.Text = .Tables(0).Rows(0)("T01Color")
                    txtColour_no.Text = .Tables(0).Rows(0)("T01Color_No")
                    cboTest_Standard.Text = .Tables(0).Rows(0)("T01Test_Std")
                    txtOther_Requierment.Text = .Tables(0).Rows(0)("T01Other_Req")
                    If Trim(.Tables(0).Rows(0)("T01Org_Ref")) = "1" Then
                        chkCounter_Sample.Checked = True
                    ElseIf Trim(.Tables(0).Rows(0)("T01Org_Ref")) = "2" Then
                        chkCustomer_Spec.Checked = True
                    End If

                    If Trim(.Tables(0).Rows(0)("T01Approval")) = "1" Then
                        chKApproval_Aesth.Checked = True
                    ElseIf Trim(.Tables(0).Rows(0)("T01Approval")) = "2" Then
                        chkApproval_Technical.Checked = True
                    End If

                    cboHanger.Text = .Tables(0).Rows(0)("T01Hangers")
                    VALUE = .Tables(0).Rows(0)("T01Yardage")
                    txtYardage.Text = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtYardage.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))

                    If Trim(.Tables(0).Rows(0)("T01Test_Rpt")) = "1" Then
                        chkTest_Internal.Checked = True
                    ElseIf Trim(.Tables(0).Rows(0)("T01Test_Rpt")) = "2" Then
                        chkTest_Outside.Checked = True
                    End If

                    If Trim(.Tables(0).Rows(0)("T01Brush")) = "1" Then
                        chkBrush_One.Checked = True
                    ElseIf Trim(.Tables(0).Rows(0)("T01Brush")) = "2" Then
                        chkBrush_Both.Checked = True
                    End If


                    If Trim(.Tables(0).Rows(0)("T01Anti_PIll")) = "1" Then
                        chkAntipill_One.Checked = True
                    ElseIf Trim(.Tables(0).Rows(0)("T01Anti_PIll")) = "2" Then
                        chkAntipill_Both.Checked = True
                    End If

                    If Trim(.Tables(0).Rows(0)("T01Sueded")) = "YES" Then
                        chkSueded.Checked = True

                    Else
                        chkSueded.Checked = False
                    End If

                    If Trim(.Tables(0).Rows(0)("T01Bio_Polish")) = "YES" Then
                        chkBio_Polish.Checked = True

                    Else
                        chkBio_Polish.Checked = False
                    End If
                End With
            End If
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try


    End Function
End Class