Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO
Public Class frmUpload_Capacity
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Dim c_dataCustomer2 As System.Data.DataTable

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = MakeDataTable_Capacity()
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '   .DisplayLayout.Bands(0).Columns(10).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(11).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(12).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(13).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left


            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(5).Width = 120
            .DisplayLayout.Bands(0).Columns(6).Width = 90
            .DisplayLayout.Bands(0).Columns(7).Width = 90
            .DisplayLayout.Bands(0).Columns(8).Width = 90
            .DisplayLayout.Bands(0).Columns(9).Width = 90


            '.DisplayLayout.Bands(0).Columns(3).Width = 60
            '.DisplayLayout.Bands(0).Columns(5).Width = 60
            '.DisplayLayout.Bands(0).Columns(8).Width = 60
            '.DisplayLayout.Bands(0).Columns(7).Width = 70
            '.DisplayLayout.Bands(0).Columns(9).Width = 60

        End With
    End Function

    Function MakeDataTable_Capacity() As DataTable
        Dim I As Integer
        Dim X As Integer
        Dim _Lastweek As Integer


        ' MsgBox(DatePart("ww", Today))
        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Ref No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Dye M/C Group", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Quality", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Trim Quality", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Body Quality", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Compisition", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Width Cm", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Gram per sqm (g/m2)", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("KG/M", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Meter/kg", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("SR_ (sensitive CLR)", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("SR_ (normal CLR)", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("SR_ (White CLR)", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("DR__ (sensitive CLR)", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("DR__ (normal CLR)", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("DR__ (White CLR)", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Changed date", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Comments", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

      
        Return dataTable
    End Function

    Private Sub frmUpload_Capacity_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Load_Gride()
        Call Upload_File()
    End Sub

    Function Upload_File()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _RefNo As String
        Dim _MC_Group As String
        Dim _Qulity As String
        Dim _Trim As String
        Dim _Body As String
        Dim _Compisition As String
        Dim _Width As String
        Dim _Gram As String
        Dim _Kg_M As String

        Dim _Mtr_Kg As String
        Dim _sensitive As String
        Dim _normal As String
        Dim _White As String
        Dim _sensitiveC As String
        Dim _normalC As String
        Dim _WhiteC As String
        Dim _Comments As String
        Dim _Change_Date As String


        'Dim nvcFieldList1 As String

        'Dim connection As SqlClient.SqlConnection
        'Dim transaction As SqlClient.SqlTransaction
        'Dim transactionCreated As Boolean
        'Dim connectionCreated As Boolean

        'connection = DBEngin.GetConnection(True)
        'connectionCreated = True
        'transaction = connection.BeginTransaction()
        'transactionCreated = True
        Dim M01 As DataSet
        Dim I As Integer
        Dim A As String
        Dim characterToRemove As String

        Dim X11 As Integer
        Dim _TollPLS As Integer
        Dim _TollMIN As Integer

        Try

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\Capacity.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)

                _RefNo = Trim(fields(0))
                _MC_Group = Trim(fields(1))
                _Qulity = Trim(fields(2))
                _Trim = Trim(fields(3))
                characterToRemove = "'"

                _Body = Trim(fields(4))
                _Compisition = Trim(fields(5))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Compisition = (Replace(_Compisition, characterToRemove, ""))

                _Width = Trim(fields(6))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Gram = Trim(fields(7))

                _Kg_M = Trim(fields(8))
                _Mtr_Kg = Trim(fields(9))
                _sensitive = Trim(fields(10))
                _normal = Trim(fields(11))
                _White = Trim(fields(12))
                _sensitiveC = Trim(fields(13))
                _normalC = Trim(fields(14))
                _WhiteC = Trim(fields(15))
                _Change_Date = Trim(fields(16))
                _Comments = Trim(fields(17))

                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("Ref No") = _RefNo
                newRow("Dye M/C Group") = _MC_Group
                newRow("Quality") = _Qulity
                newRow("Trim Quality") = _Trim
                newRow("Body Quality") = _Body
                newRow("Compisition") = _Compisition
                newRow("Width Cm") = _Width
                newRow("Gram per sqm (g/m2)") = _Gram
                newRow("KG/M") = _Kg_M
                newRow("Meter/kg") = _Mtr_Kg
                newRow("SR_ (sensitive CLR)") = _sensitive
                newRow("SR_ (normal CLR)") = _normal
                newRow("SR_ (White CLR)") = _White
                newRow("DR__ (sensitive CLR)") = _sensitiveC
                newRow("DR__ (normal CLR)") = _normalC
                newRow("DR__ (White CLR)") = _WhiteC
                newRow("Changed date") = _Change_Date
                newRow("Comments") = _Comments
                c_dataCustomer1.Rows.Add(newRow)

                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            'MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            'transaction.Commit()
            'DBEngin.CloseConnection(connection)
            'connection.ConnectionString = ""



        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            'DBEngin.CloseConnection(connection)
            'connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            'MsgBox(X11)
            'DBEngin.CloseConnection(connection)
            'connection.ConnectionString = ""
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Private Sub UltraButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton6.Click
        Call Update_Records()
    End Sub



    Function Update_Records()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        Dim M01 As DataSet
        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim P01Code As Integer
        Dim vcWhere As String
        Dim X11 As Integer
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _RefNo As String
        Dim _MC_Group As String
        Dim _Qulity As String
        Dim _Trim As String
        Dim _Body As String
        Dim _Compisition As String
        Dim _Width As String
        Dim _Gram As String
        Dim _Kg_M As String

        Dim _Mtr_Kg As String
        Dim _sensitive As String
        Dim _normal As String
        Dim _White As String
        Dim _sensitiveC As String
        Dim _normalC As String
        Dim _WhiteC As String
        Dim _Comments As String
        Dim _Change_Date As String
        Dim characterToRemove As String


        Dim ncQryType As String

        Try
            'vcWhere = "P01CODE='PRN'"
            'M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "P01"), New SqlParameter("@vcWhereClause1", vcWhere))
            'If isValidDataset(M01) Then
            '    P01Code = M01.Tables(0).Rows(0)("P01No")
            'End If

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\Capacity.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)

                _RefNo = Trim(fields(0))
                _MC_Group = Trim(fields(1))
                _Qulity = Trim(fields(2))
                _Trim = Trim(fields(3))
                characterToRemove = "'"

                _Body = Trim(fields(4))
                _Compisition = Trim(fields(5))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Compisition = (Replace(_Compisition, characterToRemove, ""))

                _Width = Trim(fields(6))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Gram = Trim(fields(7))

                _Kg_M = Trim(fields(8))
                _Mtr_Kg = Trim(fields(9))
                _sensitive = Trim(fields(10))
                _normal = Trim(fields(11))
                _White = Trim(fields(12))
                _sensitiveC = Trim(fields(13))
                _normalC = Trim(fields(14))
                _WhiteC = Trim(fields(15))
                _Change_Date = Trim(fields(16))
                _Comments = Trim(fields(17))

                nvcFieldList1 = "M49Ref=" & _RefNo & ""
                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "CGLI"), New SqlParameter("@vcWhereClause1", nvcFieldList1))
                If isValidDataset(M01) Then

                Else
                    vcWhere = ""
                    ncQryType = "ICPL"
                    nvcFieldList1 = "(M49Ref," & "M49MC_Group," & "M49Quality," & "M49Trim," & "M49Base," & "M49Composition," & "M49Width_Cm," & "M49Gram," & "M49KG_Mtr," & "M49Mtr_Kg," & "M49SR_Nomal," & "M49SR_Critical," & "M49SR_White," & "M49DR_Nomal," & "M49DR_Critical," & "M49DR_White," & "M49Change_Date," & "M49Comment) " & "values('" & _RefNo & "','" & _MC_Group & "','" & _Qulity & "','" & _Trim & "','" & _Body & "','" & _Compisition & "','" & _Width & "','" & _Gram & "','" & _Kg_M & "','" & _Mtr_Kg & "','" & _sensitive & "','" & _normal & "','" & _White & "','" & _sensitiveC & "','" & _normalC & "','" & _WhiteC & "','" & _Change_Date & "','" & _Comments & "')"
                    up_GetSetCAPACITY(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                End If
                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next

            'nvcFieldList1 = "Update P01PARAMETER set P01No=P01No + " & 1 & " where P01CODE='PRN'"
            'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            MsgBox("File Update succesfully", MsgBoxStyle.Information, "Information ...")
            transaction.Commit()
            connection.Close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function
End Class