Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO

Public Class frmProjection_Upload
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
        c_dataCustomer1 = MakeDataTable_Projection()
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 110
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            '.DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(0).Columns(3).Width = 60
            '.DisplayLayout.Bands(0).Columns(5).Width = 60
            '.DisplayLayout.Bands(0).Columns(8).Width = 60
            '.DisplayLayout.Bands(0).Columns(7).Width = 70
            '.DisplayLayout.Bands(0).Columns(9).Width = 60

        End With
    End Function

    Function MakeDataTable_Projection() As DataTable
        Dim I As Integer
        Dim X As Integer
        Dim _Lastweek As Integer


        ' MsgBox(DatePart("ww", Today))
        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Planned(Y/N)", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Project Type", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("FG Supplier", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("OS Greige", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Production Step", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("With [Proj, PO]", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Retailer", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Business Unit", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Customer", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Quality", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Shade", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Sales Month", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("USD Mtr", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("USD Kg", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty(mtr)", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Product Month", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Year", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("CF", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable
    End Function

    Private Sub frmProjection_Upload_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
    End Sub

    Function Upload_File()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _Planned As String
        Dim _ProductType As String
        Dim _FG As String
        Dim _OS As String
        Dim _Production As String
        Dim _Width As String
        Dim _Retailler As String
        Dim _Biz_Unit As String
        Dim _Customer As String
        Dim _Quality As String
        Dim _Shade As String
        Dim _Sales_Month As String
        Dim _USD_M As Double
        Dim _USD_Kg As Double
        Dim _Qty As Double
        Dim _PMonth As String
        Dim _Year As String
        Dim _CF As Double
        


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
            
            strFileName = ConfigurationManager.AppSettings("FilePath") + "\Projection.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
               
                _Planned = Trim(fields(0))
                _ProductType = Trim(fields(1))
                _FG = Trim(fields(2))
                _OS = Trim(fields(3))
                characterToRemove = "'"

                _Production = Trim(fields(4))
                _Width = Trim(fields(5))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Width = (Replace(_Width, characterToRemove, ""))

                _Retailler = Trim(fields(6))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Retailler = (Replace(_Retailler, characterToRemove, ""))
                _Biz_Unit = Trim(fields(7))
               
                _Customer = Trim(fields(8))
                Dim _Len As Integer
                If Microsoft.VisualBasic.Left(Trim(fields(9)), 1) = "Q" Then
                    _Len = Microsoft.VisualBasic.Len(Trim(fields(9)))
                    _Quality = Microsoft.VisualBasic.Right(Trim(fields(9)), _Len - 1)
                Else
                    _Quality = Trim(fields(9))
                End If
                _Shade = Trim(fields(10))
                _Sales_Month = Trim(fields(11))
                _USD_M = Trim(fields(12))
                _USD_Kg = Trim(fields(13))
                _Qty = Trim(fields(14))
                _PMonth = Trim(fields(15))
                _Year = Trim(fields(16))
                _CF = Trim(fields(17))

                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("Planned(Y/N)") = _Planned
                newRow("Project Type") = _ProductType
                newRow("FG Supplier") = _FG
                newRow("OS Greige") = _OS
                newRow("Production Step") = _Production
                newRow("With [Proj, PO]") = _Width
                newRow("Retailer") = _Retailler
                newRow("Business Unit") = _Biz_Unit
                newRow("Customer") = _Customer
                newRow("Quality") = _Quality
                newRow("Shade") = _Shade
                newRow("Sales Month") = _Sales_Month
                newRow("USD Mtr") = _USD_M
                newRow("USD Kg") = _USD_Kg
                newRow("Qty(mtr)") = _Qty
                newRow("Year") = _Year
                newRow("CF") = _CF
                newRow("Product Month") = _PMonth
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

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Load_Gride()
        Call Upload_File()

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
        Dim _Planned As String
        Dim _ProductType As String
        Dim _FG As String
        Dim _OS As String
        Dim _Production As String
        Dim _Width As String
        Dim _Retailler As String
        Dim _Biz_Unit As String
        Dim _Customer As String
        Dim _Quality As String
        Dim _Shade As String
        Dim _Sales_Month As String
        Dim _USD_M As Double
        Dim _USD_Kg As Double
        Dim _Qty As Double
        Dim _PMonth As Integer
        Dim _Year As String
        Dim _CF As Double
        Dim characterToRemove As String
        Dim ncQryType As String

        Try
            vcWhere = "P01CODE='PRN'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "P01"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                P01Code = M01.Tables(0).Rows(0)("P01No")
            End If

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\Projection.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)

                _Planned = Trim(fields(0))
                _ProductType = Trim(fields(1))
                _FG = Trim(fields(2))
                _OS = Trim(fields(3))
                characterToRemove = "'"

                _Production = Trim(fields(4))
                _Width = Trim(fields(5))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Width = (Replace(_Width, characterToRemove, ""))

                _Retailler = Trim(fields(6))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Retailler = (Replace(_Retailler, characterToRemove, ""))
                _Biz_Unit = Trim(fields(7))

                _Customer = Trim(fields(8))
                Dim _Len As Integer
                If Microsoft.VisualBasic.Left(Trim(fields(9)), 1) = "Q" Then
                    _Len = Microsoft.VisualBasic.Len(Trim(fields(9)))
                    _Quality = Microsoft.VisualBasic.Right(Trim(fields(9)), _Len - 1)
                Else
                    _Quality = Trim(fields(9))
                End If

                _Shade = Trim(fields(10))
                _Sales_Month = Trim(fields(11))
                _USD_M = Trim(fields(12))
                _USD_Kg = Trim(fields(13))
                _Qty = Trim(fields(14))
                _PMonth = Trim(fields(15))
                _Year = Trim(fields(16))
                If Trim(fields(17)) <> "" Then
                    _CF = Trim(fields(17))
                Else
                    _CF = 0
                End If

                vcWhere = ""
                ncQryType = "PRJ"
                nvcFieldList1 = "(M43Planned," & "M43Product_type," & "M43FG_Supplier," & "M43OS," & "M43Production_Step," & "M43Project_PO," & "M43Retailler," & "M43B_Unit," & "M43Customer," & "M43Quality," & "M43Shade," & "M43Sales_Month," & "M43USD_Mtr," & "M43USD_kg," & "M43Sales_Volume," & "M43Product_Month," & "M43Year," & "M43CF," & "M43Count_No," & "M43User) " & "values('" & _Planned & "','" & _ProductType & "','" & _FG & "','" & _OS & "','" & _Production & "','" & _Width & "','" & _Retailler & "','" & _Biz_Unit & "','" & _Customer & "','" & _Quality & "','" & _Shade & "','" & _Sales_Month & "','" & _USD_M & "','" & _USD_Kg & "','" & _Qty & "'," & _PMonth & "," & _Year & ",'" & _CF & "'," & P01Code & ",'" & strDisname & "')"
                up_GetSetProjection(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next

            nvcFieldList1 = "Update P01PARAMETER set P01No=P01No + " & 1 & " where P01CODE='PRN'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

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

    Private Sub UltraButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton6.Click
        Call Update_Records()
    End Sub

    Private Sub UltraGrid1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid1.InitializeLayout

    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Load_Gride()
    End Sub
End Class