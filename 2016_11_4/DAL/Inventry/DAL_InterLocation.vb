Public Class DAL_InterLocation
    Public Function MakeDataTable() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Current Wight", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("New Wight", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a End column
        colWork = New DataColumn("Current Qty", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("New Qty", GetType(String))
        dataTable.Columns.Add(colWork)
        Return dataTable

    End Function

    Public Function MakeDataTable_PO() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Rate", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Total", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTableMK_Return() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        colWork = New DataColumn("Cost Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Retail Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("Rec.Qty", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        'colWork = New DataColumn("Free Issue", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("Ex Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Total", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("##", GetType(Boolean))
        ''colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True
        Return dataTable

    End Function


    Public Function MakeDataTable_SalesTR() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        colWork = New DataColumn("Cost Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Retail Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Mtr", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Total", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function
    Public Function MakeDataTable_GRN_CANCEL_INVOICE() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Ref_No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn
        colWork = New DataColumn("Supplier Name", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        ' Keys(0) = colWork
        '  colWork.ReadOnly = True
        '  dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Voucher No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("GRN No", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Net Amount", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Chq No", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Bank Name", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Canceld by", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Cancel Time", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function


    Public Function MakeDataTable_GRN_PAYID_INVOICE() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Ref_No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn
        colWork = New DataColumn("Supplier Name", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        ' Keys(0) = colWork
        '  colWork.ReadOnly = True
        '  dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Voucher No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("GRN No", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Net Amount", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Chq No", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Bank Name", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function


    Public Function MakeDataTable_GRN_PAYID_INVOICE_SUMMERY() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Ref_No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn
        colWork = New DataColumn("Supplier Name", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        ' Keys(0) = colWork
        '  colWork.ReadOnly = True
        '  dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Voucher No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        'colWork = New DataColumn("GRN No", GetType(String))
        ''  colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("Net Amount", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Chq No", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Bank Name", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function


    Public Function MakeDataTable_GRN_PAY() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        '  colWork.ReadOnly = True
        '  dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Supplier Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("GRN No", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Com.Invoice", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Net Amount", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Discount", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("MK Return", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("VAT Amount", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("NBT", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Paid Amount", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_oUTSTANDING_PAY() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        '  colWork.ReadOnly = True
        '  dataTable.PrimaryKey = Keys
        ' create and add a Description column

        colWork = New DataColumn("#Type", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Customer Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Inv.Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Inv.No", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Inv.Amount", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Paid Amount", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Balance", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

       

        Return dataTable

    End Function

    Public Function MakeDataTable_Creditors_Statment() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        '  colWork.ReadOnly = True
        '  dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Ref No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Supplier Name", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Cr", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Dr", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Chq No", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

     

        Return dataTable

    End Function

    
    Public Function MakeDataTable_Bank_PAY() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Voucher No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        '  colWork.ReadOnly = True
        '  dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Acc_Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Acc_Name", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Amount", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

      

        Return dataTable

    End Function

    Public Function MakeDataTable_GRN_Payment1() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Supplier Name", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        '  dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("GRN No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column


        colWork = New DataColumn("Com.Invoice", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Paid Amount", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True



        Return dataTable

    End Function
    Public Function MakeDataTable_BankPay() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Pay To", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Amount", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True



        Return dataTable

    End Function

    Public Function MakeDataTable_Prize() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Cost Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        colWork = New DataColumn("Retail Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


       
        Return dataTable

    End Function

    Public Function MakeDataTableCom_Discount() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Invo.Amount", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Discount(%)", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        Return dataTable

    End Function

    Public Function MakeDataTableCus_Discount() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Chq Days", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Discount(%)", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        Return dataTable

    End Function

    Public Function MakeDataTable_BOM() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable

    End Function

    Public Function MakeDataTable_Stock_Adjestment_1() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Ref.Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Part No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("#Rack No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("#Cell No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)

        Return dataTable

    End Function
    Public Function MakeDataTable_itemIssue_Uniq() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("#Ref.No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("#Part No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        colWork = New DataColumn("Retail Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Discount%", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Free Issue", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        

        colWork = New DataColumn("Total", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

      
        Return dataTable

    End Function

    Public Function MakeDataTable_Ez_Uniq() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("#Card Name", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Card No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        colWork = New DataColumn("Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        Return dataTable

    End Function

    Public Function MakeDataTable_ChqPay_Uniq() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("#Bank Name", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Chq No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        colWork = New DataColumn("DOR", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("Chq Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        Return dataTable

    End Function

    Public Function MakeDataTable_iNVOICE_UNIQ() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Department", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("#Part No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Retail Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Discount%", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Free Issue", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True



        colWork = New DataColumn("Total", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable

    End Function

    Public Function MakeDataTableGRN() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("#Ref.No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn
        colWork = New DataColumn("#Part No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Description column
        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        colWork = New DataColumn("Cost Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Retail Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#Rack No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Cell No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Total", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("##", GetType(Boolean))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True
        Return dataTable

    End Function
    Public Function MakeDataTablePayment() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Date", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Payment", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("W.Drawals", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function
    Public Function MakeDataTable_AccStatment() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Date", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Ref.No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Credit", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Debit", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Due Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Bank Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable

    End Function

    Public Function MakeDataTableCard() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Card No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Card Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        colWork = New DataColumn("Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("Qty", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        ''colWork = New DataColumn("Rec.Qty", GetType(String))
        ''colWork.MaxLength = 250
        ''dataTable.Columns.Add(colWork)
        ''colWork.ReadOnly = True

        ''colWork = New DataColumn("Free Issue", GetType(String))
        ''colWork.MaxLength = 250
        ''dataTable.Columns.Add(colWork)
        ''colWork.ReadOnly = True

        'colWork = New DataColumn("Total", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTableSales() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        colWork = New DataColumn("Rate", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Dis%", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("Free Issue", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("Total", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function
    Public Function MakeDataTable_StockIN() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("Remark", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_ExStock() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        colWork = New DataColumn("Ex Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Current Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("New Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        Return dataTable

    End Function


    Public Function MakeDataTableItem() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Loc Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Current Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        Return dataTable

    End Function
    Public Function MakeDataTable_Unpacking() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("Recycle Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' colWork.ReadOnly = True
        colWork = New DataColumn("Remark", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)


        Return dataTable

    End Function

    Public Function MakeDataTable_StockItem() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Current Stock", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("New Stock", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_ProductionItemUnpackingD() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Unpacking No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Set Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Recycle Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function


    Public Function MakeDataTable_ProductionItemUnpacking() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Unpacking No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_ProductionItem() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Product Category", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Item Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Retail Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_StockItemsrpt2() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Category", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Supplier", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Cost", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Retail", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Total Cost", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Total Retail", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_Outstanding_detailes() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Customer Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Customer Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("##", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Invoice Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Invoice No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Invoice Value", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Paid Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Balance Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function
    Public Function MakeDataTable_StockItemsrpt1() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Category", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Supplier", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Cost Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Retail Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_StockMovement() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Date", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Remark", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Sales Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("GRN Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Mkt Return Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Transfer", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Wastage", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Balance", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function


    Public Function MakeDataTable_GRN_rptD() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Date", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("GRN No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Com Invoice", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("PO#", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Supplier Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Cost Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Retail Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Free Issue", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Total", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable

    End Function

    Public Function MakeDataTable_Wastage_App() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Wastage No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Net Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("User", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
       
        Return dataTable

    End Function

    Public Function MakeDataTable_GRN_rpt() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Date", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("GRN No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Com Invoice", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("PO#", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Supplier Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Net Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Discount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("VAT Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("NBT", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Gross Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function
    Public Function MakeDataTable_Wastage() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Wastage No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Item Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Rate", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("Free Issue", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("Total", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_Transfer() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Transfer No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("Com Invoice", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("To Location", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Rate", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("Free Issue", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("Total", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_Varions() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Date(L/Up)", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Item Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

       
        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Cost Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Last Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("New Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Variance", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Variance Value", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_MkReturn() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Return No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("Com Invoice", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("Supplier", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Rate", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("Free Issue", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("Total", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_Sales() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Date", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("#Invoice Type", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Job No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Invoice No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Vehicle No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Customer Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Invoice Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_Sales_Detailes() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Date", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("#Invoice Type", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Job No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Invoice No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Part No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Part Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("#Rate", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#Discount%", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#Total Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_Sales_Detailes_1() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Date", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys


        colWork = New DataColumn("Job No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Invoice No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Department", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True



        colWork = New DataColumn("#Rate", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_Sales_Summery() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Date", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys




        colWork = New DataColumn("#Net Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#Cash", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#Credit", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("#Chq", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("#Card", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        Return dataTable

    End Function

    Public Function MakeDataTable_Sales_Summery1() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Date", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys




        colWork = New DataColumn("#Net Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#Cash", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#Credit", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("#Chq", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("#Card", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        Return dataTable

    End Function

    Public Function MakeDataTable_CR_PAY_Invo() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Invo No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Pay Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable

    End Function
    Public Function MakeDataTable_Inv() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Invo Date", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Invoice No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Tobe Paid", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("Balance", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True



        Return dataTable

    End Function


    Public Function MakeDataTable_Outstanding_Age() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Customer Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Customer Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Inv.Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Inv.No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Inv.Value", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Paid Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Balance", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("1-30", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("30-45", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("45-60", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("Grater than 60", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        Return dataTable

    End Function

    Public Function MakeDataTable_Outstanding() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Customer Name", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Credit Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("Paid Amount", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        'colWork = New DataColumn("Balance", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True



        Return dataTable

    End Function
    Public Function MakeDataTable_CR_RECEIVED() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        '  colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Customer Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("Com Invoice", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("Customer Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Invo Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Invoice No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Invo Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Paid Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Balance Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("Total", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_CR_PAY() As DataTable

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
        colWork = New DataColumn("Pay.Doc", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("Com Invoice", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("Pay Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("Customer Code", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("Customer Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Pay Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Cash", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Cheque", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("Total", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_Customer_Summery() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Invoice No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        'colWork = New DataColumn("Date", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Customer Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Invo.Value", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

      

        Return dataTable

    End Function


    Public Function MakeDataTable_Cashier_Summery() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Date", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        'colWork = New DataColumn("Date", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("Net Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Cash Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("VISA", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Master", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Amex", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("User", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable

    End Function

    Public Function MakeDataTable_Daily_Summery() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Date", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        'colWork = New DataColumn("Date", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("Net Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Cash Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Credit Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("VISA", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Master", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Amex", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("User", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True


        Return dataTable

    End Function

    Public Function MakeDataTable_Cashier() As DataTable

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
        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Invoice No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("MC No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Net Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Cash", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Credit", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("VISA", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Master", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("AMEX", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("User", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        
        Return dataTable

    End Function

    Public Function MakeDataTable_Purchasing() As DataTable

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
        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Com Invoice", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Supplier", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Cost Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Free Issue", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Total", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_StockItemsrpt() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Category", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Item Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#Unit Cost", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#Unit Rate", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#Total Cost", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#Total Rate", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_StockItemsrpt_Brand() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Category", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        ' create and add a Description column
        colWork = New DataColumn("Brand Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        ' create and add a Description column
        colWork = New DataColumn("Item Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#Unit Cost", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#Unit Rate", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#Total Cost", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#Total Rate", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function


    Public Function MakeDataTable_rptPackingBox() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Box No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Box Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        
        Return dataTable

    End Function

    Public Function MakeDataTable_rptPackingBox_Item() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Box No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Box Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Product Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        Return dataTable

    End Function
    

    Public Function MakeDataTable_Receipt() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Ref No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Invoice No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Amount", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        

        Return dataTable

    End Function

    Public Function MakeDataTable_rptStockIN() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Batch No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Remark", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function
    Public Function MakeDataTable_rptCustomer() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Cus No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Customer Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Address", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Contact No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable

    End Function

    Public Function MakeDataTable_rptStockIN_Summery() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable

    End Function

    Public Function MakeDataTable_ProductionSet() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Product Category", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Item Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Item Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Cost Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("Retail Price", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable

    End Function

    Public Function MakeDataTable_Sales_Order() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        ' colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Line Item", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Quality", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Quantity", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Matching With", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)


        colWork = New DataColumn("Retailer", GetType(String))
        colWork.MaxLength = 120
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("1st Bulk/Repeat", GetType(String))
        colWork.MaxLength = 20
        dataTable.Columns.Add(colWork)
        '  colWork.ReadOnly = True

        colWork = New DataColumn("Lab Dye", GetType(String))
        colWork.MaxLength = 10
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True



        colWork = New DataColumn("Possible App Date", GetType(Date))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' colWork.ReadOnly = True

        colWork = New DataColumn("NPL", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' colWork.ReadOnly = True

        colWork = New DataColumn("NPL App Date", GetType(Date))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("PP", GetType(String))
        dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("PP App Date", GetType(Date))
        '  colWork.MaxLength = 20
        dataTable.Columns.Add(colWork)
        ' colWork.ReadOnly = True
        colWork = New DataColumn("Cap Tb Tkn", GetType(String))
        '  colWork.MaxLength = 20
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("GO Rolling", GetType(String))
        '  colWork.MaxLength = 20
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("GO Non Rolling", GetType(String))
        '  colWork.MaxLength = 20
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("STRO Repelenish", GetType(String))
        '  colWork.MaxLength = 20
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("STRO Reduce", GetType(String))
        '  colWork.MaxLength = 20
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Lead Time", GetType(String))
        '  colWork.MaxLength = 20
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Comment", GetType(String))
        '  colWork.MaxLength = 20
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Req.Del.Date", GetType(Date))
        '  colWork.MaxLength = 20
        dataTable.Columns.Add(colWork)

        Return dataTable

    End Function

    Public Function MakeDataTableCheck_Dep() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = False
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Department", GetType(String))
        colWork.MaxLength = 350
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        Return dataTable

    End Function

    Public Function MakeDataTable_SetCreation() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Product Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Product Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Qty", GetType(Integer))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

       

        Return dataTable

    End Function

    Public Function MakeDataTable_ShadeNew() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = False
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Shade", GetType(String))
        colWork.MaxLength = 350
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        Return dataTable

    End Function

    Public Function MakeDataTable_ShadeWithBiz() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = False
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Shade", GetType(String))
        colWork.MaxLength = 350
        dataTable.Columns.Add(colWork)
        ' colWork.ReadOnly = True
        ' create and add a Qty column

        colWork = New DataColumn("##1", GetType(Boolean))
        ' colWork.MaxLength = 350
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Biz Unit", GetType(String))
        colWork.MaxLength = 350
        dataTable.Columns.Add(colWork)

        Return dataTable

    End Function

    Public Function MakeDataTable_ProMonthWithBiz() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = False
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Biz Unit", GetType(String))
        colWork.MaxLength = 350
        dataTable.Columns.Add(colWork)
        ' colWork.ReadOnly = True
        ' create and add a Qty column

        colWork = New DataColumn("##1", GetType(Boolean))
        ' colWork.MaxLength = 350
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Product Month", GetType(String))
        colWork.MaxLength = 350
        dataTable.Columns.Add(colWork)

        Return dataTable

    End Function

    Public Function MakeDataTable_BizUnit() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = False
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Biz Unit", GetType(String))
        colWork.MaxLength = 350
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        Return dataTable

    End Function

    Public Function MakeDataTable_StockCode() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("10 Class", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Stock Code", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("GRN Date", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a End column
        colWork = New DataColumn("Qty", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Location", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Shade", GetType(String))
        dataTable.Columns.Add(colWork)
        Return dataTable

    End Function

    Public Function MakeDataTableCheck_Customer() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = False
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Customer Name", GetType(String))
        colWork.MaxLength = 350
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        Return dataTable

    End Function

    Public Function MakeDataTableCheck_Retailer() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = False
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Retailer Name", GetType(String))
        colWork.MaxLength = 350
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        Return dataTable

    End Function

    Public Function MakeDataTableCheck_Merch() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = False
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Merchant", GetType(String))
        colWork.MaxLength = 350
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        Return dataTable

    End Function

    Public Function MakeDataTableCheck_Location() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = False
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Location", GetType(String))
        colWork.MaxLength = 350
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        Return dataTable

    End Function

    Public Function MakeDataTableCheck_BU() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = False
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("BU", GetType(String))
        colWork.MaxLength = 350
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        Return dataTable

    End Function

    Public Function MakeDataTable_PendingTEC() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Req No", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        ' colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("##", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Quality No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Fabric Type", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Shade", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)


        colWork = New DataColumn("Customer Name", GetType(String))
        colWork.MaxLength = 120
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True



        Return dataTable

    End Function

    Public Function MakeDataTable_DevelopingTEC() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Req No", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        ' colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Start Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Developer", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Quality No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Development", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Knitting", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)


        colWork = New DataColumn("Dyeing", GetType(String))
        colWork.MaxLength = 120
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True



        Return dataTable

    End Function

    Public Function MakeDataTable_Dye_YarnRequest() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Yarn Name", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        ' colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Year", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Week", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Capacity", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

       
        Return dataTable

    End Function

    Public Function MakeDataTable_ProjectionMonth() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Projection", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Used", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("##", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

       


        Return dataTable

    End Function

    Public Function MakeDataTable_ProjectionMonth1() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column





        Return dataTable

    End Function

    Public Function MakeDataTable_FG() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Location", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        'Dim Keys(0) As DataColumn

        'Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        '' create and add a Description column
     

        


        colWork = New DataColumn("Material", GetType(String))
        colWork.MaxLength = 120
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Material Description", GetType(String))
        colWork.MaxLength = 220
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Retailer", GetType(String))
        colWork.MaxLength = 20
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Merchant", GetType(String))
        colWork.MaxLength = 20
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Sales Order", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Line Item", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Roll No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty(Mtr)", GetType(Double))
        ' colWork.MaxLength = 10
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty(Kg)", GetType(Double))
        'colWork.MaxLength = 10
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Ageing", GetType(String))
        '  colWork.MaxLength = 20
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("GRN Date", GetType(String))
        '  colWork.MaxLength = 20
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("New Update", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' colWork.ReadOnly = True

       
        colWork = New DataColumn("latest update", GetType(String))
        '  colWork.MaxLength = 20
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
       
        colWork = New DataColumn("Dedline Date", GetType(Date))
        '  colWork.MaxLength = 20
        dataTable.Columns.Add(colWork)

        Return dataTable

    End Function

    Public Function MakeDataTableEral() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Lot No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Machine No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Programe No", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("Programe Name", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a End column
        colWork = New DataColumn("Lot Type", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Standed Time", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Start Date", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Start Time", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("End Date", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("End Time", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Total Hour", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Quality", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Quality Group", GetType(String))
        dataTable.Columns.Add(colWork)

   

        colWork = New DataColumn("Shade Code", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Shade", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Shade Type", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Weight", GetType(String))
        dataTable.Columns.Add(colWork)

        Return dataTable

    End Function

    Public Function MakeDataTableGrige_Stock() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("20Class", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Grige Order No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("L/Item", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Stock Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Age", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Available Qty(Kg)", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty use for Order", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

        colWork = New DataColumn("Del.Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

        colWork = New DataColumn("##", GetType(Boolean))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)

        'colWork = New DataColumn("Balance to Knit", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

    

        Return dataTable

    End Function

    Public Function MakeDataTableGrige_Request_M_Maneger() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Ref No", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Requested Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Required Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Customer Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Order Qty (mtr)", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Merchant", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTableGrige_Request_Merch() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Ref No", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Requested Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Required Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Customer Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Order Qty (mtr)", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        

        Return dataTable

    End Function

    Public Function MakeDataTableGrige_TECYARN() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("No of Ply ", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Yarn Count", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Yarn Type", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Yarn Material", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Loop Lenth", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Course Length", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("No of Feeders", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

        colWork = New DataColumn("No of Needles", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

        colWork = New DataColumn("Contribution", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False


        Return dataTable

    End Function
    Public Function MakeDataTableGrige_ProcessRoot() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Code ", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys

        colWork = New DataColumn("Process Type", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable

    End Function
    Public Function MakeDataTableGrige_Fabric() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Code ", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable

    End Function

    Public Function MakeDataTableGrige_YARNMASTER() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Yarn Code ", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable

    End Function

    Public Function MakeDataTableGrige_Process() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Process Root", GetType(String))
        dataTable.Columns.Add(colWork)
        '  colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        'colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("##", GetType(Boolean))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        '  colWork.ReadOnly = True

       


        Return dataTable

    End Function


    Public Function MakeDataTableGrige_Process1() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Process No", GetType(String))
        dataTable.Columns.Add(colWork)
        '  colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        'colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("##", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        '  colWork.ReadOnly = True




        Return dataTable

    End Function
    Public Function MakeDataTableGrige_KNTSpec_Yarn() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Yarn Type", GetType(String))
        dataTable.Columns.Add(colWork)
        '  colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        'colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Yarn Count", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        '  colWork.ReadOnly = True

        colWork = New DataColumn("Yarn%", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Stich Lenth", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Course Lenth", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("No of Feeders", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        Return dataTable

    End Function

    Public Function MakeDataTableGrige_CompositionNew() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Descrition", GetType(String))
        dataTable.Columns.Add(colWork)
        '  colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        'colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Rate", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        '  colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("No of Measure", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Stage", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        Return dataTable

    End Function

    Public Function MakeDataTableGrige_Screen2() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("SPECIAL", GetType(String))
        dataTable.Columns.Add(colWork)
        ' colWork.ReadOnly = True
        ' '' add CustomerID column to key array and bind to DataTable
        'Dim Keys(0) As DataColumn

        'Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("STRUCTURE", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' colWork.ReadOnly = True

        colWork = New DataColumn("FIBER COMPOSITION", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' colWork.ReadOnly = True


        Return dataTable

    End Function

    Public Function MakeDataTableGrige_Composition() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Type", GetType(String))
        dataTable.Columns.Add(colWork)
        '  colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        'Dim Keys(0) As DataColumn

        'Keys(0) = colWork
        ''colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("%", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        '  colWork.ReadOnly = True




        Return dataTable

    End Function

    Public Function MakeDataTableConf_Stock() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Stock Code", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("##", GetType(Boolean))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)

        'colWork = New DataColumn("Balance to Knit", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True



        Return dataTable

    End Function

    Public Function MakeDataTablePreVious_Stock() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Stock Code", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Week No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Year", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Used Qty(Kg)", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

       


        Return dataTable

    End Function

    Public Function MakeDataTableYarn_Booking() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("10Class", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Stock Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Yarn Aging", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Location", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Available Qty(Kg)", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty use for Order", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

        colWork = New DataColumn("##", GetType(Boolean))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True



        Return dataTable

    End Function


    Public Function MakeDataTableKnitting_PLN() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Machine No", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Start Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("End Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Allc.Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        '    colWork.ReadOnly = True

        colWork = New DataColumn("No of Hour", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Status", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("##", GetType(Boolean))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)

        Return dataTable

    End Function

    Public Function MakeDataTable_Delivary() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Week No", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTable_Delivary_Daily() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Date", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTableDye_Plan() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("15Class", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("No of Batch", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

        colWork = New DataColumn("Allocate Yarn", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

        colWork = New DataColumn("Allocate Con", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

        colWork = New DataColumn("Allocate Dye MC", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

        


        Return dataTable

    End Function

    Public Function MakeDataTableKnt_Projection() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Shade", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        'colWork = New DataColumn("Month", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        'colWork = New DataColumn("Year", GetType(String))
        'colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        colWork = New DataColumn("Last Proj", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Projection", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Allocated", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

        colWork = New DataColumn("For Quality", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True





        Return dataTable

    End Function


    Public Function MakeDataTableYarn_Dyeing() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("15Class", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Shade", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Composition", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Dyed Yarn Req.Knit", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Dyed Yarn Stock", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

        colWork = New DataColumn("Yarn Req - Dyed Yarn", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        

        colWork = New DataColumn("Balance Qty for Allocate", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

        colWork = New DataColumn("No of Cons", GetType(String))
        ' colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

        colWork = New DataColumn("Allocate Yarn", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False



        Return dataTable

    End Function

    Public Function MakeDataTableYarn() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column

        colWork = New DataColumn("10 Class", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Stock Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Age", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

        colWork = New DataColumn("Allocation", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

        colWork = New DataColumn("Log User", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable

    End Function

    Public Function MakeDataTableYarn_Stock() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column

        colWork = New DataColumn("Stock Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Alocate Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = False

       
        colWork = New DataColumn("Log User", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable

    End Function

    Public Function MakeDataTableQuality() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Quality No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Quantity", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
       

        Return dataTable

    End Function

    Public Function MakeDataTableDNH() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Batch No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Lot No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column

        ' create and add a End column
        colWork = New DataColumn("Sub No", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Machine", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Quality", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Shade", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Dyed Quantity (Kg)", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Reject Qty", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Batch type", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Dye House Shade", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("M/C Change", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        colWork = New DataColumn("S/C Change", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        colWork = New DataColumn("Process Change", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        colWork = New DataColumn("Liquor ratio Change", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        colWork = New DataColumn("Dye Lot Change", GetType(Boolean))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Reason", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork = New DataColumn("Bulk Shaid Status", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork = New DataColumn("Wet on Wet comment", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork = New DataColumn("QC & Exam comments", GetType(String))
        dataTable.Columns.Add(colWork)

        Return dataTable

    End Function

    Public Function MakeDataTableMB51() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Material No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Plant", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Stock Loc", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("Movement Type", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a End column
        colWork = New DataColumn("No of Mat Doc", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Posting Date", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Quantity", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Unit", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Batch No", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("PO No", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("DOADE", GetType(String))
        dataTable.Columns.Add(colWork)

        
        Return dataTable

    End Function

    Public Function MakeDataTableZDCA_PURCHASE() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Material No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Plant", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Stock Loc", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("Movement Type", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a End column
        colWork = New DataColumn("No of Mat Doc", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Posting Date", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Quantity", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Unit", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Batch No", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("PO No", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("DOADE", GetType(String))
        dataTable.Columns.Add(colWork)


        Return dataTable

    End Function
    Public Function MakeDataTableCheck_Invoice() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("PI No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Supplier Invoice No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("L/J Ref No", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("Fabric Referance", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Shipment Date", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Quantity", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Amount $", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Payment Terms", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Payment Confermation", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Maneger Confermation", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True




        Return dataTable

    End Function

    Public Function MakeDataTableCheck_InvoiceSJSR() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("LJSR No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Supplier Invoice No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("L/J Ref No", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("Fabric Referance", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Shipment Date", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Quantity", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Amount $", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Payment Terms", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Payment Due Mill", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Payment Due Customer", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        colWork = New DataColumn("Maneger Confermation", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True




        Return dataTable

    End Function



    Public Function MakeDataTablePlaning() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Color Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Color Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("L/D Ref No", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("Quantity", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
      
        colWork = New DataColumn("ETD", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable

    End Function

    Public Function MakeDataTableColour() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Color Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Color Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("ETD Mill", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("Remarks", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True


        Return dataTable

    End Function

    Public Function MakeDataTableWIP() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Batch No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Customer", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Material", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("Description", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a End column

        colWork = New DataColumn("Delivary Date", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Last Confirmation Date", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty (Kg)", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Next Oparation", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Planing Comment", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork = New DataColumn("Order Type", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork = New DataColumn("Sales Order", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Line Item", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork = New DataColumn("Qty (Mtr)", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork = New DataColumn("Merchant", GetType(String))
        dataTable.Columns.Add(colWork)

        Return dataTable

    End Function


    Public Function MakeDataTable_Wip(ByVal strYDate As String, ByVal strTodayDate As String) As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Sales Order", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Line Item", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Material", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("Description", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a End column

        colWork = New DataColumn("Delivary Date", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Batch No", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Batch Qty (Kg)", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Batch Qty (Mtr)", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("No.Of.Dys In S/L", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Next Operation", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("NC Comment", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn(strYDate, GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn(strTodayDate, GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Complete Date", GetType(Date))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Possible Devivary Date", GetType(Date))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Reason", GetType(String))
        dataTable.Columns.Add(colWork)


        colWork = New DataColumn("LIB dep", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Week", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Customer", GetType(String))
        dataTable.Columns.Add(colWork)

        Return dataTable

    End Function

    Public Function MakeDataTableAdvance() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Dyed Date", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Machine", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Batch no", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("Quality", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a End column

        colWork = New DataColumn("Type", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Programme Name", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Shade code", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Shade", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty (Kg)", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Shade Type", GetType(String))
        dataTable.Columns.Add(colWork)
       
        Return dataTable

    End Function
    Public Function MakeDataTableSuppAcc() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Acc No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Payment Terms", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Intermediary Bank", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Address", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("Advising Bank", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a End column
        colWork = New DataColumn("Address1", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Shift Code", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Telex", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Tel", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Fax", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        Return dataTable

    End Function


    Public Function MakeDataTableMRS() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("SAP-Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Prc-$/Kg", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("PkSiz", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a End column
        colWork = New DataColumn("MR", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Month", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Value", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("12 MAvg", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Hgt 6 MAvg", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Hgt 3 MAvg", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("CQ-MRS", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Purchase", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("NLTW", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("NLTD", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("SS", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("$SS", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("RL", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("$RL", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("RQ", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("$RQ", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("ML", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("End Stock", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("stk-Holding", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Category", GetType(String))
        dataTable.Columns.Add(colWork)


        Return dataTable

    End Function

    Public Function MakeDataTableAlarm() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Category", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("SAP Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("SS", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("End Stock", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a End column
        colWork = New DataColumn("Avg 12", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("M6", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("N3", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("CQ-MRS", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Last Month Qty", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Last 30Day Qty", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("SD", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("2nd Last Week", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Last Week", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Com Last 30Day", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Com SD", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Com 2ndLast Week", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Com Last Week", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Com Last 14Days", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("RL", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("RQ", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("LT Days", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Outstanding P/O", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("#PO", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("PID", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("PDD", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Reg No", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("ETA", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("ETD", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("L14D - Con", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTableFDP() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("SAP Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Week", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("Qty", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
       

        Return dataTable

    End Function

    


    Public Function MakeDataTableEral1() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Ref.Doc", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        colWork.ReadOnly = True
        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Lot No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        '  colWork.ReadOnly = True

        colWork = New DataColumn("Machine No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Programe No", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("Programe Type", GetType(String))
        dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True
        ' create and add a End column
        colWork = New DataColumn("Lot Type", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Standed Time", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Start Date", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Start Time", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("End Date", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("End Time", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Total Hour", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Quality", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Quality Group", GetType(String))
        dataTable.Columns.Add(colWork)



        colWork = New DataColumn("Shade Code", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Shade", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Shade Type", GetType(String))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Weight", GetType(String))
        dataTable.Columns.Add(colWork)

        Return dataTable

    End Function


    Public Function MakeDataTableBTT() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Date", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        ' colWork.ReadOnly = True
        ' dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Invoice No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Amount", GetType(String))
        dataTable.Columns.Add(colWork)
        ' colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("Status", GetType(Boolean))
        dataTable.Columns.Add(colWork)
        '  colWork.ReadOnly = True
        ' create and add a End column
        'colWork = New DataColumn("End No", GetType(Integer))
        'dataTable.Columns.Add(colWork)


        Return dataTable

    End Function
    Public Function MakeDataTablePurchasing() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn
        colWork.ReadOnly = True
        Keys(0) = colWork

        ' dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Quntity", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        'colWork = New DataColumn("Description", GetType(String))
        'dataTable.Columns.Add(colWork)
        '' create and add a End column
        ''colWork = New DataColumn("End No", GetType(Integer))
        ''dataTable.Columns.Add(colWork)
        'colWork = New DataColumn("Color Code", GetType(String))
        'dataTable.Columns.Add(colWork)
        'colWork = New DataColumn("Color", GetType(String))
        'dataTable.Columns.Add(colWork)
        'colWork = New DataColumn("Amount", GetType(String))
        'dataTable.Columns.Add(colWork)
        Return dataTable

    End Function


    

    Public Function MakeDataTableSpairpart() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        colWork.ReadOnly = True
        Dim Keys(0) As DataColumn

        Keys(0) = colWork

        ' dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Wight(grm)", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("Quntity", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        'colWork = New DataColumn("Exp Date", GetType(Date))
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True
        colWork = New DataColumn("Price", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a End column
        colWork = New DataColumn("Discount", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Total", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Remark", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        Return dataTable

    End Function
    Public Function MakeDataTableT09PawnFlutter() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        colWork.ReadOnly = True
        Dim Keys(0) As DataColumn

        Keys(0) = colWork

        ' dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Wight(grm)", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        colWork = New DataColumn("Quntity", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        colWork = New DataColumn("Amount", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a End column
        colWork = New DataColumn("Interest Rate", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Profit", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Total", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Remark", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        Return dataTable

    End Function
    Public Function MakeDataTablePurchasorder() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        colWork.ReadOnly = True
        Dim Keys(0) As DataColumn

        Keys(0) = colWork

        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Date of Receving", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Start column
        'colWork = New DataColumn("Exp Date", GetType(Date))
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True
        colWork = New DataColumn("Quntity", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a End column
        'colWork = New DataColumn("Discount", GetType(String))
        ''colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        'colWork = New DataColumn("Total", GetType(String))
        ''colWork.MaxLength = 250
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True
        Return dataTable

    End Function
    Public Function MakeDataTableOther() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        colWork.ReadOnly = True
        Dim Keys(0) As DataColumn

        Keys(0) = colWork

        dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        'colWork = New DataColumn("Gurantee No", GetType(Integer))
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True
        ' create and add a Start column
        'colWork = New DataColumn("Exp Date", GetType(Date))
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        'colWork = New DataColumn("Amount", GetType(String))
        'dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True
        '' create and add a End column
        ''colWork = New DataColumn("End No", GetType(Integer))
        ''dataTable.Columns.Add(colWork)


        Return dataTable

    End Function

    Public Function MakeDataTable_Journal() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Acc Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        'colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Acc Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True
        ' create and add a Qty column
        colWork = New DataColumn("Debit", GetType(String))
        '  colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Credit", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        'colWork.ReadOnly = True

        Return dataTable

    End Function

    Public Function MakeDataTableCollector() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Customer Name", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork

        ' dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Invoice No", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' create and add a Qty column
        'colWork = New DataColumn("Quntity", GetType(Integer))
        'dataTable.Columns.Add(colWork)
        '' create and add a Start column
        'colWork = New DataColumn("Starting No", GetType(Integer))
        'dataTable.Columns.Add(colWork)

        'colWork = New DataColumn("Ending No", GetType(Integer))
        'dataTable.Columns.Add(colWork)
        '' create and add a End column
        ''colWork = New DataColumn("End No", GetType(Integer))
        ''dataTable.Columns.Add(colWork)
        'colWork = New DataColumn("Agent Code", GetType(String))
        'dataTable.Columns.Add(colWork)

        Return dataTable

    End Function

    Public Function MakeDataTableFundReceving() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Mode", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork

        ' dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' create and add a Qty column
        colWork = New DataColumn("Referance", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' create and add a Start column
        colWork = New DataColumn("Value", GetType(String))
        dataTable.Columns.Add(colWork)
        ' create and add a End column
        colWork = New DataColumn("Type", GetType(String))
        dataTable.Columns.Add(colWork)


        Return dataTable

    End Function

    Public Function MakeDataTableViewBookissue() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork

        ' dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' create and add a Qty column
        colWork = New DataColumn("Date", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' create and add a Start column
        colWork = New DataColumn("Start No", GetType(Integer))
        dataTable.Columns.Add(colWork)
        ' create and add a End column
        colWork = New DataColumn("End No", GetType(Integer))
        dataTable.Columns.Add(colWork)

        colWork = New DataColumn("Qty", GetType(Integer))
        dataTable.Columns.Add(colWork)

        Return dataTable

    End Function

    Public Function MakeDataTableBookBalance() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Code", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork

        ' dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Name", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' create and add a Qty column
        colWork = New DataColumn("Bal/Book", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' create and add a Start column
        colWork = New DataColumn("Reqd/Amount", GetType(String))
        dataTable.Columns.Add(colWork)
        ' create and add a End column
        colWork = New DataColumn("Avail/Balance", GetType(String))
        dataTable.Columns.Add(colWork)


        Return dataTable

    End Function

    Public Function MakeDataTableGoodReceived() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Game Type", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork

        ' dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' create and add a Qty column
        colWork = New DataColumn("Start Number", GetType(String))
        'colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' create and add a Start column
        colWork = New DataColumn("End Number", GetType(String))
        dataTable.Columns.Add(colWork)
        ' create and add a End column
        colWork = New DataColumn("Qty", GetType(String))
        dataTable.Columns.Add(colWork)


        Return dataTable

    End Function

    Public Function MakeDataTableWinners() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("Winners")
        Dim colWork As New DataColumn("Draw Date", GetType(String))
        dataTable.Columns.Add(colWork)
        Dim Keys(0) As DataColumn
        Keys(0) = colWork
        colWork = New DataColumn("Ticket Number", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork = New DataColumn("Prize Code", GetType(String))
        dataTable.Columns.Add(colWork)
        colWork = New DataColumn("Agent No", GetType(String))
        dataTable.Columns.Add(colWork)
        Return dataTable


    End Function

    Public Function MakeDataTableWinnersinfo() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Serial of VIRN No", GetType(String))
        dataTable.Columns.Add(colWork)

        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        dataTable.PrimaryKey = Keys
        ' dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Prize Code", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' create and add a Qty column
        colWork = New DataColumn("Price", GetType(String))
        dataTable.Columns.Add(colWork)
        ' create and add a Start column
        colWork = New DataColumn("Draw Date/Agent", GetType(String))
        dataTable.Columns.Add(colWork)
        ' create and add a End column
        ' colWork = New DataColumn("End No", GetType(Integer))
        'dataTable.Columns.Add(colWork)


        Return dataTable

    End Function
    Public Function MakeDataTableWinnersinfoTICK() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Price", GetType(String))
        dataTable.Columns.Add(colWork)

        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        dataTable.PrimaryKey = Keys
        ' dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("No of Tickets", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' create and add a Qty column
        'colWork = New DataColumn("Price", GetType(String))
        'dataTable.Columns.Add(colWork)
        ' create and add a Start column
        'colWork = New DataColumn("Draw Date/Agent", GetType(String))
        'dataTable.Columns.Add(colWork)
        ' create and add a End column
        ' colWork = New DataColumn("End No", GetType(Integer))
        'dataTable.Columns.Add(colWork)


        Return dataTable

    End Function


    Public Function MakeDataTableEralUpload() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Price", GetType(String))
        dataTable.Columns.Add(colWork)

        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn

        Keys(0) = colWork
        dataTable.PrimaryKey = Keys
        ' dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("No of Tickets", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' create and add a Qty column
        'colWork = New DataColumn("Price", GetType(String))
        'dataTable.Columns.Add(colWork)
        ' create and add a Start column
        'colWork = New DataColumn("Draw Date/Agent", GetType(String))
        'dataTable.Columns.Add(colWork)
        ' create and add a End column
        ' colWork = New DataColumn("End No", GetType(Integer))
        'dataTable.Columns.Add(colWork)


        Return dataTable

    End Function
End Class
