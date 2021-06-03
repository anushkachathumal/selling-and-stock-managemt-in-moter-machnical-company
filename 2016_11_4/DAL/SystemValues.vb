Imports System.Threading
Imports System.Data.SqlClient
'Imports Microsoft.ApplicationBlocks.Data

Public Class SystemValues

    Private Shared _SystemValues As SystemValues
    Private Shared _testValue As String

    Public Shared DatabaseServerName As String
    Public Shared DatabaseName As String
    Public Shared DatabaseUserID As String
    Public Shared DatabaseUserPassword As String
    Public Shared ConnectionString As String
    Public Shared DateFormat As String = DateFormats.DD_MM_YYYY
    Public Shared DateFormatString As String
    Public Shared CompanyCode As String
    Public Shared UserGroup As String
    Public Shared UserID As String
    Public Shared ComputerName As String
    Public Shared ComputerIPAddress As String
    Public Shared InstalledPath As String
    Public Shared UserLoginDateTime As String
    Public Shared DefaultCurrencyCode As String = "AED"
    Public Shared debugMode As debugModes
    Public Shared showConfugurationCheck As Boolean = False
    Public Shared X As Long
    Public Shared Y As Long
    Public Shared sessionNumber As Long
    Public Shared showCostFlag As String
    Public Shared CurrentOutletCode As String = ""
    Public Shared IsSQLDatabase As String
    Public Shared DSNName As String
    Public Shared DSNUserID As String
    Public Shared DSNPassword As String
    Public Shared PackageType As String
    Public Shared TimeFormatString As String
    Public Shared DatabasePath As String
    Public Shared POSPath As String
    Public Shared itemCodeStructureType As itemCodeStructureTypes
    Public Shared serverFolderPath As String
    Public Shared styleLength As Integer
    Public Shared sizeLength As Integer
    Public Shared MainDbConnectionString As String = ""

    Public Shared colourLength As Integer

    Public Shared dsUserRecord As DataSet
    Public Shared drCompanyProfile As DataRow
    Shared dsStockRoomDetailRecords As DataSet

    'Public Shared lastServerDateTime As Date
    Public Shared serverTimeDifference As Long

    Private Shared _mu As New Mutex


    Public Enum itemCodeStructureTypes

        Style_Only = 0
        Style_And_Size_Only = 1
        Style_And_Colour_Only = 2
        Style_Size_Colour = 3
        Style_Colour_Size = 4
        Unknown_ItemCode_Structure_Type


    End Enum

    Public Enum serverDateTimeFormats

        SystemValues_Date_Format = 0
        dd_MM_yyyy_hh_mm_ss = 1
        dd_MMM_yyyy_hh_mm_ss = 2
        MM_dd_yyyy_hh_mm_ss = 3
        MMM_dd_yyyy_hh_mm_ss = 4

    End Enum
    Public Enum debugModes

        None = 0
        Development = 1
        UserDebug = 2

    End Enum


    Public Enum DateFormats
        None = 0
        MM_DD_YYYY = 1
        DD_MM_YYYY = 2
        MMM_DD_YYYY = 3
        DD_MMM_YYYY = 4
    End Enum

    Public Shared Sub loadGlobalData()

        '   loadStockRoomDetailRecords()
        '   loadCompanyProfile()

    End Sub


    ''Public Shared Function loadCompanyProfile() As String
    ''    Dim strSQL As String
    ''    Dim dsCompanyProfile As DataSet

    ''    Try
    ''        strSQL = " SELECT * FROM CompanyProfile WHERE cCompanyNo = '" & CompanyCode & "'"

    ''        dsCompanyProfile = SqlHelper.ExecuteDataset(SystemValues.ConnectionString, CommandType.Text, strSQL)
    ''        drCompanyProfile = dsCompanyProfile.Tables(0).Rows(0)


    ''        dsCompanyProfile = Nothing

    ''    Catch ex As Exception
    ''        Throw ex
    ''    End Try
    ''End Function


    'Public Shared Function loadStockRoomDetailRecords() As String
    '    Dim strSQL As String

    '    Try
    '        strSQL = "SELECT * FROM CompanyProfileDetail WHERE cCompanyNo = '" & CompanyCode & "'"

    '        dsStockRoomDetailRecords = SqlHelper.ExecuteDataset(SystemValues.ConnectionString, CommandType.Text, strSQL)

    '        If dsStockRoomDetailRecords.Tables(0).Rows.Count > 0 Then
    '            Return "NOERROR"
    '        Else
    '            Return " No records found for Stockroom details"
    '        End If

    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function

    Public Overloads Shared Function getStockRoomDetailRecord(ByVal StockRoomCode As String, ByRef drStockRoomDetailRecord As DataRow) As String
        'Dim strSQL As String
        'Dim SQLString As String
        'Dim stockroomSQL As String

        Dim tableKeys(1) As DataColumn
        Dim keyValues(1) As Object

        Try
            tableKeys(0) = dsStockRoomDetailRecords.Tables(0).Columns("cCompanyNo")
            tableKeys(1) = dsStockRoomDetailRecords.Tables(0).Columns("cStockroomCode")

            dsStockRoomDetailRecords.Tables(0).PrimaryKey = tableKeys

            keyValues(0) = CompanyCode
            keyValues(1) = StockRoomCode

            drStockRoomDetailRecord = dsStockRoomDetailRecords.Tables(0).Rows.Find(keyValues)

            Return "NOERROR"



            'If dsStockRoomDetailRecord.Tables.Count <= 0 Then
            '    dsStockRoomDetailRecord.Tables.Add("Table1")
            'End If

            '    SQLString = "cStockroomCode = '" & StockRoomCode & "'"

            'Dim rows() As DataRow

            ' rows = dsStockRoomDetailRecords.Tables(0).Select(SQLString)

            'For Each dtRow As DataRow In rows
            '    dsStockRoomDetailRecord.Tables(0).ImportRow(dtRow)
            'Next


        Catch ex As Exception
            Throw ex
        End Try

    End Function


    Public Shared Function GetInstance() As SystemValues

        _mu.WaitOne()

        Try
            If _SystemValues Is Nothing Then
                _SystemValues = New SystemValues
                '// load defaults

            End If

        Finally
            _mu.ReleaseMutex()
        End Try

        Return _SystemValues

    End Function


    Public Overloads Shared Function getStockRoomDetailRecord(ByRef dsStochRoomDetailRecord As DataSet, _
                                              ByVal ParamArray stockrooms As String()) As String
        ' Dim strSQL As String

        Try

            Dim SQLString As String
            Dim stockroomSQL As String


            '        /Create your new datatable
            'DataTable dt = new DataTable();
            '//Clone the old DataTable structure to the new one
            'dt = dsEmployees1.Tables[0].Clone();
            '//Select a specific rows from the old DataTable
            'DataRow[] dr = dsEmployees1.Tables[0].Select("EmployeeID > 5");
            '//loop on the selected rows and import it to the new datatable using the ImportRow method
            'foreach (DataRow drEmployee in(dr))
            '{
            'dt.ImportRow(drEmployee);
            '}


            Try

                If dsStochRoomDetailRecord.Tables.Count <= 0 Then
                    dsStochRoomDetailRecord.Tables.Add("Table1")
                End If

                SQLString = "cStockroomCode IN ("

                For Each stockroom As String In stockrooms
                    stockroomSQL &= "'" & stockroom & "', "
                Next

                If stockroomSQL.EndsWith(", ") Then
                    stockroomSQL = stockroomSQL.Substring(0, InStrRev(stockroomSQL, ", ") - 1)
                End If

                SQLString &= stockroomSQL & ")"

                Dim rows() As DataRow

                rows = dsStockRoomDetailRecords.Tables(0).Select(SQLString)

                For Each dtRow As DataRow In rows
                    dsStochRoomDetailRecord.Tables(0).ImportRow(dtRow)
                Next

                Return "NOERROR"

            Catch ex As Exception
                Return ex.Message
            End Try

            Return "NOERROR"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    'Public Shared Function getText() As String
    '    Return _testValue
    'End Function


    'Public Shared Function setText(ByVal val As String) As String
    '    _testValue = val
    'End Function

    'Public Shared Function setCompanyCode(ByVal val As String) As String
    '    strCompanyCode = val
    'End Function


    'Public Shared Function getCompanyCode() As String

    '    _mu.WaitOne()

    '    Try
    '        If _singleton Is Nothing Then
    '            _singleton = New SystemValues
    '        End If

    '    Finally
    '        _mu.ReleaseMutex()
    '    End Try

    '    Return _singleton

    'End Function

    'Public Shared Sub readServerDateTime()
    '    Dim strSQL As String
    '    Dim dtServerDateTime As DateTime

    '    Try
    '        strSQL = "SELECT GETDATE()"

    '        dtServerDateTime = CType(SqlHelper.ExecuteScalar(SystemValues.ConnectionString, CommandType.Text, strSQL), DateTime)
    '        'lastServerDateTime = dtServerDateTime
    '        serverTimeDifference = DateDiff(DateInterval.Second, Now, dtServerDateTime)

    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    Public Shared Function ServerDateTime(Optional ByVal serverDateTimeFormat As serverDateTimeFormats = serverDateTimeFormats.SystemValues_Date_Format) As DateTime

        Try
            If serverDateTimeFormat = serverDateTimeFormats.SystemValues_Date_Format Then
                Select Case SystemValues.DateFormat
                    Case DateFormats.DD_MM_YYYY
                        serverDateTimeFormat = serverDateTimeFormats.dd_MM_yyyy_hh_mm_ss
                    Case DateFormats.DD_MMM_YYYY
                        serverDateTimeFormat = serverDateTimeFormats.dd_MMM_yyyy_hh_mm_ss
                    Case DateFormats.MM_DD_YYYY
                        serverDateTimeFormat = serverDateTimeFormats.MM_dd_yyyy_hh_mm_ss
                    Case DateFormats.MMM_DD_YYYY
                        serverDateTimeFormat = serverDateTimeFormats.MMM_dd_yyyy_hh_mm_ss
                    Case DateFormats.None
                        serverDateTimeFormat = serverDateTimeFormats.dd_MMM_yyyy_hh_mm_ss
                End Select
            End If

            ServerDateTime = DateAdd(DateInterval.Second, serverTimeDifference, Now)
            Select Case serverDateTimeFormat
                Case serverDateTimeFormats.dd_MM_yyyy_hh_mm_ss
                    ServerDateTime = Format(ServerDateTime, "dd MM yyyy HH:mm:ss")
                Case serverDateTimeFormats.dd_MMM_yyyy_hh_mm_ss
                    ServerDateTime = Format(ServerDateTime, "dd MMM yyyy HH:mm:ss")
                Case serverDateTimeFormats.MM_dd_yyyy_hh_mm_ss
                    ServerDateTime = Format(ServerDateTime, "MMM dd yyyy HH:mm:ss")
                Case serverDateTimeFormats.MMM_dd_yyyy_hh_mm_ss
                    ServerDateTime = Format(ServerDateTime, "MMM dd yyyy HH:mm:ss")
            End Select

            Return ServerDateTime

        Catch ex As Exception
            Throw ex
        End Try
    End Function

End Class



