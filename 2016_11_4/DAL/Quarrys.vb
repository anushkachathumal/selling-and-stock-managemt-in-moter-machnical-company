
Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Quarrys

    Public Shared Function ValidateDistrict(ByVal strCode As String) As Boolean
        Dim SQL As String
        Dim con = New SqlConnection()
        Dim dsUser As DataSet
        con = DBEngin.GetConnection()
        ' Dim ms As New Complex
        Dim returnVariable As Boolean

        strDisname = ""
        ValidateDistrict = False
        Try
            SQL = "select * from M10District where M10DistrictNo='" & strCode & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If dsUser.Tables(0).Rows.Count > 0 Then
                ValidateDistrict = True
                'Return dsUser.Tables(0).Rows(0)("M10Name")
                strDisname = dsUser.Tables(0).Rows(0)("M10Name")

            Else
                MsgBox("Wrong District Code", MsgBoxStyle.Information, "INFORMATION ........")
                ValidateDistrict = False
            End If
            'Return returnVariable
        Catch returnMessage As Exception
            MessageBox.Show(returnMessage.Message)
        End Try

    End Function

    Public Shared Function GetSettmpGoodReceived(ByVal sType As String, ByVal sCode As String) As Boolean
        Dim dsUser As DataSet
        Dim con = New SqlConnection()
        GetSettmpGoodReceived = False
        '=======================================================================
        Try
            con = DBEngin.GetConnection()
            Select Case sType
                Case "LST"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_tmpGoodReceived", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", sCode))
                    If common.isValidDataset(dsUser) Then
                        GetSettmpGoodReceived = True
                    End If
                Case "ADD"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_tmpGoodReceived", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", sCode))
                    If common.isValidDataset(dsUser) Then
                        GetSettmpGoodReceived = False
                    Else
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_tmpGoodReceived", New SqlParameter("@cQryType", "ADD"), New SqlParameter("@vcFieldList", sCode))
                        GetSettmpGoodReceived = True
                    End If
                Case "DEL"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_tmpGoodReceived", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", sCode))
                    If common.isValidDataset(dsUser) = False Then
                        GetSettmpGoodReceived = False
                    Else
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_tmpGoodReceived", New SqlParameter("@cQryType", "DEL"), New SqlParameter("@vcCode", sCode))
                        GetSettmpGoodReceived = True
                    End If
            End Select
            '===================================================================
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
    Public Shared Function GetSetT17LocationStock(ByVal sType As String, ByVal sCode As String) As Boolean
        Dim dsUser As DataSet
        Dim con = New SqlConnection()
        GetSetT17LocationStock = False
        '=======================================================================
        Try
            con = DBEngin.GetConnection()
            Select Case sType
                Case "LST"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_GetSetT17LocationStock", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", sCode))
                    If common.isValidDataset(dsUser) Then
                        GetSetT17LocationStock = True
                    End If
                Case "ADD"
                    'dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_GetSetT17LocationStock", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", sCode))
                    'If common.isValidDataset(dsUser) Then
                    'GetSetT17LocationStock = False
                    'Else
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_GetSetT17LocationStock", New SqlParameter("@cQryType", "ADD"), New SqlParameter("@vcFieldList", sCode))
                    GetSetT17LocationStock = True
                    'End If
                Case "DEL"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_GetSetT17LocationStock", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", sCode))
                    If common.isValidDataset(dsUser) = False Then
                        GetSetT17LocationStock = False
                    Else
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_GetSetT17LocationStock", New SqlParameter("@cQryType", "DEL"), New SqlParameter("@vcCode", sCode))
                        GetSetT17LocationStock = True
                    End If
            End Select
            '===================================================================
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
End Class
