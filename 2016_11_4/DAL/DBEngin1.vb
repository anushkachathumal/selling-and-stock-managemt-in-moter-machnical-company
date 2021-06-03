Imports System
Imports System.Collections
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Public Class DBEngin1
    Public Shared dTransActionDate As Date

    Public Shared Function GetConnectionString() As String
        'Return ConfigurationSettings.AppSettings(SystemValues.CompanyCode)
        Return ConfigurationManager.AppSettings("CD1")

    End Function

    Public Shared Function GetConnection() As SqlConnection

        Return New SqlConnection(GetConnectionString())

    End Function


    Public Shared Function GetConnection(ByVal open As Boolean) As SqlConnection
        Dim connection As New SqlConnection(GetConnectionString())
        If (open) Then
            connection.Open()
        End If
        Return connection
    End Function


    Public Shared Function CloseConnection(ByRef connection As SqlConnection) As Boolean
        If connection.State = ConnectionState.Open Then
            connection.Close()
            Return True
        Else
            ' connection.Close()
            Return False
        End If
    End Function


    Public Shared Sub ExecuteNonQuery(ByVal connection As SqlConnection, ByVal transaction As SqlTransaction, ByVal commandText As String)
        Dim state As ConnectionState = connection.State
        Dim command As New SqlCommand

        ' Initialize the command
        command.Connection = connection
        command.CommandText = commandText
        command.CommandType = CommandType.StoredProcedure
        command.Transaction = transaction
        command.CommandTimeout = 300
        ' Open the database connection if it isn't already opened
        If state = ConnectionState.Closed Then
            connection.Open()
        End If

        Try
            ' Execute the query
            command.ExecuteNonQuery()
        Finally
            ' If the database connection was closed before the method call, close it again
            If state = ConnectionState.Closed Then
                connection.Close()
            End If
        End Try
    End Sub
    Public Shared Sub ExecuteNonQueryText(ByVal connection As SqlConnection, ByVal transaction As SqlTransaction, ByVal commandText As String)
        Dim state As ConnectionState = connection.State
        Dim command As New SqlCommand

        ' Initialize the command
        command.Connection = connection
        command.CommandText = commandText
        command.CommandType = CommandType.Text
        command.Transaction = transaction
        command.CommandTimeout = 300
        ' Open the database connection if it isn't already opened
        If state = ConnectionState.Closed Then
            connection.Open()
        End If

        Try
            ' Execute the query
            command.ExecuteNonQuery()
        Finally
            ' If the database connection was closed before the method call, close it again
            If state = ConnectionState.Closed Then
                connection.Close()

            End If
        End Try
    End Sub

    Public Shared Sub ExecuteNonQuery(ByVal connection As SqlConnection, ByVal transaction As SqlTransaction, ByVal commandText As String, ByVal ParamArray parameters As SqlParameter())
        Dim state As ConnectionState = connection.State
        Dim command As New SqlCommand

        ' Initialize the command
        command.Connection = connection
        command.CommandText = commandText
        command.CommandType = CommandType.StoredProcedure
        command.Transaction = transaction

        ' Append each parameter to the command
        Dim parameter As SqlParameter
        For Each parameter In parameters
            command.Parameters.Add(parameter)
        Next

        ' Open the database connection if it isn't already opened
        If state = ConnectionState.Closed Then
            connection.Open()
        End If

        Try
            ' Execute the query
            command.ExecuteNonQuery()
        Finally
            ' If the database connection was closed before the method call, close it again
            If state = ConnectionState.Closed Then
                connection.Close()
            End If
        End Try
    End Sub


    Public Shared Function ExecuteReader(ByVal connection As SqlConnection, ByVal transaction As SqlTransaction, ByVal commandText As String) As SqlDataReader
        Dim state As ConnectionState = connection.State
        Dim command As New SqlCommand

        ' Initialize the command
        command.Connection = connection
        command.CommandText = commandText
        command.CommandType = CommandType.StoredProcedure
        command.Transaction = transaction

        ' Open the database connection if it isn't already opened
        If state = ConnectionState.Closed Then
            connection.Open()
            Return command.ExecuteReader(CommandBehavior.CloseConnection)
        Else
            Return command.ExecuteReader()
        End If
    End Function
    Public Shared Function ExecuteReaderText(ByVal connection As SqlConnection, ByVal transaction As SqlTransaction, ByVal commandText As String) As SqlDataReader
        Dim state As ConnectionState = connection.State
        Dim command As New SqlCommand

        ' Initialize the command
        command.Connection = connection
        command.CommandText = commandText
        command.CommandType = CommandType.Text
        command.Transaction = transaction

        ' Open the database connection if it isn't already opened
        If state = ConnectionState.Closed Then
            connection.Open()
            Return command.ExecuteReader(CommandBehavior.CloseConnection)
        Else
            Return command.ExecuteReader()
        End If
    End Function


    Public Shared Function ExecuteReader(ByVal connection As SqlConnection, ByVal transaction As SqlTransaction, ByVal commandText As String, ByVal ParamArray parameters As SqlParameter()) As SqlDataReader
        Dim state As ConnectionState = connection.State
        Dim command As New SqlCommand

        ' Initialize the command
        command.Connection = connection
        command.CommandText = commandText
        command.CommandType = CommandType.StoredProcedure
        command.Transaction = transaction

        ' Append each parameter to the command
        Dim parameter As SqlParameter
        For Each parameter In parameters
            command.Parameters.Add(parameter)
        Next

        ' Open the database connection if it isn't already opened
        If state = ConnectionState.Closed Then
            connection.Open()
            Return command.ExecuteReader(CommandBehavior.CloseConnection)
        Else
            Return command.ExecuteReader()
        End If
    End Function


    Public Shared Function ExecuteScalar(ByVal connection As SqlConnection, ByVal transaction As SqlTransaction, ByVal commandText As String) As Object
        Dim state As ConnectionState = connection.State
        Dim command As New SqlCommand

        ' Initialize the command
        command.Connection = connection
        command.CommandText = commandText
        command.CommandType = CommandType.StoredProcedure
        command.Transaction = transaction
        command.CommandTimeout = 300
        ' Open the database connection if it isn't already opened
        If state = ConnectionState.Closed Then
            connection.Open()
        End If

        Try
            ' Execute the query
            Return command.ExecuteScalar()
        Finally
            ' If the database connection was closed before the method call, close it again
            If state = ConnectionState.Closed Then
                connection.Close()
            End If
        End Try
    End Function


    Public Shared Function ExecuteScalar(ByVal connection As SqlConnection, ByVal transaction As SqlTransaction, ByVal commandText As String, ByVal ParamArray parameters As SqlParameter()) As Object
        Dim state As ConnectionState = connection.State
        Dim command As New SqlCommand

        ' Initialize the command
        command.Connection = connection
        command.CommandText = commandText
        command.CommandType = CommandType.StoredProcedure
        command.Transaction = transaction
        command.CommandTimeout = 300
        ' Append each parameter to the command
        Dim parameter As SqlParameter
        For Each parameter In parameters
            command.Parameters.Add(parameter)
        Next

        ' Open the database connection if it isn't already opened
        If state = ConnectionState.Closed Then
            connection.Open()
        End If

        Try
            ' Execute the query
            Return command.ExecuteScalar()
        Finally
            ' If the database connection was closed before the method call, close it again
            If state = ConnectionState.Closed Then
                connection.Close()
            End If
        End Try
    End Function

    Public Shared Function ExecuteDataset(ByVal connection As SqlConnection, ByVal transaction As SqlTransaction, ByVal commandText As String) As DataSet
        Dim state As ConnectionState = connection.State
        Dim command As New SqlCommand
        Dim adapter As New SqlDataAdapter
        Dim tableDataSet As New DataSet

        ' Initialize the command
        command.Connection = connection
        command.CommandText = commandText
        command.CommandType = CommandType.Text
        command.Transaction = transaction
        command.CommandTimeout = 300
        ' Open the database connection if it isn't already opened
        If state = ConnectionState.Closed Then
            connection.Open()
        End If

        adapter.SelectCommand = command

        Try
            adapter.Fill(tableDataSet)
            Return tableDataSet

        Finally
            ' If the database connection was closed before the method call, close it again
            If state = ConnectionState.Closed Then

                connection.Close()


            End If
        End Try
    End Function

    Public Shared Function ExecuteDataset(ByVal connection As SqlConnection, ByVal transaction As SqlTransaction, ByVal commandText As String, ByVal ParamArray parameters As SqlParameter()) As DataSet
        Dim state As ConnectionState = connection.State
        Dim command As New SqlCommand
        Dim adapter As New SqlDataAdapter
        Dim tableDataSet As New DataSet

        ' Initialize the command
        command.Connection = connection
        command.CommandText = commandText
        command.CommandType = CommandType.StoredProcedure
        command.Transaction = transaction
        'command.CommandTimeout = 300
        ' Append each parameter to the command
        Dim parameter As SqlParameter
        For Each parameter In parameters
            command.Parameters.Add(parameter)
        Next

        ' Open the database connection if it isn't already opened
        If state = ConnectionState.Closed Then
            connection.Open()
        End If

        Try
            ' Execute the query
            adapter.SelectCommand = command
            adapter.Fill(tableDataSet)
            Return tableDataSet
        Finally
            ' If the database connection was closed before the method call, close it again
            If state = ConnectionState.Closed Then
                connection.Close()
            End If
        End Try
    End Function
End Class
