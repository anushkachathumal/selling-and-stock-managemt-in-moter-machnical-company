Imports System.Data
Imports System.Data.SqlClient


Public Class setServer_Date
    Dim strServer As String
    Public dtServerDateTime As DateTime

    Private Shared tmrLocalTimer As New System.Windows.Forms.Timer()

    ' Flag to check whether it is First time run. 

    Dim IsFirstTime As Boolean = True

    Public Sub TimerEventProcessor(ByVal myObject As Object, ByVal myEventArgs As EventArgs)

        'Check whether it is First time run. 

        ' If yes then Get Server time and display it in the Label.

        If IsFirstTime = True Then

            Dim cmd As SqlCommand

            Dim tmpConnection As SqlConnection

            Dim strConnectionString As String

            Try

                ' Read ConnectionString from Application Configuration file.

                strConnectionString = System.Configuration.ConfigurationManager.AppSettings("ServerConnectionString")

                tmpConnection = New SqlConnection(strConnectionString)

                ' Open the Connection

                tmpConnection.Open()

                ' Initialize the Command object with commandText and SQLConnection

                cmd = New SqlCommand("Select GetDate()", tmpConnection)

                'Execute the query and return the Server Date Time

                dtServerDateTime = cmd.ExecuteScalar

                ' Display the Server Date Time in the Label.

                strServer = Format(dtServerDateTime, "dd/MM/yyyy hh:mm:ss tt")

                ' Set the Flag to False

                IsFirstTime = False

            Catch ex As Exception

                MsgBox(ex.Message)

            Finally

                If tmpConnection.State = ConnectionState.Open Then

                    cmd = Nothing

                    tmpConnection.Close()

                End If

            End Try

        Else

            'Add one Second to the dtServerDateTime

            dtServerDateTime = DateAdd(DateInterval.Second, 1, dtServerDateTime)

        End If

        tmrLocalTimer.Enabled = True

        ' Display the Server Date Time in the ToolStripStatusLabel of the MainForm

        strServer = Format(dtServerDateTime, "dd/MM/yyyy hh:mm:ss tt")

    End Sub

    Public Sub GetServerDateTime()

        AddHandler tmrLocalTimer.Tick, AddressOf TimerEventProcessor

        ' Sets the timer interval to 1 second.

        tmrLocalTimer.Interval = 1000

        tmrLocalTimer.Start()

    End Sub

End Class