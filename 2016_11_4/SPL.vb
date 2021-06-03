Imports System.Configuration
Imports System.Net.NetworkInformation
Imports System.Net
Imports DBLotVbnet.common
Imports System
Imports System.Data
Imports System.Data.SqlClient

Public Class SPL
    Dim networkcard() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces()
    
    Public Timer()

    Private Sub SPL_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        'On Error Resume Next
        

        'UltraListView1.Items.Add(1, "GFHGFH")
        Dim win As System.Security.Principal.WindowsIdentity

        Dim LocalHostName As String

        win = System.Security.Principal.WindowsIdentity.GetCurrent()
        LoggedUser = win.Name.Substring(win.Name.IndexOf("\") + 1)

        Try
            
            
            UltraListView1.Items.Add(1, "Machine Name : " & (System.Environment.MachineName))
            VMName = System.Environment.MachineName
            UltraListView1.Items.Add(2, "Machine ID : " & networkcard(0).GetPhysicalAddress.ToString())

            netCard = networkcard(0).GetPhysicalAddress.ToString

            If My.Computer.Info.OSFullName.Contains("XP") Then
                UltraListView1.Items.Add(3, "Machine IP : " & System.Net.Dns.GetHostAddresses(My.Computer.Name)(0).ToString)
                VIP = System.Net.Dns.GetHostAddresses(My.Computer.Name)(0).ToString
            ElseIf My.Computer.Info.OSFullName.Contains("2003") Then
                LocalHostName = Dns.GetHostName()
                Dim ipEnter As IPHostEntry = Dns.GetHostEntry(LocalHostName)
                Dim IpAdd() As IPAddress = ipEnter.AddressList
                For i As Integer = 0 To IpAdd.GetUpperBound(0)
                    UltraListView1.Items.Add(3, "Machine IP : " & IpAdd(i).ToString)
                    VIP = IpAdd(i).ToString
                Next
            Else
                Try
                    'UltraListView1.Items.Add(3, "Machine IP : " & System.Net.Dns.GetHostAddresses(My.Computer.Name)(2).ToString)
                    'VIP = System.Net.Dns.GetHostAddresses(My.Computer.Name)(2).ToString
                Catch
                    Timer1.Enabled = False
                    MsgBox("Unable to return Network Details", MsgBoxStyle.Critical, "Technova ....")

                    Exit Sub
                End Try
            End If
            ' UltraListView1.Items.Add(4, "Logged User : " & win.Name)
            VDominLoggedUser = win.Name
            UltraListView1.Items.Add(5, "Data Source : " & Mid(ConfigurationManager.AppSettings("CD"), 13, 13))
            VDataSource = Mid(ConfigurationManager.AppSettings("CD"), 13, 13)
            UltraListView1.Items.Add(6, "Database : " & Mid(ConfigurationManager.AppSettings("CD"), 43, 8))
            VDataBase = Mid(ConfigurationManager.AppSettings("CD"), 43, 8)
            UltraListView1.Items.Add(7, "Application Style : " & ConfigurationManager.AppSettings("APPSTL"))





        Catch returnMessage As Exception
            'If returnMessage.Message <> Nothing Then
            ' Exit Sub
            ' MessageBox.Show(returnMessage.Message)
            Me.Close()
            MDIMain.Close()
            ' End If
        End Try
    End Sub
    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Dim sRect As Rectangle
        ' Infragistics.Win.AppStyling.StyleManager.Load(ConfigurationManager.AppSettings("APPSTL"))
        Me.Top = ((800 / 2) - (Me.Height / 2)) - 50
        Me.Left = 300
        Me.Width = 500
        Me.Height = 300
        Me.Show()
        Timer1.Interval = 500
        Timer1.Enabled = True
        AddHandler Timer1.Tick, AddressOf OnTimerEvent

    End Sub

    Private Sub Form2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
    End Sub

    Public Shared Sub OnTimerEvent(ByVal [source] As Object, ByVal e As EventArgs)

        SPL.Timer1.Enabled = False
        ' frmLog.Show()
        MDIMain.Show()

    End Sub 'OnTimerEvent

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

    End Sub

    Private Sub UltraLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraLabel2.Click

    End Sub
End Class