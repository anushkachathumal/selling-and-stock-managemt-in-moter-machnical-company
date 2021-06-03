Imports System.Data.SqlClient
Imports Infragistics.Win.AppStyling
Imports System.Data
Imports System.Configuration
Imports System
Imports System.Collections


Public Class Login
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Public returnMessage As String

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Function ValidateFields() As String

        Dim SQL As String
        Dim con = New SqlConnection()

        Try
            con = DBEngin.GetConnection()
            SQL = "SELECT USERID, PASSWORD  FROM USERS WHERE (USERID ='" & txtUserName.Text & "') "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)

            If dsUser.Tables(0).Rows.Count = 0 Then
                returnMessage = "User name and pasword combination not found"
                'txtUserName.Text
                txtUserName.Focus()
                Return returnMessage
            End If

            If Trim(dsUser.Tables(0).Rows(0)("PASSWORD")) <> Trim((txtPassword.Text)) Then 'Trim(ObjSecurity.Encrypt(txtPassword.Text))
                returnMessage = "User name and pasword combination not found"
                ' txtUserName.SelectAll()
                txtUserName.Focus()
                Return returnMessage
            End If

            'Return "NOERROR"

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If

        End Try

    End Function

    Private Sub Login_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        txtUserName.Focus()
    End Sub

    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Infragistics.Win.AppStyling.StyleManager.Load(ConfigurationManager.AppSettings("APPSTL"))

    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub OK_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        ValidateFields()
        Me.Hide()
        MDIMain.Show()
    End Sub

    Private Sub cancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancel.Click
        Me.Close()
    End Sub
End Class
