Imports System.Data
Imports System.Data.SqlClient

Public Class BLogin

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.
    Private v = My.Settings.BAR
    Private Shadows m = My.Settings.TOS
    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click

        If (UsernameTextBox.Text = My.Settings.Username And PasswordTextBox.Text = My.Settings.Password) Then
            Rinfosys.Show()
            _initIdentifier()
            Me.Hide()
        Else
            MsgBox("Invalid credentials! Please try again.", MsgBoxStyle.Exclamation)
            Rinfosys.Visible = False
        End If


    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        My.Settings.Save()
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub BLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      

    End Sub
End Class
