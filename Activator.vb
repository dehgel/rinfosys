Imports System.Data
Imports System.Data.SqlClient
Imports System.Linq

Public Class Activator
    Private v = My.Application.Info.AssemblyName
    Private r As String
    Private i As Integer
    Public ConnStr As String = My.Settings.RFS_DBCONN_STR
  
    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button33.Click
        Dim n = ActivationKeyBox.Text
        Dim t = TextBox3.Text
        Using conn As New SqlConnection(ConnStr)
            Using cmd As New SqlCommand("SELECT * FROM RFS_ACTS WHERE THEK='" & n & "'", conn)
                Try
                    conn.Open()
                    Dim reader As SqlDataReader = cmd.ExecuteReader
                    While reader.Read
                        If Trim(reader("THEK") = n) Then
                            My.Computer.Registry.CurrentUser.CreateSubKey(v, Microsoft.Win32.RegistryKeyPermissionCheck.Default).SetValue(t, n, RegistryValueKind.String)
                            My.Settings.BAR = TextBox3.Text
                            My.Settings.BCM103 = TextBox2.Text
                            My.Settings.BP102 = TextBox1.Text
                            My.Settings.Username = TextBox4.Text
                            My.Settings.Password = TextBox5.Text
                            My.Settings.TOKEN = n
                            My.Settings.Save()
                            Me.Close()
                        Else
                            MsgBox("Authentication not valid.", MsgBoxStyle.MsgBoxHelp)
                        End If
                    End While
                    conn.Close()
                Catch ex As Exception

                Finally
                    _initIdentifier()
                End Try

            End Using
        End Using

       
    End Sub

    Private Sub Activator_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        My.Settings.Save()

    End Sub

    Private Sub Activator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
       
        Try

        
        Catch ex As Exception
            If MsgBox("Invalid Activation Key!", MsgBoxStyle.MsgBoxHelp) Then

            End If
        End Try
        
    End Sub

    Private Sub ActivationKeyBox_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ActivationKeyBox.KeyUp
        Try
            Using conn As New SqlConnection(ConnStr)
                Using cmd As New SqlCommand("SELECT THEK FROM RFS_ACTS WHERE THEK='" & ActivationKeyBox.Text & "'", conn)
                    conn.Open()
                    Dim reader As SqlDataReader = cmd.ExecuteReader
                    While reader.Read
                        If String.Equals(Trim(ActivationKeyBox.Text), Trim(reader("THEK"))) = True Then
                            ActivationKeyBox.BackColor = System.Drawing.Color.GreenYellow
                            Button33.Enabled = True
                            Button33.Text = "ACTIVATE NOW"
                            Button33.BackColor = System.Drawing.Color.SpringGreen
                            Button33.ForeColor = System.Drawing.Color.Black
                        Else
                            Button33.Text = "Activate Key"
                            ActivationKeyBox.BackColor = System.Drawing.Color.Coral
                            Button33.Enabled = False
                            Button33.BackColor = System.Drawing.Color.DodgerBlue
                        End If
                    End While

                    conn.Close()
                End Using
            End Using
        Catch ex As Exception
            ActivationKeyBox.BackColor = System.Drawing.Color.Coral
        Finally
            If ActivationKeyBox.Text = String.Empty Then
                Button33.Text = "Activate Key"
                Button33.BackColor = System.Drawing.Color.DodgerBlue
                Button33.ForeColor = System.Drawing.Color.White
                Button33.Enabled = False
            End If
        End Try
        
    End Sub

    Private Sub ActivationKeyBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ActivationKeyBox.TextChanged
        If ActivationKeyBox.Text = String.Empty Then
            ActivationKeyBox.BackColor = System.Drawing.Color.White
        End If
    End Sub

  
End Class