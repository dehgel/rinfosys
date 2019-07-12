Imports Microsoft.Win32.Registry
Imports System.Data
Imports System.Data.SqlClient

Module ActivatorFunctions
    Private i As Integer
    Private xhole As New Random


    Public ConnStr As String = My.Settings.RFS_DBCONN_STR
    Public conn As New SqlConnection(ConnStr)
    Public cmd As New SqlCommand
    Public reader As SqlDataReader
    Public adap As SqlDataAdapter
    Public this As String
    
    Function __GetOneRand(ByVal key As String, ByVal value As String, ByVal delete As String) As Single
        
        Using conn As New SqlConnection(ConnStr)
            Try
                Using cmd As New SqlCommand("SELECT KEY FROM RFS_ACTS WHERE KEY='" & key & "'", conn)
                    If key = value Then

                    End If
                End Using
            Catch ex As Exception
                MsgBox("This is not the time you would spent for this service.", MsgBoxStyle.Critical)
                Exit Function
            Finally
                Using cmd As New SqlCommand("DELETE FROM RFS_ACTS WHERE KEY='" & key & "'", conn)
                    conn.Open()
                    cmd.ExecuteNonQuery()
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("INSERT INTO RFS_TOKEN(RDX, RDY) VALUES(" & key & ", " & delete & ")", conn)
                    conn.Open()
                    cmd.ExecuteNonQuery()
                    conn.Close()
                End Using
            End Try
        End Using




    End Function
End Module
