Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient

Public Class ResidentProfile
    Public ConnStr As String = My.Settings.RFS_DBCONN_STR
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Dim pd As New PrintPreviewDialog
        pd.Document = Rinfosys.prtProfile
        Rinfosys.prtProfile.PrinterSettings = New PrinterSettings
        Dim w As Integer = 827
        Dim h As Integer = 1167
        Dim PpName As String = "A4"
        Dim psz As New PaperSize(PpName, w, h)
        Rinfosys.prtProfile.PrinterSettings.DefaultPageSettings.PaperSize = psz
        pd.Document = Rinfosys.prtProfile
        Rinfosys.prtProfile.DocumentName = "Print Certificate"
        pd.ShowDialog()

    End Sub

    Private Sub ResidentProfile_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Cursor = Cursors.WaitCursor 'and some various me.Cursor / current.cursor 
        Application.DoEvents()
        OK_Button.Enabled = False
        Cursor = Cursors.Default
        Dim conn As New SqlConnection(ConnStr)
        conn.Open()
        Dim com As String = "SELECT * FROM RFS_CHILD WHERE ID='" & TextBox4.Text & "'"
        Dim Adpt As New SqlDataAdapter(com, conn)
        Dim ds As New DataSet()
        Adpt.Fill(ds, "RFS_CHILD")

        DataGridView1.DataSource = ds.Tables(0)
        For i As Integer = 1 To ds.Tables(0).Rows.Count
            DataGridView1.Rows(i).Cells(0).Value = i
        Next

        conn.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DataGridView1.RowsAdded

    End Sub
End Class
