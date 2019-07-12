Public Class LogData

    Private Sub LogData_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _LOG()
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ClearDataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearDataToolStripMenuItem.Click
        _TRUNCATE("LOGGER")
    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub ContextMenuStrip1_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening
        Select Case True
            Case SplitContainer1.Panel1.HasChildren
                _TRUNCATE("LOGGER")
            Case SplitContainer1.Panel2.HasChildren
                _TRUNCATE("RFS_LOG")
            Case Else

        End Select
    End Sub
End Class