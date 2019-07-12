Imports System.Windows.Forms
Imports System.Drawing.Image
Imports System.Data
Imports System.Data.SqlClient

Public Class ConfigDiag

    Public ConnStr As String = My.Settings.RFS_DBCONN_STR
    Public conn As New SqlConnection(ConnStr)
    Public cmd As New SqlCommand
    Public reader As SqlDataReader
    Public adap As SqlDataAdapter
    Public this As String
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Try
            If System.Windows.Forms.DialogResult.OK Then
                My.Settings.DBUser = TextBox2.Text
                My.Settings.DBPass = TextBox3.Text
                My.Settings.BR101 = TextBox9.Text
                My.Settings.BAR = TextBox6.Text
                My.Settings.BChair = TextBox7.Text
                My.Settings.BCM103 = TextBox5.Text
                My.Settings.BP102 = TextBox4.Text
                My.Settings.BSec = TextBox8.Text

                If CheckBox1.Checked = True Then
                    My.Settings.Checked = True
                Else
                    My.Settings.Checked = False
                End If

                My.Settings.Save()
                My.Settings.Reload()
                Me.Close()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            With Rinfosys.ListBox1.Items
                .Add("System Configured: " & Process.GetCurrentProcess.Id)
                .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") & " - Timestamp")
            End With
        End Try


    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        My.Settings.Reload()
        Me.Close()
    End Sub

    Private Sub ConfigDiag_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            TextBox1.Text = My.Settings.CustomConnection
            TextBox2.Text = My.Settings.DBUser
            TextBox3.Text = My.Settings.DBPass
            TextBox4.Text = My.Settings.BP102
            TextBox5.Text = My.Settings.BCM103
            TextBox6.Text = My.Settings.BAR
            TextBox7.Text = My.Settings.BChair
            TextBox8.Text = My.Settings.BSec
            TextBox9.Text = My.Settings.BR101
            TextBox10.Text = My.Settings.CMLOGO
            TextBox11.Text = My.Settings.BRNGYLOGO
            If My.Settings.Checked = True Then
                CheckBox1.Checked = True
            Else
                CheckBox1.Checked = False
            End If
            If CheckBox3.Checked = True Then

                IDBackground.BackColor = System.Drawing.Color.Transparent
                HeaderBackColor.BackColor = System.Drawing.Color.Transparent
                LinePanel.BackColor = System.Drawing.Color.Transparent
                Label21.ForeColor = My.Settings.CIDHEADTEXTC
                Label22.ForeColor = My.Settings.CIDBODYTEXTC
                Label23.ForeColor = My.Settings.CIDBODYTEXTC
                Label24.ForeColor = My.Settings.CIDBODYTEXTC
                HeaderBackColor.BackColor = My.Settings.CIDHEADCOLOR
                IDBackground.BackColor = My.Settings.CIDBACKC
                LinePanel.BackColor = My.Settings.CIDLINEC
                Button8.BackColor = My.Settings.CIDHEADTEXTC
                Button9.BackColor = My.Settings.CIDBACKC
                Button10.BackColor = My.Settings.CIDHEADCOLOR
                Button13.BackColor = My.Settings.CIDBODYTEXTC
                Button11.BackColor = My.Settings.CIDLINEC
                If My.Settings.IDBI.Length > 1 Then
                    IDBackground.BackgroundImage = System.Drawing.Image.FromFile(My.Settings.IDBI)
                End If
            Else

                IDBackground.BackColor = System.Drawing.Color.Transparent
                HeaderBackColor.BackColor = System.Drawing.Color.Transparent
                LinePanel.BackColor = System.Drawing.Color.Transparent
                Label21.ForeColor = System.Drawing.Color.Black
                Label22.ForeColor = System.Drawing.Color.Black
                Label23.ForeColor = System.Drawing.Color.Black
                Label24.ForeColor = System.Drawing.Color.Black
                HeaderBackColor.BackColor = System.Drawing.Color.Transparent
                IDBackground.BackColor = System.Drawing.Color.Transparent
                LinePanel.BackColor = System.Drawing.Color.Transparent
                Button8.BackColor = System.Drawing.Color.Transparent
                Button9.BackColor = System.Drawing.Color.Transparent
                Button10.BackColor = System.Drawing.Color.Transparent
                Button13.BackColor = System.Drawing.Color.Transparent
                Button11.BackColor = System.Drawing.Color.Transparent

            End If
            IDBackground.BackgroundImageLayout = ImageLayout.Zoom
            TemplateLogoBR.BackgroundImage = System.Drawing.Image.FromFile(TextBox11.Text)

            If My.Settings.IDCheckBox = False Then
                Select Case My.Settings.IDCOMBOSELECT.ToString
                    Case My.Settings.IDCOMBOSELECT = "Basic_1"
                        IDBackground.BackgroundImage = My.Resources.Basic_1
                    Case My.Settings.IDCOMBOSELECT = "Basic_2"
                        IDBackground.BackgroundImage = My.Resources.Basic_2
                    Case My.Settings.IDCOMBOSELECT = "Basic_3"
                        IDBackground.BackgroundImage = My.Resources.Basic_3
                    Case My.Settings.IDCOMBOSELECT = "Classic_1"
                        IDBackground.BackgroundImage = My.Resources.Classic_1
                    Case My.Settings.IDCOMBOSELECT = "Classic_2"
                        IDBackground.BackgroundImage = My.Resources.Classic_2
                    Case My.Settings.IDCOMBOSELECT = "Classic_3"
                        IDBackground.BackgroundImage = My.Resources.Classic_3
                    Case My.Settings.IDCOMBOSELECT = "Modern_1"
                        IDBackground.BackgroundImage = My.Resources.Modern_1
                    Case My.Settings.IDCOMBOSELECT = "Modern_2"
                        IDBackground.BackgroundImage = My.Resources.Modern_2
                    Case My.Settings.IDCOMBOSELECT = "Modern_3"
                        IDBackground.BackgroundImage = My.Resources.Modern_3
                End Select
                ComboBox1.Text = My.Settings.IDCOMBOSELECT
            Else
                ComboBox1.Enabled = False
            End If

        Catch ex As Exception
            Me.Show()
        End Try

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If Me.CheckBox1.Checked = True Then
            DBsettings.Enabled = True
        Else
            DBsettings.Enabled = False
            My.Settings.CustomConnection = My.Settings.RFS_DBCONN_STR
        End If
    End Sub

    Private Sub OKAppBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKAppBtn.Click
        My.Settings.CMLOGO = TextBox10.Text
        My.Settings.BRNGYLOGO = TextBox11.Text
        My.Settings.RESNO = TextBox12.Text
        My.Settings.RESTITLE = TextBox13.Text
        With Rinfosys.ListBox1.Items
            .Add("Application Configured: " & Process.GetCurrentProcess.Id)
            .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") & " - Timestamp")
        End With
        My.Settings.Save()
        Me.Close()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim opd = New OpenFileDialog
        opd.Title = "Select Photo"
        opd.InitialDirectory = opd.FileName
        opd.RestoreDirectory = True
        opd.Filter = "All Images(*.*)|*.jpg;*.png;*.gif;*.bmp;*.tiff|JPEG (*.jpg) |*.jpg|PNG (*.png)|*.png"
        If opd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            TextBox10.Text = opd.FileName
            My.Settings.CMLOGO = opd.FileName

        End If
        opd.Dispose()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim opd = New OpenFileDialog
        opd.Title = "Select Photo"
        opd.InitialDirectory = opd.FileName
        opd.RestoreDirectory = True
        opd.Filter = "All Images(*.*)|*.jpg;*.png;*.gif;*.bmp;*.tiff|JPEG (*.jpg) |*.jpg|PNG (*.png)|*.png"
        If opd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            TextBox11.Text = opd.FileName
            My.Settings.BRNGYLOGO = opd.FileName
            Rinfosys.MainLogoBox.BackgroundImage = System.Drawing.Image.FromFile(My.Settings.BRNGYLOGO)
            TemplateLogoBR.BackgroundImage = System.Drawing.Image.FromFile(opd.FileName)

        End If
        opd.Dispose()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If conn.State = ConnectionState.Connecting Then
            Using conn
                With conn
                    .ConnectionString = My.Settings.CustomConnection
                    .Open()

                    Button1.Text = "Testing"
                    .Close()
                End With
            End Using

        ElseIf conn.State = ConnectionState.Broken Then
            MsgBox("Error Database Connection")
        ElseIf conn.State = ConnectionState.Open Then
            Button1.Enabled = True
            Button1.Text = "Success!"
            Button1.BackColor = System.Drawing.Color.Green
        End If

        My.Settings.CustomConnection = TextBox1.Text

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim opd = New OpenFileDialog
        opd.Title = "Select Photo"
        opd.InitialDirectory = opd.FileName
        opd.RestoreDirectory = True
        opd.Filter = "All Database Files(*.*)|*.sql;*.mdb;*.mdf|MDF (*.mdf) |*.mdf|MDB (*.mdb)|*.mdb"
        If opd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            TextBox1.Text = opd.FileName

        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim cdg As New ColorDialog

        If cdg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Button8.BackColor = cdg.Color
            Label21.ForeColor = cdg.Color
            My.Settings.CIDHEADTEXTC = cdg.Color
        End If

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim cdg As New ColorDialog

        If cdg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Button10.BackColor = cdg.Color
            HeaderBackColor.BackColor = cdg.Color
            My.Settings.CIDHEADCOLOR = cdg.Color

        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim cdg As New ColorDialog

        If cdg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Button9.BackColor = cdg.Color
            IDBackground.BackColor = cdg.Color
            IDBackground.BackgroundImage = Nothing
            My.Settings.IDBI = Nothing
            My.Settings.CIDBACKC = cdg.Color

        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim cdg As New ColorDialog

        If cdg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Button11.BackColor = cdg.Color
            LinePanel.BackColor = cdg.Color
            My.Settings.CIDLINEC = cdg.Color

        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim opd = New OpenFileDialog
        opd.Title = "Select Photo"
        opd.InitialDirectory = opd.FileName
        opd.RestoreDirectory = True
        opd.Filter = "All Images(*.*)|*.jpg;*.png;*.gif;*.bmp;*.tiff|JPEG (*.jpg) |*.jpg|PNG (*.png)|*.png"
        If opd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            HeaderBackColor.BackgroundImage = System.Drawing.Image.FromFile(opd.FileName)
            My.Settings.CIDIDHEADIMG = opd.FileName
        End If

    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim cdg As New ColorDialog

        If cdg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Label22.ForeColor = cdg.Color
            Label23.ForeColor = cdg.Color
            Label24.ForeColor = cdg.Color
            My.Settings.CIDBODYTEXTC = cdg.Color

        End If



    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim opd = New OpenFileDialog
        opd.Title = "Select Photo"
        opd.InitialDirectory = opd.FileName
        opd.RestoreDirectory = True
        opd.Filter = "All Images(*.*)|*.jpg;*.png;*.gif;*.bmp;*.tiff|JPEG (*.jpg) |*.jpg|PNG (*.png)|*.png"
        If opd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            IDBackground.BackgroundImage = System.Drawing.Image.FromFile(opd.FileName)
            IDBackground.BackgroundImageLayout = ImageLayout.Zoom
            My.Settings.IDBI = opd.FileName

        End If
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Dim opd = New OpenFileDialog
        opd.Title = "Select Photo"
        opd.InitialDirectory = opd.FileName
        opd.RestoreDirectory = True
        opd.Filter = "All Images(*.*)|*.jpg;*.png;*.gif;*.bmp;*.tiff|JPEG (*.jpg) |*.jpg|PNG (*.png)|*.png"
        If opd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            HeaderBackColor.BackgroundImage = System.Drawing.Image.FromFile(opd.FileName)
            My.Settings.IDWM = opd.FileName
            LinkLabel1.Show()

        End If
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Try
            Label21.ForeColor = System.Drawing.Color.Black
            Label22.ForeColor = System.Drawing.Color.Black
            Label23.ForeColor = System.Drawing.Color.Black
            Label24.ForeColor = System.Drawing.Color.Black
            HeaderBackColor.BackColor = System.Drawing.Color.Transparent
            IDBackground.BackColor = System.Drawing.Color.White
            LinePanel.BackColor = System.Drawing.Color.Transparent
            Button8.BackColor = System.Drawing.Color.Transparent
            Button9.BackColor = System.Drawing.Color.Transparent
            Button10.BackColor = System.Drawing.Color.Transparent
            Button13.BackColor = System.Drawing.Color.Transparent
            Button11.BackColor = System.Drawing.Color.Transparent
        Catch ex As Exception
            MsgBox("DIAGX-1001:" & ex.Message)
        End Try

    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            CustomIDPanel.Enabled = True
            My.Settings.IDCheckBox = True
            ComboBox1.Enabled = False

            IDBackground.BackColor = System.Drawing.Color.Transparent
            HeaderBackColor.BackColor = System.Drawing.Color.Transparent
            LinePanel.BackColor = System.Drawing.Color.Transparent
            Label21.ForeColor = My.Settings.CIDHEADTEXTC
            Label22.ForeColor = My.Settings.CIDBODYTEXTC
            Label23.ForeColor = My.Settings.CIDBODYTEXTC
            Label24.ForeColor = My.Settings.CIDBODYTEXTC
            HeaderBackColor.BackColor = My.Settings.CIDHEADCOLOR
            IDBackground.BackColor = My.Settings.CIDBACKC
            LinePanel.BackColor = My.Settings.CIDLINEC
            Button8.BackColor = My.Settings.CIDHEADTEXTC
            Button9.BackColor = My.Settings.CIDBACKC
            Button10.BackColor = My.Settings.CIDHEADCOLOR
            Button13.BackColor = My.Settings.CIDBODYTEXTC
            Button11.BackColor = My.Settings.CIDLINEC
            If My.Settings.IDBI.Length > 1 Then
                IDBackground.BackgroundImage = System.Drawing.Image.FromFile(My.Settings.IDBI)
            End If

        Else
            CustomIDPanel.Enabled = False
            My.Settings.IDCheckBox = False
            ComboBox1.Enabled = True

            IDBackground.BackColor = System.Drawing.Color.Transparent
            HeaderBackColor.BackColor = System.Drawing.Color.Transparent
            LinePanel.BackColor = System.Drawing.Color.Transparent
            Label21.ForeColor = System.Drawing.Color.Black
            Label22.ForeColor = System.Drawing.Color.Black
            Label23.ForeColor = System.Drawing.Color.Black
            Label24.ForeColor = System.Drawing.Color.Black
            HeaderBackColor.BackColor = System.Drawing.Color.Transparent
            IDBackground.BackColor = System.Drawing.Color.Transparent
            LinePanel.BackColor = System.Drawing.Color.Transparent
            Button8.BackColor = System.Drawing.Color.Transparent
            Button9.BackColor = System.Drawing.Color.Transparent
            Button10.BackColor = System.Drawing.Color.Transparent
            Button13.BackColor = System.Drawing.Color.Transparent
            Button11.BackColor = System.Drawing.Color.Transparent
            If My.Settings.IDBI.Length > 1 Then
                IDBackground.BackgroundImage = System.Drawing.Image.FromFile(My.Settings.IDBI)
            End If
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        My.Settings.IDWM = ""
        LinkLabel1.Hide()
    End Sub
    Private Function __SELECTHEME() As Single

        Select Case (ComboBox1.SelectedItem)
            Case "Basic_1"
                IDBackground.BackgroundImage = My.Resources.Basic_1
                My.Settings.IDCOMBOSELECT = "Basic_1"
            Case "Basic_2"
                IDBackground.BackgroundImage = My.Resources.Basic_2
                My.Settings.IDCOMBOSELECT = "Basic_2"
            Case "Basic_3"
                IDBackground.BackgroundImage = My.Resources.Basic_3
                My.Settings.IDCOMBOSELECT = "Basic_3"
            Case "Classic_1"
                IDBackground.BackgroundImage = My.Resources.Classic_1
                My.Settings.IDCOMBOSELECT = "Classic_1"
            Case "Classic_2"
                IDBackground.BackgroundImage = My.Resources.Classic_2
                My.Settings.IDCOMBOSELECT = "Classic_2"
            Case "Classic_3"
                IDBackground.BackgroundImage = My.Resources.Classic_3
                My.Settings.IDCOMBOSELECT = "Classic_3"
            Case "Modern_1"
                IDBackground.BackgroundImage = My.Resources.Modern_1
                My.Settings.IDCOMBOSELECT = "Modern_1"
            Case "Modern_2"
                IDBackground.BackgroundImage = My.Resources.Modern_2
                My.Settings.IDCOMBOSELECT = "Modern_2"
            Case "Modern_3"
                IDBackground.BackgroundImage = My.Resources.Modern_3
                My.Settings.IDCOMBOSELECT = "Modern_3"
        End Select
    End Function
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        __SELECTHEME()
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click

        If PasswordAcc.Text = ConfirmPassAcc.Text Then
            My.Settings.Username = UsernameAcc.Text
            My.Settings.Password = ConfirmPassAcc.Text
        Else
            MsgBox("Mismatched! Please try again.", MsgBoxStyle.MsgBoxHelp)
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            UsernameAcc.Enabled = True
            PasswordAcc.Enabled = True
            ConfirmPassAcc.Enabled = True
        Else
            UsernameAcc.Enabled = False
            PasswordAcc.Enabled = False
            ConfirmPassAcc.Enabled = False
        End If
    End Sub


    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Dim sfd As New SaveFileDialog
        sfd.FileName = "RFS_DB"
        sfd.InitialDirectory = System.Environment.CurrentDirectory
        sfd.RestoreDirectory = True
        sfd.Filter = "MDF (*.mdf)|*.mdf"
        If sfd.ShowDialog = Forms.DialogResult.OK Then
            Process.Start("copy '" & System.Environment.CurrentDirectory + "/RFS_DB.mdf" & " " & sfd.FileName & "'")
        End If

    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Process.Start("http://noah.up.edu.ph/#/")
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        My.Settings.Elevation = TextBox14.Text
        My.Settings.Latitude = TextBox15.Text
        My.Settings.Longitude = TextBox14.Text
        My.Settings.PostalCode = TextBox16.Text
        My.Settings.Island = TextBox17.Text
        My.Settings.Save()

    End Sub
End Class
