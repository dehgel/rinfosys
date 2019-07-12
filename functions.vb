Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Drawing.Image
Imports KeepAutomation.Barcode.Bean
Imports KeepAutomation.Barcode
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Drawing.Printing
Imports Microsoft.Office.Interop

Module functions
    Public ConnStr As String = My.Settings.RFS_DBCONN_STR
    Public conn As New SqlConnection(ConnStr)
    Public cmd As New SqlCommand
    Public reader As SqlDataReader
    Public adap As SqlDataAdapter
    Public this As String
    Public bytes As Byte()
    Public ms As New MemoryStream()
    Public TheImg As New PictureBox

    Public Function GenerateRandomNumber(ByVal min As Integer, ByVal max As Integer) As Integer
        Dim random As New Random()
        Return random.Next(min, max)
    End Function
    Function _RESETID() As Single
        Rinfosys.IDno.Text = GenerateRandomNumber(1234568, 99999999)
        Rinfosys.DIDno.Text = GenerateRandomNumber(1987654, 88888888)
    End Function
    Function _APP_UI() As Integer
        Try
            Rinfosys.PrintPanel.Hide()
            Rinfosys.TextBox19.Enabled = False
            Rinfosys.IDno.Text = GenerateRandomNumber(1234568, 99999999)
            Rinfosys.DIDno.Text = GenerateRandomNumber(1987654, 88888888)
            Rinfosys.BCID.Text = GenerateRandomNumber(2134659, 99999999)
            If Trim(My.Settings.BRNGYLOGO).Length > 0 Then
                Rinfosys.MainLogoBox.BackgroundImage = System.Drawing.Image.FromFile(My.Settings.BRNGYLOGO)
            Else
                Rinfosys.MainLogoBox.BackgroundImage = My.Resources.rinfosysIcon
            End If
            Rinfosys.TableLayoutPanel3.Hide()
            Rinfosys.StatusTextBox.Hide()
            Rinfosys.Button32.Hide()
        Catch ex As Exception

        Finally

            If Rinfosys.DataGridView1.RowCount <= 0 Then
                Rinfosys.Button12.Enabled = False
                Rinfosys.Button15.Enabled = False
                Rinfosys.Button16.Enabled = False
            Else
                Rinfosys.Button12.Enabled = True
                Rinfosys.Button15.Enabled = True
                Rinfosys.Button16.Enabled = True
            End If

            
      
        End Try

    End Function
    Function __InsertHouseholdsHousing() As Single
        Using conn As New SqlConnection(ConnStr)
            Using cmd As New SqlCommand("INSERT INTO RFS_HHM(ID, FNAME, MNAME, LNAME, RCS, CHILDREN, HOMAT) VALUES(@A, @B, @C, @D, @E, @F, @G)", conn)
                With cmd.Parameters
                    .AddWithValue("@A", Rinfosys.IDno.Text)
                    .AddWithValue("@B", Rinfosys.TextBox2.Text)
                    .AddWithValue("@C", Rinfosys.TextBox3.Text)
                    .AddWithValue("@D", Rinfosys.TextBox1.Text)
                    .AddWithValue("@E", Rinfosys.ComboBox1.Text)
                    .AddWithValue("@F", "")
                    .AddWithValue("@G", Rinfosys.ComboBox9.Text)
                End With
            End Using
        End Using
    End Function

    Function _init_db() As Single

        Try
            conn.Open()
            Dim com As String = "SELECT * FROM RFS_PID"
            Dim Adpt As New SqlDataAdapter(com, conn)
            Dim ds As New DataSet()
            Adpt.Fill(ds, "RFS_PID")
            Rinfosys.DataGridView1.DataSource = ds.Tables(0)
            conn.Close()

            With Rinfosys.DataGridView1
                .RowHeadersVisible = False
                .Columns(0).HeaderCell.Value = "ENTRY NO."
                .Columns(1).HeaderCell.Value = "CONTROL ID"
                .Columns(2).HeaderCell.Value = "TOKEN ID"
                .Columns(3).HeaderCell.Value = "LAST NAME"
                .Columns(4).HeaderCell.Value = "FIRSTNAME"
                .Columns(5).HeaderCell.Value = "MIDDLENAME"
                .Columns(6).HeaderCell.Value = "SUFFIX"
                .Columns(7).HeaderCell.Value = "BIRTHDATE"
                .Columns(8).HeaderCell.Value = "BIRTHPLACE"
                .Columns(9).HeaderCell.Value = "STATUS"
                .Columns(10).HeaderCell.Value = "GENDER"
                .Columns(11).HeaderCell.Value = "RELIGION"
                .Columns(12).HeaderCell.Value = "CITIZENSHIP"
                .Columns(13).HeaderCell.Value = "RESIDENTIAL"
                .Columns(14).HeaderCell.Value = "ADDRESS"
                .Columns(15).HeaderCell.Value = "EDUCATION"
                .Columns(16).HeaderCell.Value = "S.ADDRESS"
                .Columns(17).HeaderCell.Value = "S.YEAR"
                .Columns(18).HeaderCell.Value = "OCCUPATION"
                .Columns(19).HeaderCell.Value = "OFFICE ADD"
                .Columns(20).HeaderCell.Value = "POSITION"
                .Columns(21).HeaderCell.Value = "ORGANIZATION"
                .Columns(22).HeaderCell.Value = "TYPE"
                .Columns(23).HeaderCell.Value = "MEMBERSHIP"

            End With

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            __BUSINIT()
        End Try


    End Function
    Function __BUSINIT() As Single
        Try
            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim com As String = "SELECT * FROM RFS_BUS"
            Dim Adpt As New SqlDataAdapter(com, conn)
            Dim ds As New DataSet()
            Adpt.Fill(ds, "RFS_BUS")
            Rinfosys.BUSDGriview.DataSource = ds.Tables(0)
            conn.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            With Rinfosys.BUSDGriview
                .RowHeadersVisible = False
                .Columns(0).HeaderCell.Value = "BUS. ID"
                .Columns(1).HeaderCell.Value = "DOC. NO."
                .Columns(2).HeaderCell.Value = "OWNER"
                .Columns(3).HeaderCell.Value = "BUS. NAME"
                .Columns(4).HeaderCell.Value = "INDUSTRY"
                .Columns(5).HeaderCell.Value = "ADDRESS"
                .Columns(6).HeaderCell.Value = "NATURE"
                .Columns(7).HeaderCell.Value = "TIN NO."
                .Columns(8).HeaderCell.Value = "OR NO."
                .Columns(9).HeaderCell.Value = "PAID"
                .Columns(10).HeaderCell.Value = "DATE REG."
                .Columns(11).HeaderCell.Value = "REMARKS"
            End With
        End Try
    End Function
    Function __LISTING() As Single
        Dim r = Rinfosys
        Try
            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim com As String = "SELECT * FROM RFS_EVT"
            Dim cmd As String = "SELECT EVTNAME FROM RFS_EVT"
            Dim Adpt As New SqlDataAdapter(com, conn)
            Dim asdf As New SqlDataAdapter(cmd, conn)
            Dim dd As New DataSet

            Dim ds As New DataSet()
            Adpt.Fill(ds, "RFS_EVT")
            Rinfosys.DataGridView3.DataSource = ds.Tables(0)

            With r.ComboBox8
                asdf.Fill(dd, "RFS_EVT")
                .DataSource = dd.Tables(0)
                .DisplayMember = "EVTNAME"
            End With

            With Rinfosys.DataGridView3
                .RowHeadersVisible = False
                .Columns(0).HeaderCell.Value = "NO."
                .Columns(1).HeaderCell.Value = "NAME"
                .Columns(2).HeaderCell.Value = "EVENT"
                .Columns(3).HeaderCell.Value = "ORGANIZATION"
                .Columns(4).HeaderCell.Value = "TYPE OF ORG"
                .Columns(5).HeaderCell.Value = "DATE ENTRY"

            End With

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    Function __Upload() As Single
        Dim opd = New System.Windows.Forms.OpenFileDialog
        opd.Title = "Select Photo"
        opd.InitialDirectory = opd.FileName
        opd.RestoreDirectory = True
        opd.Filter = "All Images(*.*)|*.jpg;*.png;*.gif;*.bmp;*.tiff|JPEG (*.jpg) |*.jpg|PNG (*.png)|*.png"
        If opd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Rinfosys.PictureBox2.ImageLocation = opd.FileName
            Rinfosys.PictureBox2.SizeMode = PictureBoxSizeMode.Zoom
            'TheImg.ImageLocation = opd.FileName
        End If
        opd.Dispose()
    End Function
    Function __LogEntry() As Single
        Dim logcmd As String = String.Empty
        logcmd &= "INSERT INTO LOGGER(ID, NAME, LOGDATE, IMAGE, SP) VALUES(@ID, @N, @DT, @PIC, @S)"

        'Dim data As Byte() = DirectCast(dr("Picture"), Byte())
        'Dim ms As New MemoryStream(data)
        'Edt_PictureBox.Image = Image.FromStream(ms)

        Using conn As New SqlConnection(ConnStr)
            Using cmd As New SqlCommand

                Dim ThisName = Trim(Rinfosys.TextBox1.Text & ", " & Rinfosys.TextBox2.Text & " " & Rinfosys.TextBox3.Text).ToUpper
                With cmd
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = logcmd
                    .Parameters.AddWithValue("@ID", Rinfosys.IDno.Text)
                    .Parameters.AddWithValue("@N", ThisName)
                    .Parameters.AddWithValue("@DT", DateTime.Today)

                    If Rinfosys.PictureBox2.ImageLocation.Length > 1 Then
                        Using fs As FileStream = New FileStream(Rinfosys.PictureBox2.ImageLocation, FileMode.Open, FileAccess.Read)
                            Dim bytes(fs.Length) As Byte
                            fs.Read(bytes, 0, fs.Length)
                            ms = New MemoryStream(bytes)
                        End Using
                        .Parameters.AddWithValue("@PIC", Rinfosys.PictureBox2.Image).Value = ms.ToArray
                    Else
                        Dim LocalImg = System.Environment.CurrentDirectory & "../Images/profile.jpg"
                        Using fs As FileStream = New FileStream(LocalImg, FileMode.Open, FileAccess.Read)
                            Dim bytes(fs.Length) As Byte
                            fs.Read(bytes, 0, fs.Length)
                            ms = New MemoryStream(bytes)
                        End Using
                        .Parameters.AddWithValue("@PIC", Rinfosys.PictureBox2.Image).Value = ms.ToArray
                    End If
                    If Rinfosys.CheckBox4.Checked = True Then
                        .Parameters.AddWithValue("@S", 1)
                    Else
                        .Parameters.AddWithValue("@S", 0)
                    End If
                    Try
                        conn.Open()
                        cmd.ExecuteNonQuery()
                        conn.Close()
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    Finally
                        If File.Exists(Rinfosys.PictureBox2.ImageLocation) Then
                            File.Delete(Rinfosys.PictureBox2.ImageLocation)
                        Else

                        End If
                    End Try

                End With
               
            End Using
        End Using
    End Function
    Function _REG_NEW() As Integer
        Dim ThisCmd As String = ""

        ' If Rinfosys.CheckBox1.Checked = True Then
        ThisCmd &= "INSERT INTO RFS_PID(ID, DID, RLNAME,RFNAME, RMNAME, REXTNAME, RDB, RPB, RCS, RSEX, RREL, RCZ, RBT, RSADD, RHEA, RSCHADD, RSCHYYYY, ROCCU, ROOFFADD, RPOST, RORG, RORGTYPE, RORGMS, CONNO, MAIL, ALIAS, LRN, PHILHEALTH, SSS, PAGIBIG, PASSPORT, ISVOTER, VPRCNCT, VBARANGAY, VIDNO, DENTRY) VALUES (@A, @B, @C, @D, @E, @F, @G, @H, @I, @J, @K, @L, @M, @N, @O, @P, @Q, @R, @S, @T, @U, @V, @W, @X, @Y, @Z, @BA, @BB, @BC, @BD, @BE, @BF, @BG, @BH, @BI, @BJ)"
        '  Else
        '    ThisCmd &= "INSERT INTO RFS_PID(ID, DID, RLNAME,RFNAME, RMNAME, REXTNAME, RDB, RPB, RCS, RSEX, RREL, RCZ, RBT, RSADD, CONNO, MAIL, ALIAS, LRN, PHILHEALTH, SSS, PAGIBIG, PASSPORT, ISVOTER, VPRCNCT, VBARANGAY, VIDNO, DENTRY) VALUES (@A, @B, @C, @D, @E, @F, @G, @H, @I, @J, @K, @L, @M, @N, @O, @P, @Q, @R, @S, @T, @U, @V, @Wyn, @X, @Y, @Z, @BA)"
        'End If

        Using conn As New SqlConnection(ConnStr)
            Using cmd As New SqlCommand
                With cmd
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = ThisCmd
                    .Parameters.AddWithValue("@A", Rinfosys.IDno.Text)
                    .Parameters.AddWithValue("@B", Rinfosys.DIDno.Text)
                    .Parameters.AddWithValue("@C", Rinfosys.TextBox1.Text)
                    .Parameters.AddWithValue("@D", Rinfosys.TextBox2.Text)
                    .Parameters.AddWithValue("@E", Rinfosys.TextBox3.Text)
                    .Parameters.AddWithValue("@F", Rinfosys.TextBox4.Text)
                    .Parameters.AddWithValue("@G", Rinfosys.RDB1.Text)
                    .Parameters.AddWithValue("@H", Rinfosys.TextBox5.Text)
                    .Parameters.AddWithValue("@I", Trim(Rinfosys.ComboBox1.Text))
                    .Parameters.AddWithValue("@J", Trim(Rinfosys.ComboBox2.Text))
                    .Parameters.AddWithValue("@K", Trim(Rinfosys.ComboBox3.Text))
                    .Parameters.AddWithValue("@L", Trim(Rinfosys.ComboBox4.Text))
                    .Parameters.AddWithValue("@M", Trim(Rinfosys.ComboBox5.Text))
                    .Parameters.AddWithValue("@N", Rinfosys.TextBox14.Text)
                    'True
                    .Parameters.AddWithValue("@O", Rinfosys.ContactNo.Text)
                    .Parameters.AddWithValue("@P", Rinfosys.TextBox9.Text)
                    .Parameters.AddWithValue("@Q", Rinfosys.DateTimePicker2.Text)
                    .Parameters.AddWithValue("@R", Rinfosys.TextBox13.Text)
                    .Parameters.AddWithValue("@S", Rinfosys.TextBox12.Text)
                    .Parameters.AddWithValue("@T", Rinfosys.TextBox11.Text)
                    .Parameters.AddWithValue("@U", Rinfosys.TextBox16.Text)
                    .Parameters.AddWithValue("@V", Trim(Rinfosys.ComboBox6.Text))
                    .Parameters.AddWithValue("@W", Trim(Rinfosys.ComboBox7.Text))
                    'Continue

                    .Parameters.AddWithValue("@X", Rinfosys.ContactNo.Text)
                    .Parameters.AddWithValue("@Y", Rinfosys.TextBox27.Text)
                    .Parameters.AddWithValue("@Z", Rinfosys.TextBox41.Text)
                    .Parameters.AddWithValue("@BA", Rinfosys.LRNNo.Text)
                    .Parameters.AddWithValue("@BB", Rinfosys.PhilHealthNo.Text)
                    .Parameters.AddWithValue("@BC", Rinfosys.SSSno.Text)
                    .Parameters.AddWithValue("@BD", Rinfosys.PagibigNo.Text)
                    .Parameters.AddWithValue("@BE", Rinfosys.TextBox45.Text)
                    If Rinfosys.VoterCheckBox.Checked = True Then
                        .Parameters.AddWithValue("@BF", "YES")
                    Else
                        .Parameters.AddWithValue("@BF", "NO")
                    End If
                    .Parameters.AddWithValue("@BG", Rinfosys.TextBox44.Text)
                    .Parameters.AddWithValue("@BH", Rinfosys.TextBox43.Text)
                    .Parameters.AddWithValue("@BI", Rinfosys.TextBox42.Text)
                    .Parameters.AddWithValue("@BJ", DateTime.Now.ToString("MM/dd/yyyy"))
                End With

                Try

                    If MsgBox("Control No. " & Rinfosys.IDno.Text & " has been recorded! Do you want to add new?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes Then
                        conn.Open()
                        cmd.ExecuteNonQuery()
                        conn.Close()

                    Else
                        Return Nothing
                    End If

                Catch ex As Exception
                    MsgBox("System cannot accomodate data for some reason. Please try again." & ex.Message)
                End Try

            End Using
        End Using


    End Function

    Function _DELETE() As Single
        Dim r = Rinfosys.DataGridView1
        Using conn As New SqlConnection(ConnStr)
            Using cmd As New SqlCommand("DELETE FROM RFS_PID WHERE ID=@A", conn)
                With cmd
                    .Parameters.AddWithValue("@A", r.SelectedRows(0).Cells(1).Value)
                End With
                Try
                    If MsgBox("Are you deleting this profile?", MsgBoxStyle.YesNo) = DialogResult.Yes Then
                        conn.Open()
                        cmd.ExecuteNonQuery()
                        conn.Close()
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End Using
        End Using
    End Function

    Function _TRUNCATE(ByVal theTable As String) As Single
        Using conn As New SqlConnection(ConnStr)
            Using cmd As New SqlCommand("TRUNCATE TABLE " & theTable & "", conn)
                With cmd
                    If MsgBox("Do you really want to remove all data from this table? Think about it.", MsgBoxStyle.YesNo) Then
                        Try
                            conn.Open()
                            .ExecuteNonQuery()
                            conn.Close()
                        Catch ex As Exception
                            MsgBox("Unable to clear database " & theTable & ".", MsgBoxStyle.Critical)
                        End Try

                    End If
                End With
            End Using
        End Using

    End Function
    Function _dgvSearch() As Single

    End Function
    Function _TEXTSEARCH() As Single
        Try

            Dim conn As New SqlConnection(ConnStr)
            Dim cmd As New SqlCommand("SELECT RLNAME, RFNAME, RMNAME FROM RFS_PID WHERE ID = '" & Rinfosys.SearchBox.Text & "'", conn)
            conn.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader()
            Select Case reader.Read()
                Case True
                    Rinfosys.Sresult.ForeColor = System.Drawing.Color.Green
                    Rinfosys.SearchBox.BackColor = System.Drawing.Color.PaleGreen
                    Rinfosys.TableLayoutPanel3.Show()
                    Rinfosys.Sresult.Text = "RESULT: " & Convert.ToString(reader("RLNAME").ToUpper & ", " & reader("RFNAME") & " " & reader("RMNAME"))
                Case False
                    Rinfosys.Sresult.ForeColor = System.Drawing.Color.Red
                    Rinfosys.SearchBox.BackColor = System.Drawing.Color.LightCoral
                    Rinfosys.TableLayoutPanel3.Hide()
                    Rinfosys.Sresult.Text = "No Data Exist"
                    If Rinfosys.SearchBox.Text = "" Then
                        Rinfosys.SearchBox.BackColor = System.Drawing.Color.White
                        Rinfosys.Sresult.ForeColor = System.Drawing.Color.Black
                        Rinfosys.Sresult.Text = "(Enter Resident Control No. <e.g. CN-12345678> )"
                    End If
            End Select

        Catch ex As Exception

            Rinfosys.Sresult.Text = ""

        Finally
            conn.Close()

        End Try
    End Function
    Function ___sprint() As Single

        Dim thisCmd = "SELECT * FROM RFS_PID WHERE ID = '@DP'"
        Using conn As New SqlConnection(ConnStr)
            Using cmd As New SqlCommand
                With cmd
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = thisCmd
                    .Parameters.AddWithValue("@PD", Rinfosys.SearchBox.Text)

                End With
                Try
                    conn.Open()
                    cmd.ExecuteNonQuery()
                    conn.Close()
                Catch ex As Exception

                End Try
            End Using
        End Using
    End Function

    Function __NEWBUS() As Single
        Dim r = Rinfosys.prtBusiness
        Using conn As New SqlConnection(ConnStr)
            Using cmd As New SqlCommand("INSERT INTO RFS_BUS(ID, DID, ANAME, BNAME, INAME, BPLACE, NATURE, TIN, BOR, PAID, DATE, REMARKS) VALUES(@A, @B, @C, @D, @E, @F, @G, @H, @I, @J, @K, @L)", conn)
                With cmd.Parameters

                    .AddWithValue("@A", Rinfosys.BCID.Text)
                    .AddWithValue("@B", Rinfosys.TextBox29.Text)
                    .AddWithValue("@C", Rinfosys.BusNameOfApplicant.Text)
                    .AddWithValue("@D", Rinfosys.TextBox18.Text)
                    .AddWithValue("@E", Rinfosys.TextBox22.Text)
                    .AddWithValue("@F", Rinfosys.TextBox20.Text)
                    .AddWithValue("@G", Rinfosys.TextBox21.Text)
                    .AddWithValue("@H", Rinfosys.TextBox23.Text)
                    .AddWithValue("@I", Rinfosys.TextBox24.Text)
                    .AddWithValue("@J", Rinfosys.TextBox25.Text)
                    .AddWithValue("@K", DateTime.Today.ToString("MM-dd-yyyy"))
                    .AddWithValue("@L", Rinfosys.TextBox26.Text)
                    Try
                        conn.Open()
                        cmd.ExecuteNonQuery()
                        If MsgBox("Business has certified to operate in " & My.Settings.BAR & ". Do you want to print Business Certificate?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            r.PrinterSettings = New PrinterSettings
                            Dim w As Integer = 850
                            Dim h As Integer = 1167
                            Dim pd As New PrintPreviewDialog
                            Dim psz As New PaperSize("A4", w, h)
                            r.PrinterSettings.DefaultPageSettings.PaperSize = psz
                            pd.Document = r
                            r.DocumentName = "Print Business Certificate"
                            pd.ShowDialog()
                            With Rinfosys.ListBox1.Items
                                .Add("Print Business Permit: " & Rinfosys.BCID.Text)
                                .Add("EventAction - " & DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
                            End With
                        End If
                        conn.Close()
                    Catch ex As Exception
                        MsgBox("System cannot accomodate the data for some reason. Please try again.", MsgBoxStyle.Information)
                    End Try
                End With
            End Using
        End Using

    End Function
    Function __Religion() As Single
        Try
            Using conn As New SqlConnection(ConnStr)
                Using cmd As New SqlCommand("SELECT COUNT(*) FROM RFS_PID WHERE RREL='ISLAM'", conn)
                    conn.Open()
                    Dim n = cmd.ExecuteScalar
                    Rinfosys.ReligionChart.Series(0).Points.AddY(n)

                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) FROM RFS_PID WHERE RREL='CATHOLIC'", conn)
                    conn.Open()
                    Dim n = cmd.ExecuteScalar
                    Rinfosys.ReligionChart.Series(1).Points.AddY(n)
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) FROM RFS_PID WHERE RREL='IGLESIA'", conn)
                    conn.Open()
                    Dim n = cmd.ExecuteScalar
                    Rinfosys.ReligionChart.Series(2).Points.AddY(n)
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) FROM RFS_PID WHERE RREL='PROTESTANT'", conn)
                    conn.Open()
                    Dim n = cmd.ExecuteScalar
                    Rinfosys.ReligionChart.Series(3).Points.AddY(n)
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) FROM RFS_PID WHERE RREL='BUDDHIST'", conn)
                    conn.Open()
                    Dim n = cmd.ExecuteScalar

                    Rinfosys.ReligionChart.Series(4).Points.AddY(n)
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) FROM RFS_PID WHERE RREL='OTHER'", conn)
                    conn.Open()
                    Dim n = cmd.ExecuteScalar
                    Rinfosys.ReligionChart.Series(5).Points.AddY(n)

                    conn.Close()
                End Using
            End Using
        Catch ex As Exception

        End Try
    End Function
    Function __HousingMaterials() As Single
        Try
            Using conn As New SqlConnection(ConnStr)
                Using cmd As New SqlCommand("SELECT COUNT(HOMAT) FROM RFS_HHM WHERE HOMAT='CONCRETE'", conn)
                    conn.Open()
                    Dim n = cmd.ExecuteScalar
                    Rinfosys.HousingChart.Series("Housing").Points(0).SetValueXY(n, n)
                    Rinfosys.HousingChart.Series("Housing").Points(0).Label = n
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(HOMAT) FROM RFS_HHM WHERE HOMAT='SEMI-CONCRETE'", conn)
                    conn.Open()
                    Dim n = cmd.ExecuteScalar
                    Rinfosys.HousingChart.Series("Housing").Points(1).SetValueXY(n, n)
                    Rinfosys.HousingChart.Series("Housing").Points(1).Label = n
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(HOMAT) FROM RFS_HHM WHERE HOMAT='LIGHT'", conn)
                    conn.Open()
                    Dim n = cmd.ExecuteScalar
                    Rinfosys.HousingChart.Series("Housing").Points(2).SetValueXY(n, n)
                    Rinfosys.HousingChart.Series("Housing").Points(2).Label = n
                    conn.Close()
                End Using
            End Using
        Catch ex As Exception

        Finally
            Using conn As New SqlConnection(ConnStr)
                Using cmd As New SqlCommand("SELECT COUNT(*) FROM RFS_PID", conn)
                    With cmd
                        conn.Open()
                        Dim nn As Integer = cmd.ExecuteScalar
                        If nn > 0 Then
                            Dim n = Rinfosys.HousingChart.Series("Housing").Points(0).XValue + Rinfosys.HousingChart.Series("Housing").Points(1).XValue + Rinfosys.HousingChart.Series("Housing").Points(2).XValue
                            For i As Integer = 0 To Rinfosys.HousingChart.Series("Housing").Points.Count - 1
                                Dim x = Rinfosys.HousingChart.Series("Housing").Points(i).XValue
                                Dim ff = FormatPercent(x / n, -1)
                                Rinfosys.HousingChart.Series("Housing").Points(i).Label = ff
                            Next i
                        Else
                            MsgBox("Empty Pie: ", MsgBoxStyle.Information)
                        End If
                        conn.Close()
                    End With
                End Using
            End Using
        End Try
        
    End Function
    Function __POPULATION() As Single
        Try
            Dim r = Rinfosys.Chart1
            '1-10
            Using conn As New SqlConnection(ConnStr)
                Using cmd As New SqlCommand("SELECT COUNT(*)FROM RFS_PID WHERE (DATEDIFF(yy,RDB, GETDATE()) <= 10) AND RSEX='MALE'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Dim d = "1-10"
                        r.Series("Men").XValueType = ChartValueType.DateTime
                        r.Series("Men").Points.AddXY(d, n)
                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) FROM RFS_PID WHERE (DATEDIFF(yy,RDB, GETDATE()) <= 10) AND RSEX='FEMALE'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Dim d = "1-10"
                        r.Series("Women").XValueType = ChartValueType.DateTime
                        r.Series("Women").Points.AddXY(d, n)
                    End With
                    conn.Close()
                End Using
                '11-20
                Using cmd As New SqlCommand("SELECT COUNT(*)FROM RFS_PID WHERE (DATEDIFF(yy,RDB, GETDATE()) >= 11 AND DATEDIFF(yy,RDB, GETDATE()) <= 20) AND RSEX='MALE'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Dim d = "11-20"
                        r.Series("Men").XValueType = ChartValueType.DateTime
                        r.Series("Men").Points.AddXY(d, n)

                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) FROM RFS_PID WHERE (DATEDIFF(yy,RDB, GETDATE()) >= 11 AND DATEDIFF(yy,RDB, GETDATE()) <= 20) AND RSEX='FEMALE'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Dim d = "11-20"
                        r.Series("Women").XValueType = ChartValueType.DateTime
                        r.Series("Women").Points.AddXY(d, n)
                    End With
                    conn.Close()
                End Using
                '20-30
                Using cmd As New SqlCommand("SELECT COUNT(*)FROM RFS_PID WHERE (DATEDIFF(yy,RDB, GETDATE()) >= 21 AND DATEDIFF(yy,RDB, GETDATE()) <= 30) AND RSEX='MALE'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Dim d = "21-30"
                        r.Series("Men").XValueType = ChartValueType.DateTime
                        r.Series("Men").Points.AddXY(d, n)

                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) FROM RFS_PID WHERE (DATEDIFF(yy,RDB, GETDATE()) >= 21 AND DATEDIFF(yy,RDB, GETDATE()) <= 30) AND RSEX='FEMALE'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Dim d = "21-30"
                        r.Series("Women").XValueType = ChartValueType.DateTime
                        r.Series("Women").Points.AddXY(d, n)
                    End With
                    conn.Close()
                End Using
                '31-40
                Using cmd As New SqlCommand("SELECT COUNT(*)FROM RFS_PID WHERE (DATEDIFF(yy,RDB, GETDATE()) >= 31 AND DATEDIFF(yy,RDB, GETDATE()) <= 50) AND RSEX='MALE'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Dim d = "31-40"
                        r.Series("Men").XValueType = ChartValueType.DateTime
                        r.Series("Men").Points.AddXY(d, n)

                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) FROM RFS_PID WHERE (DATEDIFF(yy,RDB, GETDATE()) >= 31 AND DATEDIFF(yy,RDB, GETDATE()) <= 40) AND RSEX='FEMALE'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Dim d = "31-40"
                        r.Series("Women").XValueType = ChartValueType.DateTime
                        r.Series("Women").Points.AddXY(d, n)
                    End With
                    conn.Close()
                End Using
                '41-58
                Using cmd As New SqlCommand("SELECT COUNT(*)FROM RFS_PID WHERE (DATEDIFF(yy,RDB, GETDATE()) >= 41 AND DATEDIFF(yy,RDB, GETDATE()) <= 57) AND RSEX='MALE'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Dim d = "41-57"
                        r.Series("Men").XValueType = ChartValueType.DateTime
                        r.Series("Men").Points.AddXY(d, n)

                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) FROM RFS_PID WHERE (DATEDIFF(yy,RDB, GETDATE()) >= 41 AND DATEDIFF(yy,RDB, GETDATE()) <= 57) AND RSEX='FEMALE'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Dim d = "41-57"
                        r.Series("Women").XValueType = ChartValueType.DateTime
                        r.Series("Women").Points.AddXY(d, n)
                    End With
                    conn.Close()
                End Using
                '58++
                Using cmd As New SqlCommand("SELECT COUNT(*)FROM RFS_PID WHERE (DATEDIFF(yy,RDB, GETDATE()) >= 58) AND RSEX='MALE'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Dim d = "58++"
                        r.Series("Men").XValueType = ChartValueType.DateTime
                        r.Series("Men").Points.AddXY(d, n)

                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) FROM RFS_PID WHERE (DATEDIFF(yy,RDB, GETDATE()) >= 56) <= 58) AND RSEX='FEMALE'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Dim d = "58++"
                        r.Series("Women").XValueType = ChartValueType.DateTime
                        r.Series("Women").Points.AddXY(d, n)
                    End With
                    conn.Close()
                End Using
            End Using
        Catch ex As Exception
            'MsgBox("Error loading chart! " & ex.Message)
        Finally


        End Try

    End Function
    Function _Residency() As Single
        Try
            Using conn As New SqlConnection(ConnStr)

                'Residency
                Using cmd As New SqlCommand("SELECT COUNT(RBT) FROM RFS_PID WHERE RBT='SINGLE-FAMILY'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Rinfosys.ResidencyChart.Series("Residency").Points(0).SetValueXY(n, n)
                        Rinfosys.ResidencyChart.Series("Residency").Points(0).Label = n
                        Rinfosys.ResidencyChart.Series("Residency").Points(0).LegendText = "Single-Family"
                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(RBT) FROM RFS_PID WHERE RBT='APARTMENT'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Rinfosys.ResidencyChart.Series("Residency").Points(1).SetValueXY(n, n)
                        Rinfosys.ResidencyChart.Series("Residency").Points(1).Label = n
                        Rinfosys.ResidencyChart.Series("Residency").Points(1).LegendText = "Apartment"
                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(RBT) FROM RFS_PID WHERE RBT='CONDOMINIUM'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Rinfosys.ResidencyChart.Series("Residency").Points(2).SetValueXY(n, n)
                        Rinfosys.ResidencyChart.Series("Residency").Points(2).Label = n
                        Rinfosys.ResidencyChart.Series("Residency").Points(2).LegendText = "Condo Unit"
                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(RBT) FROM RFS_PID WHERE RBT='SHARED'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Rinfosys.ResidencyChart.Series("Residency").Points(3).SetValueXY(n, n)
                        Rinfosys.ResidencyChart.Series("Residency").Points(3).Label = n
                        Rinfosys.ResidencyChart.Series("Residency").Points(3).LegendText = "Sharer"
                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(RBT) FROM RFS_PID WHERE RBT='BOARDER'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Rinfosys.ResidencyChart.Series("Residency").Points(4).SetValueXY(n, n)
                        Rinfosys.ResidencyChart.Series("Residency").Points(4).Label = n
                        Rinfosys.ResidencyChart.Series("Residency").Points(4).LegendText = "Boarder"
                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(RBT) FROM RFS_PID WHERE RBT='LANDOWNER'", conn)
                    conn.Open()
                    With cmd
                        Dim n = cmd.ExecuteScalar
                        Rinfosys.ResidencyChart.Series("Residency").Points(5).SetValueXY(n, n)

                        Rinfosys.ResidencyChart.Series("Residency").Points(5).Label = n
                        Rinfosys.ResidencyChart.Series("Residency").Points(5).LegendText = "Landowner"
                    End With
                    conn.Close()
                End Using
                'END RESIDENCY CHART
            End Using
        Catch ex As Exception
        Finally
            Using conn As New SqlConnection(ConnStr)
                Using cmd As New SqlCommand("SELECT COUNT(*) FROM RFS_PID", conn)
                    With cmd
                        conn.Open()
                        Dim nn As Integer = cmd.ExecuteScalar
                        If nn > 0 Then
                            Dim n = Rinfosys.ResidencyChart.Series("Residency").Points(0).XValue + Rinfosys.ResidencyChart.Series("Residency").Points(1).XValue + Rinfosys.ResidencyChart.Series("Residency").Points(2).XValue + Rinfosys.ResidencyChart.Series("Residency").Points(3).XValue + Rinfosys.ResidencyChart.Series("Residency").Points(4).XValue + Rinfosys.ResidencyChart.Series("Residency").Points(5).XValue
                            For i As Integer = 0 To Rinfosys.ResidencyChart.Series("Residency").Points.Count - 1
                                Dim x = Rinfosys.ResidencyChart.Series("Residency").Points(i).XValue
                                Dim ff = FormatPercent(x / n, -1)
                                Rinfosys.ResidencyChart.Series("Residency").Points(i).Label = ff
                            Next i
                        Else
                            MsgBox("No data exist in the database ", MsgBoxStyle.Information)
                        End If
                        conn.Close()
                    End With
                End Using
            End Using
        End Try
       
    End Function
    Function __GrowthRates() As Single
        Try
            Using conn As New SqlConnection(ConnStr)
                'Growth Rates
                Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE YEAR(DENTRY)<='" & DateTime.Now.AddYears(-5).ToString("yyyy") & "'", conn)
                    conn.Open()
                    With cmd
                        Dim y = cmd.ExecuteScalar
                        Rinfosys.Chart2.Series("Population").Points(0).SetValueY(y)
                        Rinfosys.Chart2.Series("Population").Points(0).AxisLabel = DateTime.Now.AddYears(-5).ToString("yyyy")
                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE YEAR(DENTRY) <='" & DateTime.Now.AddYears(-4).ToString("yyyy") & "'", conn)
                    conn.Open()
                    With cmd
                        Dim y = cmd.ExecuteScalar
                        Rinfosys.Chart2.Series("Population").Points(1).SetValueY(y)
                        Rinfosys.Chart2.Series("Population").Points(1).AxisLabel = DateTime.Now.AddYears(-4).ToString("yyyy")
                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE YEAR(DENTRY) <='" & DateTime.Now.AddYears(-3).ToString("yyyy") & "'", conn)
                    conn.Open()
                    With cmd
                        Dim y = cmd.ExecuteScalar

                        Rinfosys.Chart2.Series("Population").Points(2).SetValueY(y)
                        Dim past = Rinfosys.Chart2.Series("Population").Points(1).YValues(0)
                        Rinfosys.Chart2.Series("Population").Points(2).AxisLabel = DateTime.Now.AddYears(-3).ToString("yyyy")

                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE YEAR(DENTRY) <='" & DateTime.Now.AddYears(-2).ToString("yyyy") & "'", conn)
                    conn.Open()
                    With cmd
                        Dim y = cmd.ExecuteScalar
                        Rinfosys.Chart2.Series("Population").Points(3).SetValueY(y)
                        Dim past = Rinfosys.Chart2.Series("Population").Points(2).YValues(0)
                        Rinfosys.Chart2.Series("Population").Points(3).AxisLabel = DateTime.Now.AddYears(-2).ToString("yyyy")
                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE YEAR(DENTRY) <='" & DateTime.Now.AddYears(-1).ToString("yyyy") & "'", conn)
                    conn.Open()
                    With cmd
                        Dim y = cmd.ExecuteScalar
                        Rinfosys.Chart2.Series("Population").Points(4).SetValueY(y)
                        Rinfosys.Chart2.Series("Population").Points(4).AxisLabel = DateTime.Now.AddYears(-1).ToString("yyyy")
                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE YEAR(DENTRY) <='" & DateTime.Now.AddYears(0).ToString("yyyy") & "'", conn)
                    conn.Open()
                    With cmd
                        Dim y = cmd.ExecuteScalar

                        Rinfosys.Chart2.Series("Population").Points(5).AxisLabel = DateTime.Today.ToString("yyyy")
                        Dim past = Rinfosys.Chart2.Series("Population").Points(4).YValues(0)
                        ' Dim pp = (y - past)
                        'Dim dv = pp / past
                        'Dim tr = dv * 100
                        'Dim final = tr / 100
                        'Rinfosys.Chart2.Series("Population").Points(5).Label = y
                        Rinfosys.Chart2.Series("Population").Points(5).SetValueY(y)
                        'Rinfosys.Chart2.Series("Population").Points(5).LabelForeColor = System.Drawing.Color.RoyalBlue
                    End With
                    conn.Close()
                End Using
            End Using
        Catch ex As Exception
        Finally 'Get Growth Rates
            Using conn As New SqlConnection(ConnStr)
                'Growth Rates
                Using cmd As New SqlCommand("SELECT SUM(CASE WHEN DENTRY<=DATEADD(YEAR, -5, GETDATE()) THEN 1 ELSE 0 END) AS CURPOP, SUM(CASE WHEN DENTRY<=DATEADD(YEAR, -10,GETDATE()) THEN 1 ELSE 0 END) AS LFY FROM RFS_PID", conn)
                    conn.Open()
                    Dim reader As SqlDataReader = cmd.ExecuteReader
                    While reader.Read()
                        Dim minus = (reader("CURPOP") / reader("LFY")) ^ (1 / 10) - 1
                        'Dim divide = minus / reader("LFY")
                        Dim mults = minus * 100
                        'Dim final = mults / 5
                        Rinfosys.Chart2.Series("Growth Rates").Points(0).SetValueY(mults)
                        Rinfosys.Chart2.Series("Growth Rates").Points(0).Label = Math.Round(mults, 1, MidpointRounding.ToEven) & "%"
                        Rinfosys.Chart2.Series("Growth Rates").Points(0).AxisLabel = DateTime.Now.AddYears(-5).ToString("yyyy")
                    End While
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT SUM(CASE WHEN DENTRY<=DATEADD(YEAR, -4, GETDATE()) THEN 1 ELSE 0 END) AS CURPOP, SUM(CASE WHEN DENTRY<=DATEADD(YEAR, -5, GETDATE()) THEN 1 ELSE 0 END) AS LFY FROM RFS_PID", conn)
                    conn.Open()
                    Dim reader As SqlDataReader = cmd.ExecuteReader
                    While reader.Read()
                        Dim minus = (reader("CURPOP") / reader("LFY")) ^ (1 / 1) - 1
                        'Dim divide = minus / reader("LFY")
                        Dim mults = minus * 100
                        'Dim final = mults / 5
                        Rinfosys.Chart2.Series("Growth Rates").Points(1).SetValueY(mults)
                        Rinfosys.Chart2.Series("Growth Rates").Points(1).Label = Math.Round(mults, 1, MidpointRounding.ToEven) & "%"
                        Rinfosys.Chart2.Series("Growth Rates").Points(1).AxisLabel = DateTime.Now.AddYears(-4).ToString("yyyy")
                    End While
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT SUM(CASE WHEN DENTRY<=DATEADD(YEAR, -3, GETDATE()) THEN 1 ELSE 0 END) AS CURPOP, SUM(CASE WHEN DENTRY<=DATEADD(YEAR, -4, GETDATE()) THEN 1 ELSE 0 END) AS LFY FROM RFS_PID", conn)
                    conn.Open()
                    Dim reader As SqlDataReader = cmd.ExecuteReader
                    While reader.Read()
                        Dim minus = (reader("CURPOP") / reader("LFY")) ^ (1 / 1) - 1
                        ' Dim divide = minus / reader("LFY")
                        Dim mults = minus * 100
                        ' Dim final = mults / 5
                        Rinfosys.Chart2.Series("Growth Rates").Points(2).SetValueY(minus)
                        Rinfosys.Chart2.Series("Growth Rates").Points(2).Label = Math.Round(minus, 1, MidpointRounding.ToEven) & "%"
                        Rinfosys.Chart2.Series("Growth Rates").Points(2).AxisLabel = DateTime.Now.AddYears(-3).ToString("yyyy")

                    End While
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT SUM(CASE WHEN DENTRY<=DATEADD(YEAR, -2, GETDATE()) THEN 1 ELSE 0 END) AS CURPOP, SUM(CASE WHEN DENTRY<=DATEADD(YEAR, -3, GETDATE()) THEN 1 ELSE 0 END) AS LFY FROM RFS_PID", conn)
                    conn.Open()
                    Dim reader As SqlDataReader = cmd.ExecuteReader
                    While reader.Read()
                        Dim minus = (reader("CURPOP") / reader("LFY")) ^ (1 / 3) - 1
                        ' Dim divide = minus / reader("LFY")
                        Dim mults = minus * 100
                        'Dim final = mults / 5
                        Rinfosys.Chart2.Series("Growth Rates").Points(3).SetValueY(mults)
                        Rinfosys.Chart2.Series("Growth Rates").Points(3).Label = Math.Round(mults, 1, MidpointRounding.ToEven) & "%"
                        Rinfosys.Chart2.Series("Growth Rates").Points(3).AxisLabel = DateTime.Now.AddYears(-2).ToString("yyyy")
                    End While
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT SUM(CASE WHEN DENTRY<=DATEADD(YEAR, -1, GETDATE()) THEN 1 ELSE 0 END) AS CURPOP, SUM(CASE WHEN DENTRY<=DATEADD(YEAR, -2, GETDATE()) THEN 1 ELSE 0 END) AS LFY FROM RFS_PID", conn)
                    conn.Open()
                    Dim reader As SqlDataReader = cmd.ExecuteReader
                    While reader.Read()
                        Dim minus = (reader("CURPOP") / reader("LFY")) ^ (1 / 5) - 1
                        'Dim divide = minus / reader("LFY")
                        Dim mults = minus * 100
                        'Dim final = mults / 5
                        Rinfosys.Chart2.Series("Growth Rates").Points(4).SetValueY(mults)
                        Rinfosys.Chart2.Series("Growth Rates").Points(4).Label = Math.Round(mults, 1, MidpointRounding.ToEven) & "%"
                        Rinfosys.Chart2.Series("Growth Rates").Points(4).AxisLabel = DateTime.Now.AddYears(-1).ToString("yyyy")
                    End While
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) AS CURPOP, SUM(CASE WHEN DENTRY<=DATEADD(YEAR, -1, GETDATE()) THEN 1 ELSE 0 END) AS LFY FROM RFS_PID", conn)
                    conn.Open()
                    Dim reader As SqlDataReader = cmd.ExecuteReader
                    While reader.Read()
                        Dim minus = (reader("CURPOP") / reader("LFY")) ^ (1 / 5) - 1
                        'Dim divide = minus / reader("LFY")
                        Dim mults = minus * 100
                        '  Dim final = mults / 5
                        Rinfosys.Chart2.Series("Growth Rates").Points(5).SetValueY(mults)
                        Rinfosys.Chart2.Series("Growth Rates").Points(5).Label = Math.Round(mults, 1, MidpointRounding.ToEven) & "%"
                        Rinfosys.Chart2.Series("Growth Rates").Points(5).AxisLabel = DateTime.Today.ToString("yyyy")

                    End While
                    conn.Close()
                End Using
            End Using
        End Try
    End Function
    Function __Last5Pop() As Single
        Dim r = Rinfosys.Chart1
        Using conn As New SqlConnection(ConnStr)
            Dim d = "1-4"
            Using cmd As New SqlCommand("SELECT COUNT(*), DATEIFF(YEAR, 4, RDB) AS BD FROM RFS_PID WHERE RSEX='MALE' AND YEAR(DENTRY) <='" & DateTime.Now.AddYears(-5).ToString("yyyy") & "'", conn)
                conn.Open()
                With cmd
                    Dim n = cmd.ExecuteScalar
                    r.Series("Men").XValueType = ChartValueType.DateTime
                    r.Series("Men").Points.AddXY(d, n)
                End With
                conn.Close()
            End Using
            Using cmd As New SqlCommand("SELECT COUNT(*), DATEIFF(YEAR, 4, RDB) THEN 1 ELSE 0 END) AS BD FROM RFS_PID WHERE RSEX='FEMALE' AND YEAR(DENTRY) <='" & DateTime.Now.AddYears(-5).ToString("yyyy") & "'", conn)
                conn.Open()
                With cmd
                    Dim n = "1-4"
                    r.Series("Women").XValueType = ChartValueType.DateTime
                    r.Series("Women").Points.AddXY(d, n)
                End With
                conn.Close()
            End Using
            
        End Using
    End Function
    Function __MaritalStats() As Single
        Try
            Using conn As New SqlConnection(ConnStr)
                'Marital Status
                Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE RCS='SINGLE'", conn)
                    conn.Open()
                    With cmd
                        Dim y = cmd.ExecuteScalar
                        Rinfosys.MaritalChart.Series("Marital").Points(0).SetValueXY(y, y)
                        Rinfosys.MaritalChart.Series("Marital").Points(0).Label = y
                        Rinfosys.MaritalChart.Series("Marital").Points(0).AxisLabel = DateTime.Now.AddYears(0).ToString("yyyy")
                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE RCS='MARRIED'", conn)
                    conn.Open()
                    With cmd
                        Dim y = cmd.ExecuteScalar
                        Rinfosys.MaritalChart.Series("Marital").Points(1).SetValueXY(y, y)
                        Rinfosys.MaritalChart.Series("Marital").Points(1).Label = y
                        Rinfosys.MaritalChart.Series("Marital").Points(1).AxisLabel = DateTime.Now.AddYears(0).ToString("yyyy")
                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE RCS='DIVORCED'", conn)
                    conn.Open()
                    With cmd
                        Dim y = cmd.ExecuteScalar
                        Rinfosys.MaritalChart.Series("Marital").Points(2).SetValueXY(y, y)
                        Rinfosys.MaritalChart.Series("Marital").Points(2).Label = y
                        Rinfosys.MaritalChart.Series("Marital").Points(2).AxisLabel = DateTime.Now.AddYears(0).ToString("yyyy")
                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE RCS='WIDOW'", conn)
                    conn.Open()
                    With cmd
                        Dim y = cmd.ExecuteScalar
                        Rinfosys.MaritalChart.Series("Marital").Points(3).SetValueXY(y, y)
                        Rinfosys.MaritalChart.Series("Marital").Points(3).Label = y
                        Rinfosys.MaritalChart.Series("Marital").Points(3).AxisLabel = DateTime.Now.AddYears(0).ToString("yyyy")
                    End With
                    conn.Close()
                End Using
                Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE RCS='LIVE-IN'", conn)
                    conn.Open()
                    With cmd
                        Dim y = cmd.ExecuteScalar
                        Rinfosys.MaritalChart.Series("Marital").Points(4).SetValueXY(y, y)
                        Rinfosys.MaritalChart.Series("Marital").Points(4).Label = y
                        Rinfosys.MaritalChart.Series("Marital").Points(4).AxisLabel = DateTime.Now.AddYears(0).ToString("yyyy")
                    End With
                    conn.Close()
                End Using
            End Using
        Catch ex As Exception

        Finally
            Using conn As New SqlConnection(ConnStr)
                Using cmd As New SqlCommand("SELECT COUNT(*) FROM RFS_PID", conn)
                    With cmd
                        conn.Open()
                        Dim nn As Integer = cmd.ExecuteScalar
                        If nn > 0 Then
                            Dim n = Rinfosys.MaritalChart.Series("Marital").Points(0).XValue + Rinfosys.MaritalChart.Series("Marital").Points(1).XValue + Rinfosys.MaritalChart.Series("Marital").Points(2).XValue + Rinfosys.MaritalChart.Series("Marital").Points(3).XValue + Rinfosys.MaritalChart.Series("Marital").Points(4).XValue
                            For i As Integer = 0 To Rinfosys.MaritalChart.Series("Marital").Points.Count - 1
                                Dim x = Rinfosys.MaritalChart.Series("Marital").Points(i).XValue
                                Dim ff = FormatPercent(x / n, -1)
                                Rinfosys.MaritalChart.Series("Marital").Points(i).Label = ff
                            Next i
                        Else
                            MsgBox("No data exist in the database ", MsgBoxStyle.Information)
                        End If
                        conn.Close()
                    End With
                End Using
            End Using

        End Try

    End Function
    Function __isDeadRes() As Single
        Using conn As New SqlConnection(ConnStr)
            'Growth Rates Deceased
            Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE ISDEAD='YES' AND YEAR(DENTRY)<='" & DateTime.Now.AddYears(-5).ToString("yyyy") & "'", conn)
                conn.Open()
                With cmd
                    Dim y = cmd.ExecuteScalar
                    Rinfosys.Chart2.Series("Deceased").Points(0).SetValueY(y)
                    Dim pp = y
                    Dim gr = (pp / 100)
                    Dim tr = gr * 100
                    Rinfosys.Chart2.Series("Deceased").Points(0).Label = Math.Round(tr, 2) & "%"
                    Rinfosys.Chart2.Series("Deceased").Points(0).AxisLabel = DateTime.Now.AddYears(-5).ToString("yyyy")
                End With
                conn.Close()
            End Using
            Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE ISDEAD='YES' AND YEAR(DENTRY) <='" & DateTime.Now.AddYears(-4).ToString("yyyy") & "'", conn)
                conn.Open()
                With cmd
                    Dim y = cmd.ExecuteScalar
                    Rinfosys.Chart2.Series("Deceased").Points(1).SetValueY(y)
                    Dim past = Rinfosys.Chart2.Series("Deceased").Points(0).YValues(0)
                    Dim pp = (y - past) / past
                    Dim gr = (pp / 100)
                    Dim tr = gr * 100
                    Rinfosys.Chart2.Series("Deceased").Points(1).Label = Math.Round(tr, 2) & "%"
                    Rinfosys.Chart2.Series("Deceased").Points(1).AxisLabel = DateTime.Now.AddYears(-4).ToString("yyyy")
                End With
                conn.Close()
            End Using
            Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE ISDEAD='YES' AND YEAR(DENTRY) <='" & DateTime.Now.AddYears(-3).ToString("yyyy") & "'", conn)
                conn.Open()
                With cmd
                    Dim y = cmd.ExecuteScalar

                    Rinfosys.Chart2.Series("Deceased").Points(2).SetValueY(y)
                    Dim past = Rinfosys.Chart2.Series("Deceased").Points(1).YValues(0)
                    Dim pp = (y - past) / past
                    Dim gr = (pp / 100)
                    Dim tr = gr * 100
                    Rinfosys.Chart2.Series("Deceased").Points(2).Label = Math.Round(tr, 2) & "%"
                    Rinfosys.Chart2.Series("Deceased").Points(2).AxisLabel = DateTime.Now.AddYears(-3).ToString("yyyy")

                End With
                conn.Close()
            End Using
            Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE ISDEAD='YES' AND YEAR(DENTRY) <='" & DateTime.Now.AddYears(-2).ToString("yyyy") & "'", conn)
                conn.Open()
                With cmd
                    Dim y = cmd.ExecuteScalar
                    Rinfosys.Chart2.Series("Deceased").Points(3).SetValueY(y)
                    Dim past = Rinfosys.Chart2.Series("Deceased").Points(2).YValues(0)
                    Dim pp = (y - past) / past
                    Dim gr = (pp / 100)
                    Dim tr = gr * 100
                    Rinfosys.Chart2.Series("Deceased").Points(3).Label = Math.Round(tr, 2) & "%"
                    Rinfosys.Chart2.Series("Deceased").Points(3).AxisLabel = DateTime.Now.AddYears(-2).ToString("yyyy")
                End With
                conn.Close()
            End Using
            Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE ISDEAD='YES' AND YEAR(DENTRY) <='" & DateTime.Now.AddYears(-1).ToString("yyyy") & "'", conn)
                conn.Open()
                With cmd
                    Dim y = cmd.ExecuteScalar
                    Rinfosys.Chart2.Series("Deceased").Points(4).SetValueY(y)
                    Dim past = Rinfosys.Chart2.Series("Deceased").Points(3).YValues(0)
                    Dim pp = (y - past) / past
                    Dim gr = (pp / 100)
                    Dim tr = gr * 100
                    Rinfosys.Chart2.Series("Deceased").Points(4).Label = Math.Round(tr, 2) & "%"
                    Rinfosys.Chart2.Series("Deceased").Points(4).AxisLabel = DateTime.Now.AddYears(-1).ToString("yyyy")
                End With
                conn.Close()
            End Using
            Using cmd As New SqlCommand("SELECT COUNT(*) AS NUM FROM RFS_PID WHERE ISDEAD='YES' AND YEAR(DENTRY) <='" & DateTime.Now.AddYears(0).ToString("yyyy") & "'", conn)
                conn.Open()
                With cmd
                    Dim y = cmd.ExecuteScalar

                    Rinfosys.Chart2.Series("Deceased").Points(5).SetValueY(y)
                    Rinfosys.Chart2.Series("Deceased").Points(5).AxisLabel = DateTime.Today.ToString("yyyy")
                    Dim past = Rinfosys.Chart2.Series("Deceased").Points(4).YValues(0)
                    Dim pp = (y - past) / past
                    Dim gr = (pp / 100)
                    Dim tr = gr * 100
                    Rinfosys.Chart2.Series("Deceased").Points(5).Label = Math.Round(tr, 2) & "%"
                    Rinfosys.Chart2.Series("Deceased").Points(5).LabelForeColor = System.Drawing.Color.RoyalBlue
                End With
                conn.Close()
            End Using
        End Using
    End Function
    'Education
    Function __EductionStats() As Single
        Using conn As New SqlConnection(ConnStr)
            Using cmd As New SqlCommand("", conn)

            End Using
        End Using
    End Function

    Function __INSERTTOLIST() As Single

        Dim r = Rinfosys
        Using conn As New SqlConnection(ConnStr)
            Using cmd As New SqlCommand("INSERT INTO RFS_EVT(NAME, EVTNAME, EVENTAG, TYPEOFAG, DATEFILE) VALUES (@B, @C, @D, @E, @F)", conn)
                With cmd.Parameters

                    .AddWithValue("@B", r.TextBox7.Text)
                    .AddWithValue("@C", r.ComboBox8.Text)
                    .AddWithValue("@D", r.TextBox39.Text)
                    .AddWithValue("@E", r.TextBox39.Text)
                    .AddWithValue("@F", DateTime.Now.ToString("MM-dd-yyyy"))
                    Try
                        conn.Open()
                        r.Timer1.Interval = 3000
                        r.Timer1.Start()
                        r.StatusTextBox.Show()
                        r.StatusTextBox.ForeColor = System.Drawing.Color.LimeGreen
                        r.StatusTextBox.Text = "Resident Listed Successfully!"
                        cmd.ExecuteNonQuery()
                        conn.Close()
                    Catch ex As Exception
                        r.StatusTextBox.Show()
                        r.Timer1.Interval = 3000
                        r.Timer1.Start()
                        r.StatusTextBox.ForeColor = System.Drawing.Color.Red
                        r.StatusTextBox.Text = "Resident have not listed!"
                    End Try

                End With

            End Using
        End Using


    End Function
    Function SaveAsImage(ByVal filename As String, ByVal image As System.Drawing.Image, ByVal w As Integer, ByVal h As Integer) As Single
        Dim path As String = System.IO.Path.Combine(My.Application.Info.DirectoryPath, filename)

        Using bmp As New Bitmap(w, h)
            Dim gfx As Graphics = Graphics.FromImage(bmp)
            gfx.DrawImage(image, System.Drawing.Point.Empty)
            bmp.Save(filename, System.Drawing.Imaging.ImageFormat.Png)
            gfx.Dispose()
            bmp.Dispose()
        End Using

    End Function
    Function __SAVECHARTSASIMAGE(ByVal cm As ContextMenuStrip, ByVal i As Integer, ByRef sfd As System.Windows.Forms.SaveFileDialog, ByVal ch As Chart) As Single

        With cm
            ch.ContextMenuStrip = cm
            .Items.Add("Save Image")
            sfd.Filter = "PNG (*.png)|*.png"
            sfd.FileName = "Chart_" & DateTime.Now.ToString("MMMM-dd-yyyyy")
            sfd.RestoreDirectory = True
            AddHandler .Items(i).Click, AddressOf sfd.ShowDialog
            If sfd.ShowDialog = DialogResult.OK Then
                ch.SaveImage(sfd.FileName, ChartImageFormat.Png)
            End If

        End With

    End Function
    Function __GETEVENTSBYNAME() As Single
        Dim r = Rinfosys
        Dim d = r.DataGridView3
        Using conn As New SqlConnection(ConnStr)
            Using cmd As New SqlCommand("SELECT EVTNAME FROM RFS_EVT", conn)
                Dim rs As SqlDataReader = cmd.ExecuteReader
                Dim dt As DataTable = New DataTable
                dt.Load(rs)
                r.ComboBox8.ValueMember = "EVTNAME"
                r.ComboBox8.DisplayMember = "EVTNAME"
                r.ComboBox8.DataSource = dt

            End Using
        End Using
    End Function

    Function _insertolog(ByVal name As String, ByVal ctrl As String) As Single
        Using conn As New SqlConnection(ConnStr)
            Using cmd As New SqlCommand("INSERT INTO RFS_LOG(RCTRL,NAME, LD) VALUES(@A, @B, @C)", conn)
                With cmd.Parameters
                    .AddWithValue("@A", ctrl)
                    .AddWithValue("@B", name)
                    .AddWithValue("@C", DateTime.Now.ToString("MM/dd/yyyy"))

                    Try
                        conn.Open()
                        cmd.ExecuteNonQuery()
                        Console.WriteLine("Logged")
                        conn.Close()
                    Catch ex As Exception
                    Finally
                        _LOG()
                    End Try
                End With
            End Using
        End Using
    End Function

    Function _initIdentifier() As Single
      
        If Trim(My.Settings.TOKEN.ToString) = My.Computer.Registry.CurrentUser.OpenSubKey(My.Application.Info.AssemblyName).GetValue(My.Settings.BAR, 1).ToString Then
            Rinfosys.SearchBox.Enabled = True
            Rinfosys.TabPage3.Enabled = True
            Rinfosys.PrintPanel.Enabled = True
            Rinfosys.TabControl2.Enabled = True
            Rinfosys.TabControl3.Enabled = True
            Rinfosys.SplitContainer4.Enabled = True
            Rinfosys.Cursor = System.Windows.Forms.Cursors.Hand
            Rinfosys.Label74.Hide()
            Rinfosys.Button13.Hide()
            Rinfosys.Label75.Text = "User Key: " & My.Settings.TOS.Replace("-", "").ToUpper
        Else
            Rinfosys.SearchBox.Enabled = False
            Rinfosys.TabPage3.Enabled = False
            Rinfosys.PrintPanel.Enabled = False
            Rinfosys.TabControl2.Enabled = False
            Rinfosys.TabControl3.Enabled = False
            Rinfosys.SplitContainer4.Enabled = False
            Rinfosys.TabPage2.Cursor = System.Windows.Forms.Cursors.Hand
            Rinfosys.Cursor = System.Windows.Forms.Cursors.No
            Rinfosys.Label74.Show()
            Rinfosys.Button13.Show()
            Rinfosys.Label75.Text = "Not Activated"
            My.Settings.BP102 = "PROVINCE OF LANAO DEL NORTE"
            My.Settings.BCM103 = "MUNICIPALITY OF PANTAR"
            My.Settings.BAR = "BARANGAY POONA-PUNOD"
        End If

    End Function
    Function _LOG() As Single
        conn.Open()
        Dim com As String = "SELECT * FROM RFS_LOG"
        Dim Adpt As New SqlDataAdapter(com, conn)
        Dim ds As New DataSet()
        Adpt.Fill(ds, "LOG")
        LogData.DataGridView2.DataSource = ds.Tables(0)
        conn.Close()
    End Function
    Function _GetList(ByVal thisEvent As String) As Single
        Using conn As New SqlConnection(ConnStr)
            Using cmd As New SqlCommand("SELECT * FROM RFS_EVT WHERE EVTNAME='" & thisEvent & "'", conn)
                Dim sdr As SqlDataReader = cmd.ExecuteReader
                While sdr.Read()
                    conn.Open()


                    conn.Close()
                End While
            End Using
        End Using
    End Function

    Function _exportDGV(ByRef dgv As DataGridView, ByVal loc As String) As Single

        With dgv
            Dim xlApp As New Excel.Application
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet
            Dim misValue As Object = System.Reflection.Missing.Value
            Dim i, j As Integer
            Dim Data(0 To 100) As String

            xlApp = New Excel.ApplicationClass
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheet = xlWorkBook.Sheets("sheet1")

            For i = 0 To dgv.RowCount - 1
                For j = 0 To dgv.ColumnCount - 1
                    xlWorkSheet.Cells(i + 2, j + 1) = dgv(j, i).Value.ToString()
                Next
            Next
            xlWorkSheet.SaveAs(loc)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()
            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)
            KillExcelProcess()
        End With
    End Function
    Private Sub KillExcelProcess()
        Try
            Dim Xcel() As Process = Process.GetProcessesByName("EXCEL")
            For Each Process As Process In Xcel
                Process.Kill()
            Next
        Catch ex As Exception
        End Try
    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Function __QRCode(ByVal pb As PictureBox, ByVal toEncode As String) As Single
        Dim qrcode As KeepAutomation.Barcode.Bean.BarCode = New KeepAutomation.Barcode.Bean.BarCode
        qrcode.Symbology = KeepAutomation.Barcode.Symbology.QRCode
        qrcode.QRCodeVersion = KeepAutomation.Barcode.QRCodeVersion.V5
        qrcode.QRCodeECL = KeepAutomation.Barcode.QRCodeECL.H
        qrcode.QRCodeDataMode = KeepAutomation.Barcode.QRCodeDataMode.Auto
        qrcode.CodeToEncode = toEncode
        qrcode.BarcodeUnit = KeepAutomation.Barcode.BarcodeUnit.Pixel
        qrcode.X = 2
        qrcode.Y = 2
        qrcode.LeftMargin = 8
        qrcode.RightMargin = 8
        qrcode.TopMargin = 8
        qrcode.BottomMargin = 8
        pb.Image = qrcode.generateBarcodeToBitmap

    End Function

    Function __PrintLists(ByVal NumPages As Integer, ByVal Rpt1 As String, ByVal Rpt2 As String) As Single
        Dim i As Integer
        Dim dv = Rinfosys.DataGridView1
        With dv
            'Set the page number loop and alternate printing the report pages.
            For i = 0 To NumPages
                Dim row = .Rows(i).Cells(i).Value

            Next i
        End With
    End Function

    Function __isDeceased() As Single
        Try
            Using conn As New SqlConnection(ConnStr)
                Using cmd As New SqlCommand("UPDATE RFS_PID SET ISDEAD=@A WHERE ID='" & Rinfosys.DataGridView1.SelectedRows(0).Cells(1).Value & "'", conn)
                    conn.Open()
                    cmd.Parameters.AddWithValue("@A", "YES")
                    cmd.ExecuteNonQuery()
                    conn.Close()
                End Using
            End Using
        Catch ex As Exception
        Finally
            _init_db()
        End Try
      
    End Function
    Function __EconomyChart() As Single
        Try
            Using conn As New SqlConnection(ConnStr)
                Using cmd As New SqlCommand("SELECT SUM(CASE WHEN NATURE='' THEN 1 ELSE 0 END) AS B FROM RFS_BUS", conn)

                End Using
            End Using
        Catch ex As Exception

        End Try
    End Function
    Function __GenInit() As Single
        Dim T = My.Settings.TOKEN
        Dim m = My.Settings.BAR
        Dim a = My.Application.Info.AssemblyName
        Dim n = My.Computer.Registry.CurrentUser.OpenSubKey(a).GetValue(a, T)
        Dim B = My.Computer.Registry.CurrentUser.OpenSubKey(T).GetValue(T, T)


        Dim regKey = n.GetValue(T)

        If String.Equals(m, regKey) = True Then

        End If
    End Function
   
End Module
