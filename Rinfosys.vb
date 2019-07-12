Imports System.Configuration
Imports System.Drawing.Printing
Imports System.Data
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.IO
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Windows.Forms.DataVisualization
Imports Microsoft.Office.Interop
Imports Microsoft.VisualBasic.PowerPacks

Public Class Rinfosys

    Public ConnStr As String = My.Settings.RFS_DBCONN_STR
    Public conn As New SqlConnection(ConnStr)
    Public cmd As New SqlCommand
    Public reader As SqlDataReader
    Public adap As SqlDataAdapter
    Public this As String

    Private Sub EncryptConfigSection()
        Dim Config As Configuration
        Dim Section As ConfigurationSection

        Config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        Section = Config.GetSection("connectionStrings")

        If (Section IsNot Nothing) Then

            If (Not Section.SectionInformation.IsProtected) Then

                If (Not Section.ElementInformation.IsLocked) Then
                    Section.SectionInformation.ProtectSection("DataProtectionConfigurationProvider")
                    Section.SectionInformation.ForceSave = True
                    Config.Save(ConfigurationSaveMode.Full)
                End If
            End If
        End If

    End Sub

    Private Sub Rinfosys_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Me.Dispose()
    End Sub
    Private Sub Rinfosys_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Dim sw As StreamWriter
            Dim fs As FileStream
            Dim strFile As String = System.Environment.CurrentDirectory & "/EL_" & DateTime.Now.ToString("dd-MMM-yyyy") & ".txt"
            If (Not File.Exists(strFile)) Then
                Try
                    fs = File.Create(strFile)
                Catch ex As Exception
                    MsgBox("Error Creating Log File" & ex.Message)
                End Try
            Else
                sw = File.AppendText(strFile)
                For i = 0 To ListBox1.Items.Count - 1
                    sw.WriteLine(DateTime.Now.ToString("MMMM dd, yyyy HH:mm") & ": " & ListBox1.Items(i).ToString)
                Next
                sw.Close()
            End If
            My.Settings.Save()
        Catch ex As Exception

        Finally
            With ListBox1.Items
                .Add("End Session")
                .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
            End With
            _insertolog("Reprint Clearance", SearchBox.Text)

            System.Windows.Forms.Application.Exit()

        End Try
    End Sub

    Private Sub Rinfosys_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Shift + e.KeyCode = Keys.A) Then
            MsgBox("")
        End If
    End Sub
    Private Sub Rinfosys_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        EncryptConfigSection()
        Try
            _APP_UI()
            _init_db()
            __LISTING()

            'Charting 
            __POPULATION()
            __HousingMaterials()
            __GrowthRates()
            _Residency()
            __MaritalStats()
            __Religion()
        Catch ex As Exception

        Finally
            __QRCode(QrCodeBox, IDno.Text)
            _initIdentifier()

            If DataGridView1.RowCount < 0 Then
                Button12.Enabled = False
                Button15.Enabled = False
                Button16.Enabled = False
            Else
                Button12.Enabled = True
                Button15.Enabled = True
                Button16.Enabled = True
            End If
            PictureBox2.ImageLocation = System.Environment.CurrentDirectory + "/Images/profile.jpg"
            Me.Text = Trim(My.Settings.BAR).ToUpper & " | " & My.Application.Info.AssemblyName
            Label38.Text = Trim(My.Settings.BP102).ToUpper
            Label39.Text = Trim(My.Settings.BCM103).ToUpper
            BarangayLabel.Text = Trim(My.Settings.BAR).ToUpper
            SplitContainer1.Panel1.BackColor = System.Drawing.Color.DodgerBlue

            DataGridView1.ContextMenuStrip = RinfosysContextMenu

            Dim ComputerName = System.Net.Dns.GetHostName
            Dim AppProcssID = Process.GetCurrentProcess.Id
            With ListBox1.Items
                .Add("Application PID: " & AppProcssID)
                .Add("Host: " & ComputerName)
                .Add("Application/db: " & My.Settings.RFS_DBCONN_STR.Substring(57))
                .Add("Date: " & DateTime.Today.ToString("MM/dd/yyyy"))
                .Add("Welcome to Rinfosys!")
                .Add("------------xxxxxx------------")
            End With


        End Try


    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Me.Close()
        Me.Dispose()
    End Sub
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            GroupBox2.Enabled = True
        Else
            GroupBox2.Enabled = False
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        __Upload()
        ' Dim opd = New System.Windows.Forms.OpenFileDialog
        'opd.Title = "Select Photo"
        'opd.InitialDirectory = opd.FileName
        'opd.RestoreDirectory = True
        'opd.Filter = "All Images(*.*)|*.jpg;*.png;*.gif;*.bmp;*.tiff|JPEG (*.jpg) |*.jpg|PNG (*.png)|*.png"
        'If opd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
        'PictureBox2.ImageLocation = opd.FileName
        ' PictureBox2.SizeMode = PictureBoxSizeMode.Zoom
        ' ImgURL.Text = opd.FileName
        ' End If
        ' opd.Dispose()
    End Sub
    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        _RESETID()
        __QRCode(QrCodeBox, IDno.Text)
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Try
            
            _REG_NEW()
            __LogEntry()
        Catch ex As Exception

        Finally
            Button12.Enabled = True
            Button15.Enabled = True
            Button16.Enabled = True
            Dim thenewName = TextBox1.Text.ToUpper & ", " & TextBox2.Text.ToUpper & " " & TextBox3.Text.ToUpper
            With ListBox1.Items
                .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") & " > Added data: " & IDno.Text & " > " & thenewName)
                .Add("-----------------------------------")
            End With
            _init_db()
            _RESETID()
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        Select Case True Or False
            Case My.Settings.IDCheckBox = True
                ConfigDiag.CheckBox3.Checked = True
                ConfigDiag.CustomIDPanel.Enabled = True
            Case My.Settings.IDCheckBox = False
                ConfigDiag.CheckBox3.Checked = False
                ConfigDiag.CustomIDPanel.Enabled = False
            Case My.Settings.IDWM.Length > 1
                DehLink.Show()
            Case My.Settings.IDWM.Length <= 0
                DehLink.Hide()

        End Select
        ConfigDiag.TextBox1.Text = My.Settings.CustomConnection
        ConfigDiag.TopMost = True
        ConfigDiag.ShowDialog()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click

        Dim ppsize As System.Drawing.Printing.PaperSize
        Dim pd As New System.Windows.Forms.PrintDialog
        For i = 0 To pd.PrinterSettings.PaperSizes.Count - 1
            ppsize = pd.PrinterSettings.PaperSizes.Item(i)
            Dim PSSelector = PaperSizesBox.SelectedIndex
            PaperSizesBox.Items.Add(ppsize.PaperName)
            If pd.PrinterSettings Is ppsize Then
                PSSelector = i
            End If
        Next
        PrintPanel.Show()
    End Sub


    Private Sub TextBox6_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles SearchBox.KeyUp
        _TEXTSEARCH()
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Select Case True
            Case NewBornCheck.Checked
                Dim pd As New PrintPreviewDialog
                Dim psz As New PaperSize("A4", 827, 1168)
                pd.Document = PrintNewBorn
                PrintNewBorn.DocumentName = "Print New Born Certificate"
                pd.ShowDialog()
                With ListBox1.Items
                    Dim TheInfo = Trim(TextBox2.Text).ToUpper & " " & Trim(TextBox3.Text.First).ToUpper & ". " & Trim(TextBox1.Text).ToUpper
                    .Add("Print Newborn Indorsement: [" & IDno.Text & "]" & TheInfo)
                    .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
                End With
                _insertolog("Printed Newborn Indorsement", IDno.Text)

            Case CheckBox3.Checked
                prtCert.PrinterSettings = New PrinterSettings
                Dim w As Integer = PpWidthSize.Text * 100
                Dim h As Integer = PpHeightSize.Text * 100
                Dim PpName As String = PaperSizesBox.Text
                Dim pd As New PrintPreviewDialog
                Dim psz As New PaperSize(PpName, w, h)
                PrintIndigentPeople.PrinterSettings.DefaultPageSettings.PaperSize = psz
                pd.Document = PrintIndigentPeople
                PrintIndigentPeople.DocumentName = "Print Indigenious People Certificate"
                pd.ShowDialog()
                With ListBox1.Items
                    Dim TheInfo = Trim(TextBox2.Text).ToUpper & " " & Trim(TextBox3.Text.First).ToUpper & ". " & Trim(TextBox1.Text).ToUpper
                    .Add("Print Indigenious People Certificate: [" & IDno.Text & "]" & TheInfo)
                    .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
                End With
                _insertolog("Print Indigenious People Certificate", IDno.Text)
            Case CheckBox4.Checked
                PrintSoloParent.PrinterSettings = New PrinterSettings
                Dim w As Integer = PpWidthSize.Text * 100
                Dim h As Integer = PpHeightSize.Text * 100
                Dim PpName As String = PaperSizesBox.Text
                Dim pd As New PrintPreviewDialog
                Dim psz As New PaperSize(PpName, w, h)
                PrintSoloParent.PrinterSettings.DefaultPageSettings.PaperSize = psz
                pd.Document = PrintSoloParent
                prtCert.DocumentName = "Print Barangay Certificate"
                pd.ShowDialog()
                With ListBox1.Items
                    Dim TheInfo = Trim(TextBox2.Text).ToUpper & " " & Trim(TextBox3.Text.First).ToUpper & ". " & Trim(TextBox1.Text).ToUpper
                    .Add("Print Solo Parent: [" & IDno.Text & "]" & TheInfo)
                    .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
                End With
                _insertolog("Printed Resident Certificate", IDno.Text)

            Case CheckBox5.Checked

                PrintHomeSharer.PrinterSettings = New PrinterSettings
                Dim w As Integer = PpWidthSize.Text * 100
                Dim h As Integer = PpHeightSize.Text * 100
                Dim PpName As String = PaperSizesBox.Text
                Dim pd As New PrintPreviewDialog
                Dim psz As New PaperSize(PpName, w, h)
                PrintHomeSharer.PrinterSettings.DefaultPageSettings.PaperSize = psz
                pd.Document = PrintHomeSharer
                prtCert.DocumentName = "Print Barangay Certificate"
                pd.ShowDialog()
                With ListBox1.Items
                    Dim TheInfo = Trim(TextBox2.Text).ToUpper & " " & Trim(TextBox3.Text.First).ToUpper & ". " & Trim(TextBox1.Text).ToUpper
                    .Add("equested Home Sharer Certificate: [" & IDno.Text & "]" & TheInfo)
                    .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
                End With
                _insertolog("Requested Home Sharer Certificate", IDno.Text)

            Case Else
                Dim pd As New PrintPreviewDialog
                Dim psz As New PaperSize("A4", 827, 1168)
                pd.Document = prtCert
                prtCert.DocumentName = "Print Resident Certificate"
                pd.ShowDialog()
                With ListBox1.Items
                    Dim TheInfo = Trim(TextBox2.Text).ToUpper & " " & Trim(TextBox3.Text.First).ToUpper & ". " & Trim(TextBox1.Text).ToUpper
                    .Add("Print Resident Certificate: [" & IDno.Text & "]" & TheInfo)
                    .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
                End With
                _insertolog("Print Resident Certificate", IDno.Text)


        End Select
      

    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        PrintPanel.Hide()

    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        Try
            Dim pd As New PrintPreviewDialog
            pd.Document = prtCert
            prtCert.PrinterSettings.PrinterName = CStr(PaperSizesBox.SelectedItem)
            pd.ShowDialog()
        Catch ex As Exception
        End Try
    End Sub
    ' PageSettings Object:
    Dim pgsettings As New System.Drawing.Printing.PageSettings
    ' You use a PrintDocument object to print with
    Dim DocPrint As New PrintDocument

    Private Sub prtCert_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles prtCert.PrintPage
        Dim today = DateTime.Today
        Dim day = today.Day
        Dim month = today.ToString("MMMM")
        Dim year = today.Year
        Dim pinF As New PointF(20, 20)

        'Fonts
        Dim RegFont As Font = New Font("Calibri", 12, FontStyle.Regular)
        Dim BoldFont As Font = New Font("Verdana", 12, FontStyle.Bold)
        Dim Title As Font = New Font("Calibri", 16, FontStyle.Bold)
        Dim DocFont As Font = New Font("Verdana", 20, FontStyle.Bold)

        'Page area
        'Dim rect As New Rectangle()

        'Barcode
        Dim barcode As KeepAutomation.Barcode.Bean.BarCode = New KeepAutomation.Barcode.Bean.BarCode
        barcode.Symbology = KeepAutomation.Barcode.Symbology.Codabar
        barcode.DisplayText = True
        barcode.BarcodeUnit = KeepAutomation.Barcode.BarcodeUnit.Cm
        barcode.CodeToEncode = DIDno.Text
        barcode.X = 2
        barcode.Y = 40
        barcode.BarCodeWidth = 100
        barcode.BarCodeHeight = 70
        barcode.DPI = 100

        Dim ThisBarcode As Image = barcode.generateBarcodeToBitmap

        'Text Alignment
        Dim sf As StringFormat = StringFormat.GenericTypographic
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word

        Dim rt As New StringFormat
        rt.FormatFlags = StringFormatFlags.DirectionVertical

        Dim cx, y, w, h As Integer
        cx = 414.5
        y = 20
        w = 827
        h = 1168
        If Trim(My.Settings.CMLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.CMLOGO), 650, 20, 120, 120)
        Else

        End If
        If Trim(My.Settings.BRNGYLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.BRNGYLOGO), 75, 20, 120, 120)

        Else

        End If

        Dim Purpose = TextBox19.Text
        Select Case CheckBox2.Checked
            Case True
                Purpose = TextBox19.Text.ToUpper
            Case False
                Purpose = "APPLICATION"
        End Select

        Dim ThisName = Trim(TextBox2.Text & " " & TextBox3.Text.Substring(0) & " " & TextBox1.Text).ToUpper
        Dim LetterBody = "          THIS IS TO CERTIFY that per records available in this office, " & ThisName & " is a bonafide resident of " & Trim(My.Settings.BAR).ToUpper & ", " & Trim(My.Settings.BCM103).ToUpper & ", " & Trim(My.Settings.BP102).ToUpper & ", Philippines. " & vbCrLf & vbCrLf & _
                      "         It is further certifies that above - named person has no pending and/or derogatory record filed against him/her as of this date." & vbCrLf & vbCrLf & _
                      "         This certification is issued upon the request of the above mentioned person to support him/her for " & Trim(Purpose).ToUpper & " and whatsoever legal purposes that may serve him/her best." & vbCrLf & vbCrLf & _
                      "         Issued this " & day & " day of " & month & ", " & year & " at Barangay Hall, " & Trim(My.Settings.BAR).ToUpper & ", " & Trim(My.Settings.BCM103).ToUpper & ", " & Trim(My.Settings.BP102).ToUpper & ", Philippines."
        Dim HeaderText = "Republic of the Philippines" & vbCrLf & _
                         "Department of the Interior and Local Government" & vbCrLf & _
                         Trim(My.Settings.BP102).ToUpper & vbCrLf & _
                         Trim(My.Settings.BCM103).ToUpper

        Dim Brngy = Trim(My.Settings.BAR).ToUpper
        Dim PresideOffice = "OFFICE OF THE BARANGAY CHAIRMAN"
        Dim ToWhom = "TO WHOM IT MAY CONCERN:"
        Dim sss = "Certified By: " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                    "_____________________" & vbCrLf & _
                     " " & My.Settings.BChair.ToUpper
        Dim Prinfo = "Printed on: " & DateTime.Today.ToShortDateString & " via " & System.Windows.Forms.Application.ProductName & " " & System.Windows.Forms.Application.ProductVersion
        e.Graphics.DrawLine(Pens.Black, 50, 150, 750, 150)

        'Align texts
        e.Graphics.DrawString(HeaderText, RegFont, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 20, w * 0.52, 100), sf)
        e.Graphics.DrawString(Brngy, Title, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 80), sf)
        e.Graphics.DrawString(PresideOffice, Title, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 200), sf)

        e.Graphics.DrawString("CERTIFICATION", DocFont, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 100, w * 0.52, 300), sf)
        e.Graphics.DrawString("PERMANENT RESIDENT", BoldFont, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 120, w * 0.52, 300), sf)

        e.Graphics.DrawString(ToWhom, BoldFont, System.Drawing.Brushes.Black, New Rectangle(75, y * 16, w - 150, 100))
        e.Graphics.DrawString(LetterBody, RegFont, System.Drawing.Brushes.Black, New Rectangle(75, y * 18, w - 150, 450), StringFormat.GenericDefault)

        'Officials' Certification
        e.Graphics.DrawString(sss, BoldFont, System.Drawing.Brushes.Black, 480, 700)

        'Document Identification
        e.Graphics.DrawString("__________________", RegFont, System.Drawing.Brushes.Black, 78, 920)
        e.Graphics.DrawString("Specimen Signature", RegFont, System.Drawing.Brushes.Black, 78, 940)
        e.Graphics.DrawImage(ThisBarcode, 70, 1000, 170, 45)
        'Printed on datetime
        e.Graphics.DrawString(Prinfo, New Font("Calibri", 8, FontStyle.Regular), System.Drawing.Brushes.Gainsboro, New RectangleF(5, 50, 200, 200), rt)
        e.Graphics.DrawString("Doc. No.:" & DIDno.Text, RegFont, System.Drawing.Brushes.Black, 78, 980)
        If ConfigDiag.CheckBox2.Checked = True Then
            Dim reso = "Any alteration on this Certificate may null and void. Provided under Barangay " & My.Settings.RESNO & ", this Barangay."
            e.Graphics.DrawString(reso, New Font("Calibri", 8, FontStyle.Italic), System.Drawing.Brushes.Black, New Rectangle(75, 1050, w - 150, 50), sf)
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox19.Enabled = True
        Else
            TextBox19.Enabled = False
        End If
    End Sub




    Private Sub prtID_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles prtID.PrintPage

        Dim csz As New PaperSize("ID Size", 350, 230)
        prtID.PrinterSettings.DefaultPageSettings.PaperSize = csz

        'Head Colors
        Dim ThisColor As New SolidBrush(My.Settings.CIDHEADTEXTC)
        Dim ToThisColor As System.Drawing.Color = New System.Drawing.Pen(ThisColor).Color
        Dim ThisBrush = New System.Drawing.Pen(ThisColor)
        Dim HeadTextColor = New SolidBrush(ThisBrush.Color)

        'Body Text Color
        Dim BThisColor As New SolidBrush(My.Settings.CIDBODYTEXTC)
        Dim BToThisColor As System.Drawing.Color = New System.Drawing.Pen(BThisColor).Color
        Dim BThisBrush = New System.Drawing.Pen(BToThisColor)
        Dim BodyTextColor = New SolidBrush(BThisBrush.Color)

        'Text Formatting
        Dim sf As StringFormat = StringFormat.GenericTypographic
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word

        'Body Background Color
        Dim BcThisColor As New SolidBrush(My.Settings.CIDBACKC)
        Dim BcToThisColor As Color = New Pen(BcThisColor).Color
        Dim BcThisBrush = New Pen(BcToThisColor)
        Dim BodyBackColor = New SolidBrush(BcThisBrush.Color)
        If My.Settings.IDCheckBox = False Then
            Select Case My.Settings.IDCOMBOSELECT
                Case "Basic_1"
                    e.Graphics.DrawImage(My.Resources.Basic_1, 0, 0)
                Case "Basic_2"
                    e.Graphics.DrawImage(My.Resources.Basic_2, 0, 0)
                Case "Basic_3"
                    e.Graphics.DrawImage(My.Resources.Basic_3, 0, 0)
                Case "Classic_1"
                    e.Graphics.DrawImage(My.Resources.Classic_1, 0, 0)
                Case "Classic_2"
                    e.Graphics.DrawImage(My.Resources.Classic_2, 0, 0)
                Case "Classic_3"
                    e.Graphics.DrawImage(My.Resources.Classic_3, 0, 0)
                Case "Modern_1"
                    e.Graphics.DrawImage(My.Resources.Modern_1, 0, 0)
                Case "Modern_2"
                    e.Graphics.DrawImage(My.Resources.Modern_2, 0, 0)
                Case "Modern_3"
                    e.Graphics.DrawImage(My.Resources.Modern_3, 0, 0)
            End Select
        Else
            If Not My.Settings.IDBI = String.Empty Then
                e.Graphics.DrawImage(Image.FromFile(My.Settings.IDBI), 0, 0)

            ElseIf Not My.Settings.CIDBACKC = Color.Empty Then
                e.Graphics.FillRectangle(BodyBackColor, New Rectangle(0, 0, 350, 230))
            Else
                e.Graphics.DrawImage(My.Resources.bg3, 0, 0)
            End If
        End If

        Dim today = DateTime.Today
        Dim day = today.Day
        Dim month = today.Month.ToString("MMMM")
        Dim year = today.Year
        Dim pinF As New PointF(20, 20)
        e.Graphics.DrawRectangle(Pens.Gainsboro, 0, 0, 350, 230)
        'Logo
        If Not My.Settings.BRNGYLOGO = String.Empty Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.BRNGYLOGO), 5, 5, 50, 50)
        End If

        'The QR Code
        Dim qrcode As KeepAutomation.Barcode.Bean.BarCode = New KeepAutomation.Barcode.Bean.BarCode
        qrcode.Symbology = KeepAutomation.Barcode.Symbology.QRCode
        qrcode.QRCodeVersion = KeepAutomation.Barcode.QRCodeVersion.V5
        qrcode.QRCodeECL = KeepAutomation.Barcode.QRCodeECL.H
        qrcode.QRCodeDataMode = KeepAutomation.Barcode.QRCodeDataMode.Auto
        qrcode.CodeToEncode = IDno.Text
        qrcode.BarcodeUnit = KeepAutomation.Barcode.BarcodeUnit.Cm
        qrcode.DPI = 100
        qrcode.X = 2
        qrcode.Y = 2
        qrcode.LeftMargin = 8
        qrcode.RightMargin = 8
        qrcode.TopMargin = 8
        qrcode.BottomMargin = 8
        Dim QRcodePrint As Image = qrcode.generateBarcodeToBitmap

        'ID Information 
        Dim HeaderText = "Republic of the Philippines" & vbCrLf & _
                          Trim(My.Settings.BP102).ToUpper & vbCrLf & _
                          Trim(My.Settings.BCM103).ToUpper
        Dim Brngy = Trim(My.Settings.BAR).ToUpper
        Dim InfoLabel = "Firstname/M.I/Lastname" & vbCrLf & vbCrLf & vbCrLf & _
                            "Birthdate" & vbCrLf & vbCrLf & vbCrLf & _
                            "Status"
        Dim TheSpice = "___________________" & vbCrLf & "Signature"
        e.Graphics.DrawString(TheSpice, New Font("Calibri", 8, FontStyle.Regular), Brushes.Black, 115, 175)

        Dim datetoday = DateTime.Now.ToString("MM/dd/yyyy")
        Dim datelast = DateTime.Now.AddYears(3).ToString("MM/dd/yyyy")
        Dim voterno = ""
        If TextBox42.Text.Length > 1 Then
            voterno = "Voter No.: " & TextBox42.Text.ToUpper
        Else
            voterno = TextBox42.Text.ToUpper
        End If
        e.Graphics.DrawString(datetoday, New Font("Calibri", 6, FontStyle.Bold), Brushes.Black, 320, 130, sf)
        e.Graphics.DrawString(datelast, New Font("Calibri", 6, FontStyle.Bold), Brushes.Black, 320, 140, sf)
        e.Graphics.DrawString(voterno.ToUpper, New Font("Calibri", 8, FontStyle.Bold), Brushes.Black, 5, 200, StringFormat.GenericDefault)

        If Trim(TextBox2.Text).Length > 1 Then
            Dim TheInfo = Trim(TextBox2.Text).ToUpper & " " & Trim(TextBox3.Text.First).ToUpper & ". " & Trim(TextBox1.Text).ToUpper & vbCrLf & vbCrLf & _
                              Trim(RDB1.Text).ToUpper & vbCrLf & vbCrLf & _
                              Trim(ComboBox1.Text).ToUpper
            e.Graphics.DrawString(TheInfo, New Font("Calibri", 9, FontStyle.Bold), BodyTextColor, 115, 90)
        Else
            MsgBox("Add information")
        End If

        e.Graphics.DrawString("RESIDENT ID", New Font("Calibri", 14, FontStyle.Bold), HeadTextColor, 220, 12)
        e.Graphics.DrawString("R-" & IDno.Text, New Font("Calibri", 10, FontStyle.Regular), HeadTextColor, 222, 30)

        e.Graphics.DrawRectangle(Pens.Black, 5, 80, 100, 120)

        If PictureBox2.ImageLocation.Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(PictureBox2.ImageLocation), 6, 81, 98, 120)
        Else
        End If
        e.Graphics.DrawString(HeaderText, New Font("Calibri", 6, FontStyle.Regular), HeadTextColor, 60, 5)
        e.Graphics.DrawString(Brngy, New Font("Calibri", 8, FontStyle.Bold), HeadTextColor, 59, 35)
        e.Graphics.DrawString(InfoLabel, New Font("Calibri", 6, FontStyle.Italic), BodyTextColor, 115, 80)
        e.Graphics.DrawImage(QRcodePrint, 265, 145, 80, 80)
        ' e.Graphics.FillRectangle(Brushes.Black, 5, 80, 100, 120)
        'Execute texts

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim csz As New PaperSize("ID Size", 350, 230)
        Dim pd As New PrintPreviewDialog
        prtID.PrinterSettings = New PrinterSettings
        prtID.PrinterSettings.DefaultPageSettings.PaperSize = csz
        prtID.DocumentName = "Resident ID"
        pd.Document = prtID
        pd.ShowDialog()
        Button32.Show()
        With ListBox1.Items
            Dim TheInfo = Trim(TextBox2.Text).ToUpper & " " & Trim(TextBox3.Text.First).ToUpper & ". " & Trim(TextBox1.Text).ToUpper
            .Add("Print Resident ID: [" & IDno.Text & "]" & TheInfo)
            .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
        End With
        _insertolog("Printed ID Card", IDno.Text)
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintClearanceButton.Click
        Try
            Dim pd As New PrintPreviewDialog
            pd.Document = Reprint
            Reprint.DocumentName = "Reprint Certificate"

            pd.ShowDialog()

        Catch ex As Exception
        Finally
            _insertolog("Reprint Certification", SearchBox.Text)
        End Try
    End Sub

    Private Sub Reprint_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles Reprint.PrintPage
        Dim today = DateTime.Today
        Dim day = today.Day
        Dim month = today.ToString("MMMM")
        Dim year = today.Year
        Dim pinF As New PointF(20, 20)

        'Fonts
        Dim RegFont As Font = New Font("Verdana", 12, FontStyle.Regular)
        Dim BoldFont As Font = New Font("Verdana", 12, FontStyle.Bold)
        Dim Title As Font = New Font("Calibri", 16, FontStyle.Bold)
        Dim DocFont As Font = New Font("Verdana", 20, FontStyle.Bold)

        'Page area
        Dim rect As New Rectangle()

        'Text Alignment
        Dim sf As StringFormat = StringFormat.GenericTypographic
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word

        Dim rt As New StringFormat
        rt.FormatFlags = StringFormatFlags.DirectionVertical

        Dim cx, y, w, h As Integer
        cx = 414.5
        y = 20
        w = 827
        h = 1168
        If Trim(My.Settings.CMLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.CMLOGO), 650, 20, 120, 120)
        Else

        End If
        If Trim(My.Settings.BRNGYLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.BRNGYLOGO), 75, 20, 120, 120)

        Else

        End If

        Dim Purpose = TextBox19.Text
        Select Case CheckBox2.Checked
            Case True
                Purpose = TextBox19.Text.ToUpper
            Case False
                Purpose = "APPLICATION"
        End Select
        Dim HeaderText = "Republic of the Philippines" & vbCrLf & _
                        "Department of the Interior and Local Government" & vbCrLf & _
                        "PROVINCE OF " & Trim(My.Settings.BP102).ToUpper & vbCrLf & _
                        "MUNICIPALITY " & Trim(My.Settings.BCM103).ToUpper

        Dim LeadClear = "THIS IS TO CERTIFY the person whose name, picture, and signature appear hereon has requested a CLEARANCE from this office."
        Dim DeLblInfo = "NAME: " & vbCrLf & _
                        "BIRTHDATE: " & vbCrLf & _
                        "PLACE OF BIRTH: " & vbCrLf & _
                        "GENDER: " & vbCrLf & _
                        "CIVIL STATUS: " & vbCrLf & _
                        "PURPOSE: " & vbCrLf & _
                        "ADDRESS: "


        Dim FurLetter = "THIS IS TO FURTHER CERTIFY that the above-mentioned person is known to me good moral character and is a law-abiding citizen. He/she has no pending derogatory record in this office."
        Dim IssuesText = "Issued this " & day & " day of " & month & ", " & year & " at Barangay Hall, " & Trim(My.Settings.BAR) & ", " & Trim(My.Settings.BCM103) & ", " & Trim(My.Settings.BP102) & ", Philippines."


        Dim Brngy = "BARANGAY " & Trim(My.Settings.BAR).ToUpper
        Dim PresideOffice = "OFFICE OF THE BARANGAY CHAIRMAN"
        Dim ToWhom = "TO WHOM IT MAY CONCERN:"
        Dim sss = "Certified By: " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                    "_____________________" & vbCrLf & _
                     " " & My.Settings.BChair.ToUpper
        Dim Prinfo = "Printed on: " & DateTime.Today.ToShortDateString & " via " & System.Windows.Forms.Application.ProductName & " " & System.Windows.Forms.Application.ProductVersion
        e.Graphics.DrawLine(Pens.Black, 50, 150, 800, 150)

        'Align texts
        e.Graphics.DrawString(HeaderText, RegFont, Brushes.Black, New Rectangle(cx * 0.5, 20, w * 0.52, 100), sf)
        e.Graphics.DrawString(Brngy, Title, Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 80), sf)
        e.Graphics.DrawString(PresideOffice, Title, Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 200), sf)

        'Letters
        e.Graphics.DrawString("BARANGAY CLEARANCE", DocFont, Brushes.Black, New Rectangle(cx * 0.5, 90, w * 0.52, 300), sf)
        e.Graphics.DrawString(ToWhom, BoldFont, Brushes.Black, New Rectangle(75, 300, w - 150, 100))
        e.Graphics.DrawString(LeadClear, RegFont, Brushes.Black, New Rectangle(75, 350, w - 150, 450), StringFormat.GenericDefault)
        'Print information from database

        Try
            Dim conn As New SqlConnection(ConnStr)
            Dim cmd As New SqlCommand("SELECT DID, RLNAME, RFNAME, RMNAME, RPB, RDB, RCS, RSEX, RSADD FROM RFS_PID WHERE ID = '" & SearchBox.Text & "'", conn)
            conn.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader()
            While reader.Read()
                Dim DocNo = reader("DID")
                'Barcode
                Dim barcode As New KeepAutomation.Barcode.Bean.BarCode()
                barcode.Symbology = KeepAutomation.Barcode.Symbology.Code128B
                barcode.CodeToEncode = DocNo
                barcode.ChecksumEnabled = True
                barcode.X = 1
                barcode.Y = 50
                barcode.BarCodeWidth = 100
                barcode.BarCodeHeight = 70
                barcode.Orientation = KeepAutomation.Barcode.Orientation.Degree0
                barcode.BarcodeUnit = KeepAutomation.Barcode.BarcodeUnit.Pixel
                barcode.DPI = 300

                Dim thisbaracode As Image = barcode.generateBarcodeToBitmap
                e.Graphics.DrawImage(thisbaracode, 75, 1020, 170, 45)

                Dim ThisName = reader("RFNAME").ToUpper & " " & reader("RMNAME").ToString(0) & ". " & reader("RLNAME").ToUpper
                Dim theInfo = ThisName & vbCrLf & _
                       reader("RDB") & vbCrLf & _
                       reader("RPB").ToUpper & vbCrLf & _
                       reader("RCS").ToUpper & vbCrLf & _
                       reader("RSEX").ToUpper & vbCrLf & _
                       "APPLICATION" & vbCrLf & reader("RSADD").ToUpper
                e.Graphics.DrawString(DeLblInfo, RegFont, Brushes.Black, 210, 420)
                e.Graphics.DrawString(theInfo, BoldFont, Brushes.Black, New Rectangle(370, 420, w * 0.48, 300), StringFormat.GenericDefault)
                e.Graphics.DrawString("Doc. No.:" & DocNo, RegFont, Brushes.Black, 78, 980)

            End While
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            'The Profile Photo
            Dim connn As New SqlConnection(ConnStr)
            Dim imgcmd As New SqlCommand("SELECT IMAGE FROM LOGGER WHERE ID=" & SearchBox.Text & "", connn)
            connn.Open()
            Dim sdr As SqlDataReader = imgcmd.ExecuteReader()
            While sdr.Read()
                Dim ImgStream As New IO.MemoryStream(CType(sdr("IMAGE"), Byte()))
                e.Graphics.DrawImage(Image.FromStream(ImgStream), 80, 410, 120, 155)
            End While
            connn.Close()
        End Try


        e.Graphics.DrawString(FurLetter, RegFont, Brushes.Black, New Rectangle(75, 600, w - 150, 450), StringFormat.GenericDefault)
        e.Graphics.DrawString(IssuesText, RegFont, Brushes.Black, New Rectangle(75, 675, w - 150, 450), StringFormat.GenericDefault)
        e.Graphics.DrawLine(Pens.Black, 80, 590, 735, 590)
        e.Graphics.DrawLine(Pens.Black, 80, 400, 735, 400)

        'Officials' Certification
        e.Graphics.DrawString(sss, BoldFont, Brushes.Black, 480, 780)
        'Thumbprints
        e.Graphics.DrawRectangle(Pens.Black, 75, 780, 100, 120)
        e.Graphics.DrawRectangle(Pens.Black, 175, 780, 100, 120)
        'Document Identification
        e.Graphics.DrawString("__________________", RegFont, Brushes.Black, 78, 940)
        e.Graphics.DrawString("Specimen Signature", RegFont, Brushes.Black, 78, 960)

        e.Graphics.DrawString("Left" & vbCrLf & " Thumbmark", New Font("Calibri", 10), Brushes.Black, 125, 800, sf)
        e.Graphics.DrawString("Right" & vbCrLf & " Thumbmark", New Font("Calibri", 10), Brushes.Black, 225, 800, sf)
        'Printed on datetime
        e.Graphics.DrawString(Prinfo, New Font("Calibri", 8, FontStyle.Regular), Brushes.Gainsboro, New RectangleF(5, 50, 200, 200), rt)
        e.Graphics.DrawString("NOTE: This clearance is valid only for six (6) months from the date of issue.", RegFont, Brushes.Black, w * 0.5, 1100, sf)
        If ConfigDiag.CheckBox2.Checked = True Then
            Dim reso = "Any alteration on this Certificate may null and void. Provided under Barangay " & My.Settings.RESNO & ", this Barangay."
            e.Graphics.DrawString(reso, New Font("Calibri", 8, FontStyle.Italic), Brushes.Black, New Rectangle(75, 1050, w - 150, 50), sf)
        End If


    End Sub



    Private Sub TransferRadioBtn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TransferRadioBtn.CheckedChanged
        If TransferRadioBtn.Checked = True Then
            BusinessTransferGroup.Enabled = True
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            BusinessTransferGroup.Enabled = False
            GroupBox3.Enabled = True
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            BusinessTransferGroup.Enabled = False
            GroupBox3.Enabled = True
        End If
    End Sub

    Private Sub PrintSettingsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub prtClear_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles prtClear.PrintPage
        Dim today = DateTime.Today
        Dim day = today.Day
        Dim month = today.Month.ToString("MMMM")
        Dim year = today.Year
        Dim pinF As New PointF(20, 20)

        'Fonts
        Dim RegFont As Font = New Font("Calibri", 12, FontStyle.Regular)
        Dim BoldFont As Font = New Font("Calibri", 12, FontStyle.Bold)
        Dim Title As Font = New Font("Calibri", 16, FontStyle.Bold)
        Dim DocFont As Font = New Font("Verdana", 20, FontStyle.Bold)

        'Page area
        'Dim rect As New Rectangle()

        'Text Alignment
        Dim sf As StringFormat = StringFormat.GenericTypographic
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word

        'The Barcode
        Dim barcode As KeepAutomation.Barcode.Bean.BarCode = New KeepAutomation.Barcode.Bean.BarCode
        barcode.Symbology = KeepAutomation.Barcode.Symbology.Codabar
        barcode.DisplayText = True
        barcode.BarcodeUnit = KeepAutomation.Barcode.BarcodeUnit.Cm
        barcode.CodeToEncode = DIDno.Text
        barcode.X = 2
        barcode.Y = 40
        barcode.BarCodeWidth = 100
        barcode.BarCodeHeight = 70
        barcode.DPI = 100

        Dim ThisBarcode As Image = barcode.generateBarcodeToBitmap

        Dim rt As New StringFormat
        rt.FormatFlags = StringFormatFlags.DirectionVertical

        Dim cx, y, w, h As Integer
        cx = 414.5
        y = 20
        w = 827
        h = 1168
        If Trim(My.Settings.CMLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.CMLOGO), 640, 20, 110, 110)
        Else

        End If
        If Trim(My.Settings.BRNGYLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.BRNGYLOGO), 70, 20, 110, 110)

        Else

        End If

        Dim Purpose = TextBox19.Text
        Select Case CheckBox2.Checked
            Case True
                Purpose = TextBox19.Text.ToUpper
            Case False
                Purpose = "APPLICATION"
        End Select
        Dim HeaderText = "Republic of the Philippines" & vbCrLf & _
                        "Department of the Interior and Local Government" & vbCrLf & _
                        Trim(My.Settings.BP102).ToUpper & vbCrLf & _
                        Trim(My.Settings.BCM103).ToUpper

        Dim ThisName = Trim(TextBox2.Text & " " & TextBox3.Text.First() & ". " & TextBox1.Text).ToUpper
        Dim LeadClear = "THIS IS TO CERTIFY the person whose name, picture, and signature appear hereon has requested a CLEARANCE from this office."
        Dim DeLblInfo = "NAME: " & vbCrLf & _
                        "BIRTHDATE: " & vbCrLf & _
                        "PLACE OF BIRTH: " & vbCrLf & _
                        "GENDER: " & vbCrLf & _
                        "CIVIL STATUS: " & vbCrLf & _
                        "PURPOSE: " & vbCrLf & _
                        "ADDRESS: "
        Dim theInfo = ThisName & vbCrLf & _
                        Trim(RDB1.Text).ToUpper & vbCrLf & _
                        Trim(TextBox5.Text).ToUpper & vbCrLf & _
                        Trim(ComboBox1.Text).ToUpper & vbCrLf & _
                        Trim(ComboBox2.Text).ToUpper & vbCrLf & _
                        Purpose & vbCrLf & Trim(TextBox14.Text).ToUpper

        Dim FurLetter = "THIS IS TO FURTHER CERTIFY that the above-mentioned person is known to me good moral character and is a law-abiding citizen. He/she has no pending derogatory record in this office."
        Dim IssuesText = "Issued this " & day & " day of " & month & ", " & year & " at Barangay Hall, " & Trim(My.Settings.BAR) & ", " & Trim(My.Settings.BCM103) & ", " & Trim(My.Settings.BP102) & ", Philippines."


        Dim Brngy = Trim(My.Settings.BAR).ToUpper
        Dim PresideOffice = "OFFICE OF THE BARANGAY CHAIRMAN"
        Dim ToWhom = "TO WHOM IT MAY CONCERN:"
        Dim sss = "Certified By: " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                    "_____________________" & vbCrLf & _
                     " " & My.Settings.BChair.ToUpper
        Dim Prinfo = "Printed on: " & DateTime.Today.ToShortDateString & " via " & System.Windows.Forms.Application.ProductName & " " & System.Windows.Forms.Application.ProductVersion
        e.Graphics.DrawLine(Pens.Black, 50, 150, 750, 150)

        'Align texts
        e.Graphics.DrawString(HeaderText, RegFont, Brushes.Black, New Rectangle(cx * 0.5, 20, w * 0.52, 100), sf)
        e.Graphics.DrawString(Brngy, Title, Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 80), sf)
        e.Graphics.DrawString(PresideOffice, Title, Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 200), sf)

        'Letters
        e.Graphics.DrawString("BARANGAY CLEARANCE", DocFont, Brushes.Black, New Rectangle(cx * 0.5, 90, w * 0.52, 300), sf)
        e.Graphics.DrawString(ToWhom, BoldFont, Brushes.Black, New Rectangle(75, 300, w - 150, 100))
        e.Graphics.DrawString(LeadClear, RegFont, Brushes.Black, New Rectangle(75, 350, w - 150, 450), StringFormat.GenericDefault)
        e.Graphics.DrawString(DeLblInfo, RegFont, Brushes.Black, 210, 420)
        e.Graphics.DrawString(theInfo, BoldFont, Brushes.Black, New Rectangle(340, 420, w * 0.48, 300), StringFormat.GenericDefault)
        'Picture of applicant
        If PictureBox2.ImageLocation.Length > 1 Then
            e.Graphics.DrawImage(PictureBox2.Image, 80, 410, 120, 155)
        Else
            e.Graphics.DrawString("Photo" & vbCrLf & "(Passport Size)", New Font("Calibri", 10), Brushes.Black, 140, 480, sf)
            e.Graphics.DrawRectangle(Pens.Black, 79, 415, 121, 156)
        End If

        e.Graphics.DrawString(FurLetter, RegFont, Brushes.Black, New Rectangle(75, 600, w - 150, 450), StringFormat.GenericDefault)
        e.Graphics.DrawString(IssuesText, RegFont, Brushes.Black, New Rectangle(75, 665, w - 150, 450), StringFormat.GenericDefault)
        e.Graphics.DrawLine(Pens.Black, 80, 590, 735, 590)
        e.Graphics.DrawLine(Pens.Black, 80, 400, 735, 400)

        'Officials' Certification
        e.Graphics.DrawString(sss, BoldFont, Brushes.Black, 480, 780)
        'Thumbprints
        e.Graphics.DrawRectangle(Pens.Black, 75, 780, 100, 120)
        e.Graphics.DrawRectangle(Pens.Black, 175, 780, 100, 120)
        'Document Identification
        e.Graphics.DrawString("__________________", RegFont, Brushes.Black, 78, 940)
        e.Graphics.DrawString("Specimen Signature", RegFont, Brushes.Black, 78, 960)
        e.Graphics.DrawImage(ThisBarcode, 70, 1000, 170, 45)
        e.Graphics.DrawString("Left" & vbCrLf & " Thumbmark", New Font("Calibri", 10), Brushes.Black, 125, 800, sf)
        e.Graphics.DrawString("Right" & vbCrLf & " Thumbmark", New Font("Calibri", 10), Brushes.Black, 225, 800, sf)
        'Printed on datetime
        e.Graphics.DrawString(Prinfo, New Font("Calibri", 8, FontStyle.Regular), Brushes.Gainsboro, New RectangleF(5, 50, 200, 200), rt)
        e.Graphics.DrawString("Doc. No.:" & DIDno.Text, RegFont, Brushes.Black, 78, 980)
        e.Graphics.DrawString("NOTE: This clearance is valid only for six (6) months from the date of issue.", New Font("Calibri", 10, FontStyle.Italic), Brushes.Black, w * 0.5, 1080, sf)
        If ConfigDiag.CheckBox2.Checked = True Then
            Dim reso = "Any alteration on this Certificate may null and void. Provided under Barangay " & My.Settings.RESNO & ", this Barangay."
            e.Graphics.DrawString(reso, New Font("Calibri", 8, FontStyle.Italic), Brushes.Black, New Rectangle(75, 1050, w - 150, 50), sf)
        End If
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        prtClear.PrinterSettings = New PrinterSettings
        Dim w As Integer = PpWidthSize.Text * 100
        Dim h As Integer = PpHeightSize.Text * 100
        Dim PpName As String = PaperSizesBox.Text
        Dim pd As New PrintPreviewDialog
        Dim psz As New PaperSize(PpName, w, h)
        prtClear.PrinterSettings.DefaultPageSettings.PaperSize = psz
        pd.Document = prtClear
        prtClear.DocumentName = "Print Barangay Certificate"
        pd.ShowDialog()
        _insertolog("Printed Clearance", IDno.Text)
        With ListBox1.Items
            Dim TheInfo = Trim(TextBox2.Text).ToUpper & " " & Trim(TextBox3.Text.First).ToUpper & ". " & Trim(TextBox1.Text).ToUpper
            .Add("Print Clearance: [" & IDno.Text & "]" & TheInfo)
            .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
        End With

    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        Try
            __NEWBUS()
        Catch ex As Exception
            With ListBox1.Items
                .Add("Error Registering Business " & BCID.Text)
                .Add("EventAction - " & DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
            End With
        Finally
            _insertolog("New Business " & TextBox18.Text, BCID.Text)
            With ListBox1.Items
                .Add("New Business: " & BCID.Text)
                .Add("EventAction - " & DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
            End With
            __BUSINIT()
        End Try



    End Sub

    Private Sub DehLink_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles DehLink.LinkClicked
        Process.Start("https://dehgel.wordpress.com")
    End Sub


    Private Sub RefreshToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        DataGridView1.Update()
        DataGridView1.Refresh()
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        _DELETE()
    End Sub

    Private Sub ProfileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProfileToolStripMenuItem.Click
        Dim rp = ResidentProfile
        Dim pr = DataGridView1.SelectedRows(0)
        Try
            Using conn As New SqlConnection(ConnStr)
                Using cmd As New SqlCommand("SELECT IMAGE FROM LOGGER WHERE ID ='" & pr.Cells(1).Value & "'", conn)
                    conn.Open()
                    Dim img As Byte() = DirectCast(cmd.ExecuteScalar(), Byte())
                    Dim ms As MemoryStream = New MemoryStream(img)
                    ResidentProfile.PictureBox1.Image = Image.FromStream(ms)
                    ResidentProfile.PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                    conn.Close()
                End Using
            End Using

        Catch ex As Exception
        Finally
            Dim ProfileDialogTitle = "Profile: [" & pr.Cells(1).Value & "] " & pr.Cells(4).Value & " " & pr.Cells(5).Value.ToString(0) & ". " & pr.Cells(3).Value
            ResidentProfile.Text = ProfileDialogTitle
            With pr

                rp.TextBox4.Text = pr.Cells(1).Value
                rp.TextBox1.Text = pr.Cells(4).Value & " " & pr.Cells(5).Value & " " & pr.Cells(3).Value
                rp.DateTimePicker1.Value = pr.Cells(7).Value
                rp.TextBox3.Text = pr.Cells(8).Value.ToString
                rp.ComboBox1.Text = Trim(.Cells(10).Value.ToString)
                rp.CivilStats2.Text = Trim(pr.Cells(9).Value)
                rp.TextBox6.Text = pr.Cells(14).Value.ToString
                rp.TextBox11.Text = pr.Cells(25).Value.ToString
                rp.TextBox12.Text = pr.Cells(26).Value.ToString
                rp.TextBox18.Text = pr.Cells(15).Value.ToString
                rp.TextBox20.Text = pr.Cells(16).Value.ToString
                rp.TextBox2.Text = pr.Cells(20).Value.ToString
                rp.TextBox5.Text = pr.Cells(21).Value.ToString
                rp.TextBox7.Text = pr.Cells(22).Value.ToString
                rp.TextBox14.Text = pr.Cells(28).Value.ToString
                'rp.DateTimePicker3.Value = pr.Cells(17).Value
                rp.TextBox17.Text = pr.Cells(32).Value.ToString
                rp.TextBox15.Text = pr.Cells(36).Value.ToString
                rp.TextBox16.Text = pr.Cells(35).Value.ToString
            End With
        End Try

        ResidentProfile.Show()

    End Sub

    Private Sub EditProfileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim rp = ResidentProfile
        Dim pr = DataGridView1.SelectedRows(0)
        ResidentProfile.Text = "Revise Information for Control No.: " & pr.Cells(0).Value

        Dim ProfileDialogTitle = "Profile: [" & pr.Cells(1).Value & "] " & pr.Cells(4).Value & " " & pr.Cells(5).Value.ToString(0) & ". " & pr.Cells(3).Value

        ResidentProfile.Text = ProfileDialogTitle
        With rp
            .TextBox4.Text = pr.Cells(1).Value
            .TextBox1.Text = pr.Cells(4).Value & " " & pr.Cells(5).Value.ToString(0) & ". " & pr.Cells(3).Value
            .DateTimePicker1.Text = pr.Cells(7).Value
            .TextBox3.Text = pr.Cells(8).Value
            .ComboBox1.Text = pr.Cells(10).Value
            .CivilStats2.Text = pr.Cells(9).Value
            .TextBox6.Text = pr.Cells(14).Value
           
        End With

        rp.Button1.Enabled = False
        ResidentProfile.Show()
        ResidentProfile.Show()

    End Sub



    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        prtBusiness.PrinterSettings = New PrinterSettings
        Dim w As Integer = PpWidthSize.Text * 100
        Dim h As Integer = PpHeightSize.Text * 100
        Dim PpName As String = PaperSizesBox.Text
        Dim pd As New PrintPreviewDialog
        Dim psz As New PaperSize(PpName, w, h)
        prtBusiness.PrinterSettings.DefaultPageSettings.PaperSize = psz
        pd.Document = prtBusiness
        prtBusiness.DocumentName = "Print Business Certificate"
        pd.ShowDialog()
        With ListBox1.Items
            .Add("Print Business Permit: " & BCID.Text)
            .Add("EventAction - " & DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
        End With
        _insertolog("Printed Business Certificate", BCID.Text)
    End Sub

    Private Sub SearchBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SearchBox.TextChanged
        _TEXTSEARCH()
    End Sub

    Private Sub PrintClearanceButton_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PrintClearanceButton.MouseHover
        Dim tt As New System.Windows.Forms.ToolTip
        tt.SetToolTip(PrintClearanceButton, "Print Barangay Clearance")

    End Sub

    Private Sub PrintIdButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintIdButton.Click
        Dim csz As New PaperSize("ID Size", 350, 230)

        ReprintID.PrinterSettings = New PrinterSettings
        ReprintID.PrinterSettings.DefaultPageSettings.PaperSize = csz
        Dim pd As New PrintPreviewDialog
        pd.Name = "Print Resident ID"
        ReprintID.DocumentName = "Resident ID"
        pd.Document = ReprintID
        pd.ShowDialog()
        With ListBox1.Items
            Dim TheInfo = Trim(TextBox2.Text).ToUpper & " " & Trim(TextBox3.Text.First).ToUpper & ". " & Trim(TextBox1.Text).ToUpper
            .Add("Print Resident ID: [" & IDno.Text & "]" & TheInfo)
            .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
        End With
        _insertolog("Reprint ID Card", SearchBox.Text)
    End Sub

    Private Sub PrintIdButton_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PrintIdButton.MouseHover
        Dim tt As New System.Windows.Forms.ToolTip
        tt.SetToolTip(PrintIdButton, "Print Identification Card")
    End Sub

    Private Sub Button18_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Dim pd As New PrintPreviewDialog
        pd.Document = RePrintCert
        RePrintCert.PrinterSettings = New PrinterSettings
        Dim w As Integer = 827
        Dim h As Integer = 1167
        Dim PpName As String = "A4"
        Dim psz As New PaperSize(PpName, w, h)
        RePrintCert.PrinterSettings.DefaultPageSettings.PaperSize = psz
        pd.Document = RePrintCert
        RePrintCert.DocumentName = "Print Certificate"
        pd.ShowDialog()
        With ListBox1.Items
            Dim TheInfo = Trim(TextBox2.Text).ToUpper & " " & Trim(TextBox3.Text.First).ToUpper & ". " & Trim(TextBox1.Text).ToUpper
            .Add("Print Resident Certificate: [" & IDno.Text & "]" & TheInfo)
            .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
        End With
        _insertolog("Reprint Clearance", SearchBox.Text)
    End Sub

    Private Sub Button18_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button18.MouseHover
        Dim tt As New System.Windows.Forms.ToolTip
        tt.SetToolTip(Button18, "Print Residential Certificate")
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        AboutRinfosys.Show()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
     
        Dim pd As New PrintPreviewDialog
        pd.Name = "Print List of Residents"
        TheResList.DocumentName = "List of Residents"
        pd.Document = TheResList
        pd.ShowDialog()

        With ListBox1.Items
            .Add("Printed List of Residents")
            .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
        End With

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SourceEconBtn.Click

        Dim pd As New PrintPreviewDialog
        pd.Name = "Economic Features"
        PrintEconomy.DocumentName = Replace(My.Settings.BAR, " ", "_") & "_Economy"
        pd.Document = PrintEconomy
        pd.ShowDialog()
        With ListBox1.Items
            .Add("Print Economy")
            .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
        End With


    End Sub

    Private Sub TextBox15_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox15.GotFocus

    End Sub


    Private Sub TextBox15_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox15.KeyUp

    End Sub
    Private Sub prtBackID_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles prtBackID.PrintPage
        Dim csz As New PaperSize("ID Size", 350, 230)
        prtID.PrinterSettings.DefaultPageSettings.PaperSize = csz

        'Head Colors
        Dim ThisColor As New SolidBrush(My.Settings.CIDHEADTEXTC)
        Dim ToThisColor As Color = New Pen(ThisColor).Color
        Dim ThisBrush = New Pen(ThisColor)
        Dim HeadTextColor = New SolidBrush(ThisBrush.Color)

        'Body Text Color
        Dim BThisColor As New SolidBrush(My.Settings.CIDBODYTEXTC)
        Dim BToThisColor As Color = New Pen(BThisColor).Color
        Dim BThisBrush = New Pen(BToThisColor)
        Dim BodyTextColor = New SolidBrush(BThisBrush.Color)

        'Text Formatting
        Dim sf As StringFormat = StringFormat.GenericTypographic
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word


        Dim today = DateTime.Today
        Dim day = today.Day
        Dim month = today.Month.ToString("MMMM")
        Dim year = today.Year
        Dim pinF As New PointF(20, 20)
        e.Graphics.DrawRectangle(Pens.Gainsboro, 0, 0, 350, 230)

        'The Barcode Code
        Dim tbarcode As KeepAutomation.Barcode.Bean.BarCode = New KeepAutomation.Barcode.Bean.BarCode
        tbarcode.Symbology = KeepAutomation.Barcode.Symbology.Codabar

        tbarcode.CodeToEncode = DIDno.Text
        tbarcode.BarcodeUnit = KeepAutomation.Barcode.BarcodeUnit.Cm
        tbarcode.DisplayText = False
        tbarcode.BearerBarTop = 2
        tbarcode.BarCodeHeight = 20
        tbarcode.BarCodeWidth = 350
        Dim QRcodePrint As Image = tbarcode.generateBarcodeToBitmap

        e.Graphics.DrawImage(QRcodePrint, 0, 180, 350, 20)
        ' e.Graphics.FillRectangle(Brushes.Black, 5, 80, 100, 120)
        'Execute texts
        'e.Graphics.DrawString(DateTime.Now.ToString("MMMM dd, yyyy"), New Font("Verdana", 10, FontStyle.Bold), Brushes.Black, New Rectangle(5, 10, 340, 20), sf)
        e.Graphics.DrawString("Issued on " & DateTime.Now.ToString("MMMM dd, yyyy") & " as a courtesy by this office for identification purposes. The cardholder is a bonafide resident of this Barangay. Valid only when signed by Barangay Chairman.", New Font("Calibri", 10, FontStyle.Regular), Brushes.Black, New Rectangle(5, 20, 340, 100), sf)
        'e.Graphics.DrawString("The cardholder is a bonafide resident of this Barangay." & vbCrLf & "Valid only when signed by Barangay Official.", New Font("Calibri", 10, FontStyle.Regular), Brushes.Black, 350 * 0.5, 100, sf)
        e.Graphics.DrawString(Trim(My.Settings.BChair).ToUpper, New Font("Calibri", 10, FontStyle.Bold), Brushes.Black, 350 * 0.5, 160, sf)

    End Sub

    Private Sub Button32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button32.Click
        Dim csz As New PaperSize("ID Size", 350, 230)
        Dim pd As New PrintPreviewDialog
        prtBackID.PrinterSettings = New PrinterSettings
        prtBackID.PrinterSettings.DefaultPageSettings.PaperSize = csz
        prtBackID.DocumentName = "Resident ID"
        pd.Document = prtBackID
        pd.ShowDialog()

        With ListBox1.Items
            .Add("Print ID Back: " & IDno.Text)
            .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
        End With
    End Sub

    Private Sub prtBusiness_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles prtBusiness.PrintPage
        Dim today = DateTime.Today
        Dim day = today.Day
        Dim month = today.ToString("MMMM")
        Dim year = today.Year
        Dim pinF As New PointF(20, 20)

        'Fonts
        Dim RegFont As Font = New Font("Calibri", 12, FontStyle.Regular)
        Dim BoldFont As Font = New Font("Calibri", 12, FontStyle.Bold)
        Dim Title As Font = New Font("Calibri", 16, FontStyle.Bold)
        Dim DocFont As Font = New Font("Verdana", 20, FontStyle.Bold)

        'Page area
        'Dim rect As New Rectangle()

        'Text Alignment
        Dim sf As StringFormat = StringFormat.GenericTypographic
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word

        'The Barcode
        Dim barcode As KeepAutomation.Barcode.Bean.BarCode = New KeepAutomation.Barcode.Bean.BarCode
        barcode.Symbology = KeepAutomation.Barcode.Symbology.Codabar
        barcode.DisplayText = True
        barcode.BarcodeUnit = KeepAutomation.Barcode.BarcodeUnit.Cm
        barcode.CodeToEncode = BCID.Text
        barcode.X = 2
        barcode.Y = 40
        barcode.BarCodeWidth = 100
        barcode.BarCodeHeight = 70
        barcode.DPI = 100

        Dim ThisBarcode As Image = barcode.generateBarcodeToBitmap

        Dim rt As New StringFormat
        rt.FormatFlags = StringFormatFlags.DirectionVertical

        Dim cx, y, w, h As Integer
        cx = 414.5
        y = 20
        w = 827
        h = 1168
        If Trim(My.Settings.CMLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.CMLOGO), 640, 20, 110, 110)
        Else

        End If
        If Trim(My.Settings.BRNGYLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.BRNGYLOGO), 70, 20, 110, 110)

        Else

        End If

        Dim Purpose = TextBox19.Text
        Select Case CheckBox2.Checked
            Case True
                Purpose = TextBox19.Text.ToUpper
            Case False
                Purpose = "APPLICATION"
        End Select
        Dim HeaderText = "Republic of the Philippines" & vbCrLf & _
                        "Department of the Interior and Local Government" & vbCrLf & _
                        Trim(My.Settings.BP102).ToUpper & vbCrLf & _
                        Trim(My.Settings.BCM103).ToUpper

        Dim LeadClear = "Pursuant to the provisions of the 1991 Local Government Code, Permission is HEREBY GRANTED to the following applicant:"
        Dim DeLblInfo = "NAME OF OWNER: " & vbCrLf & _
                        "NAME OF BUSINESS: " & vbCrLf & _
                        "INDUSTRY OF BUSINESS: " & vbCrLf & _
                        "NATURE OF BUSINESS: " & vbCrLf & _
                        "TIN NO.: " & vbCrLf & _
                        "PLACE OF BUSINESS: "
        Dim theInfo = Trim(BusNameOfApplicant.Text).ToUpper & vbCrLf & _
                        Trim(TextBox18.Text).ToUpper & vbCrLf & _
                        Trim(TextBox22.Text).ToUpper & vbCrLf & _
                        Trim(TextBox21.Text).ToUpper & vbCrLf & _
                        Trim(TextBox23.Text).ToUpper & vbCrLf & _
                        Trim(TextBox20.Text)

        Dim FurLetter = "Applicant is further hereby advised to follow existing strict ordinance in relation with the conduct of his/her business. Violation of the same is a ground for the revocation of this certificate."
        Dim IssuesText = "Issued this " & day & " day of " & month & ", " & year & " at Barangay Hall, " & Trim(My.Settings.BAR) & ", " & Trim(My.Settings.BCM103) & ", " & Trim(My.Settings.BP102) & ", Philippines."


        Dim Brngy = Trim(My.Settings.BAR).ToUpper
        Dim PresideOffice = "OFFICE OF THE BARANGAY CHAIRMAN"
        Dim ToWhom = "TO WHOM IT MAY CONCERN:"
        Dim sss = "Certified By: " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                    "_____________________" & vbCrLf & _
                     " " & My.Settings.BChair.ToUpper
        Dim Prinfo = "Printed on: " & DateTime.Today.ToShortDateString & " via " & System.Windows.Forms.Application.ProductName & " " & System.Windows.Forms.Application.ProductVersion
        e.Graphics.DrawLine(Pens.Black, 50, 150, 750, 150)

        'Align texts
        e.Graphics.DrawString(HeaderText, RegFont, Brushes.Black, New Rectangle(cx * 0.5, 20, w * 0.52, 100), sf)
        e.Graphics.DrawString(Brngy, Title, Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 80), sf)
        e.Graphics.DrawString(PresideOffice, Title, Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 200), sf)

        'Letters
        e.Graphics.DrawString("BUSINESS CERTIFICATE", DocFont, Brushes.Black, New Rectangle(cx * 0.5, 90, w * 0.52, 300), sf)
        e.Graphics.DrawString(ToWhom, BoldFont, Brushes.Black, New Rectangle(75, 300, w - 150, 100))
        e.Graphics.DrawString(LeadClear, RegFont, Brushes.Black, New Rectangle(75, 350, w - 150, 450), StringFormat.GenericDefault)
        e.Graphics.DrawString(DeLblInfo, RegFont, Brushes.Black, 75, 420)
        e.Graphics.DrawString(theInfo, BoldFont, Brushes.Black, New Rectangle(250, 420, 485, 300), StringFormat.GenericDefault)

        e.Graphics.DrawString(FurLetter, RegFont, Brushes.Black, New Rectangle(75, 600, w - 150, 450), StringFormat.GenericDefault)
        e.Graphics.DrawString(IssuesText, RegFont, Brushes.Black, New Rectangle(75, 675, w - 150, 450), StringFormat.GenericDefault)
        e.Graphics.DrawLine(Pens.Black, 80, 590, 735, 590)
        e.Graphics.DrawLine(Pens.Black, 80, 400, 735, 400)

        'Officials' Certification
        e.Graphics.DrawString(sss, BoldFont, Brushes.Black, 480, 780)
        'Thumbprints
        'e.Graphics.DrawRectangle(Pens.Black, 75, 780, 100, 120)
        'e.Graphics.DrawRectangle(Pens.Black, 175, 780, 100, 120)
        'Document Identification
        e.Graphics.DrawString("__________________", RegFont, Brushes.Black, 78, 940)
        e.Graphics.DrawString("Specimen Signature", RegFont, Brushes.Black, 78, 960)
        e.Graphics.DrawImage(ThisBarcode, 70, 1000, 170, 45)
        e.Graphics.DrawString("Not valid " & vbCrLf & "without Official Seal", New Font("Calibri", 10, FontStyle.Italic), Brushes.Black, 160, 820, sf)

        'Printed on datetime
        e.Graphics.DrawString(Prinfo, New Font("Calibri", 8, FontStyle.Regular), Brushes.Gainsboro, New RectangleF(5, 50, 200, 200), rt)
        e.Graphics.DrawString("Doc. No.:" & TextBox29.Text, RegFont, Brushes.Black, 78, 980)
        e.Graphics.DrawString("NOTE: This clearance is valid only for six (6) months from the date of issue.", New Font("Calibri", 10, FontStyle.Italic), Brushes.Black, w * 0.5, 1080, sf)
        If ConfigDiag.CheckBox2.Checked = True Then
            Dim reso = "Any alteration on this Certificate may null and void. Provided under Barangay " & My.Settings.RESNO & ", this Barangay."
            e.Graphics.DrawString(reso, New Font("Calibri", 8, FontStyle.Italic), Brushes.Black, New Rectangle(75, 1050, w - 150, 50), sf)
        End If
    End Sub

    Private Sub RegisterABusinessToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegisterABusinessToolStripMenuItem.Click
        TabControl1.SelectTab(TabPage4)

        Dim rp = Me
        Dim pr = DataGridView1.SelectedRows(0)
        Dim OwnersName = pr.Cells(4).Value & " " & pr.Cells(5).Value.ToString(0) & ". " & pr.Cells(3).Value

        With rp
            .BCID.Text = pr.Cells(1).Value
            .TextBox29.Text = pr.Cells(2).Value
            .BusNameOfApplicant.Text = OwnersName
        End With


    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Try
            __INSERTTOLIST()
            _insertolog(ComboBox8.Text & "listed ", ComboBox8.Text)
            ComboBox8.Text = ComboBox8.SelectedItem.ToString
        Catch ex As Exception
        Finally
            __LISTING()
        End Try



    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim i As Integer = 3
        TimeSpan.FromSeconds(i)
        If TimeSpan.FromSeconds(i).TotalMilliseconds Then
            Timer1.Stop()
            StatusTextBox.Hide()
        End If

    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
        Process.Start("https://rinfosys.wordpress.com")
    End Sub

    Private Sub QrCodeBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles QrCodeBox.Click
        Dim cm As New ContextMenuStrip
        Dim sfd As New System.Windows.Forms.SaveFileDialog

        QrCodeBox.ContextMenuStrip = cm
        cm.Items.Add("Save Image")
        'bsender = sender
        ' Dim ptLowerLeft = New Point(0, 0)
        ' ptLowerLeft = bsender.PointToScreen(ptLowerLeft)
        'cm.Show(ptLowerLeft)
        With cm
            AddHandler .Items(0).Click, AddressOf sfd.ShowDialog
            sfd.Title = "Save Barcode"
            sfd.Filter = "PNG (*.png)|*.png|JPG (*.jpg)|*.jpg|BMP (*.bmp)|*.bmp"
            sfd.FileName = "CN-" & IDno.Text
            sfd.RestoreDirectory = True
            If sfd.ShowDialog = System.Windows.Forms.DialogResult.OK Then

                SaveAsImage(sfd.FileName, QrCodeBox.Image, QrCodeBox.Width, QrCodeBox.Height)
            End If
        End With

    End Sub


    Private Sub Chart1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Chart1.Click

    End Sub

    Private Sub Chart1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Chart1.MouseClick
        Dim cm As New ContextMenuStrip
        Dim sfd As New System.Windows.Forms.SaveFileDialog
        Chart1.ContextMenuStrip = cm
        cm.Items.Add("Save Image")

        With cm
            AddHandler .Items(0).Click, AddressOf sfd.ShowDialog
            With sfd
                .Title = "Save Chart"
                .Filter = "PNG (*.png)|*.png"
                .FileName = "Chart"
                .RestoreDirectory = True
                If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
                    Chart1.SaveImage(sfd.FileName, DataVisualization.Charting.ChartImageFormat.Png)
                End If
            End With
        End With

    End Sub

    Private Sub DataGridView3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView3.Click

    End Sub

    Private Sub DataGridView3_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView3.MouseDown

    End Sub

    Private Sub DeleteAllDataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteAllDataToolStripMenuItem.Click

        Select Case True
            Case TabControl1.SelectedTab Is TabPage3
                Try
                    _TRUNCATE("RFS_PID")
                    _TRUNCATE("RFS_CHILD")
                    _TRUNCATE("RFS_HHM")
                    _TRUNCATE("RFS_DECEASED")

                    _init_db()
                Catch ex As Exception

                Finally

                    Button12.Enabled = False
                    Button15.Enabled = False
                    Button16.Enabled = False
                End Try
                


            Case TabControl1.SelectedTab Is TabPage7
                _TRUNCATE("RFS_EVT")
                __LISTING()
                ProfileToolStripMenuItem.Enabled = False
                RegisterABusinessToolStripMenuItem.Enabled = False
                PersonalInformationToolStripMenuItem.Enabled = False
                ListOfResidentsToolStripMenuItem.Enabled = False
                ListOfBusinessesToolStripMenuItem.Enabled = False
            Case TabControl2.SelectedTab Is TabPage9
                _TRUNCATE("RFS_BUS")
                __BUSINIT()

            Case Else

        End Select
    End Sub



    Private Sub Button33_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)


    End Sub


    Private Sub ListOfResidentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListOfResidentsToolStripMenuItem.Click
        Dim pd As New PrintPreviewDialog
        pd.Document = TheResList
        TheResList.PrinterSettings = New PrinterSettings
        Dim w As Integer = 827
        Dim h As Integer = 1167
        Dim PpName As String = "A4"
        Dim psz As New PaperSize(PpName, w, h)
        TheResList.PrinterSettings.DefaultPageSettings.PaperSize = psz
        pd.Document = TheResList
        TheResList.DocumentName = "Print List of Residents"
        pd.ShowDialog()
        With ListBox1.Items
            Dim TheInfo = Trim(TextBox2.Text).ToUpper & " " & Trim(TextBox3.Text.First).ToUpper & ". " & Trim(TextBox1.Text).ToUpper
            .Add("Print Resident Certificate: [" & IDno.Text & "]" & TheInfo)
            .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
        End With
    End Sub

    Private Sub RePrintCert_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles RePrintCert.PrintPage
        Dim today = DateTime.Today
        Dim day = today.Day
        Dim month = today.ToString("MMMM")
        Dim year = today.Year
        Dim pinF As New PointF(20, 20)

        'Fonts
        Dim RegFont As Font = New Font("Calibri", 12, FontStyle.Regular)
        Dim BoldFont As Font = New Font("Verdana", 12, FontStyle.Bold)
        Dim Title As Font = New Font("Calibri", 16, FontStyle.Bold)
        Dim DocFont As Font = New Font("Verdana", 20, FontStyle.Bold)

        'Page area
        'Dim rect As New Rectangle()

        'Text Alignment
        Dim sf As StringFormat = StringFormat.GenericTypographic
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word

        Dim rt As New StringFormat
        rt.FormatFlags = StringFormatFlags.DirectionVertical

        Dim cx, y, w, h As Integer
        cx = 414.5
        y = 20
        w = 827
        h = 1168
        If Trim(My.Settings.CMLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.CMLOGO), 650, 20, 120, 120)
        Else

        End If
        If Trim(My.Settings.BRNGYLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.BRNGYLOGO), 75, 20, 120, 120)

        Else

        End If

        Dim Purpose = "APPLICATION"
        Dim conn As New SqlConnection(ConnStr)
        Dim cmd As New SqlCommand("SELECT DID,RLNAME,RFNAME,RMNAME FROM RFS_PID WHERE ID=" & SearchBox.Text & "", conn)
        conn.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader()
        While reader.Read()
            'Barcode
            Dim barcode As KeepAutomation.Barcode.Bean.BarCode = New KeepAutomation.Barcode.Bean.BarCode
            barcode.Symbology = KeepAutomation.Barcode.Symbology.Codabar
            barcode.DisplayText = True
            barcode.BarcodeUnit = KeepAutomation.Barcode.BarcodeUnit.Cm
            barcode.CodeToEncode = reader("DID")
            barcode.X = 2
            barcode.Y = 40
            barcode.BarCodeWidth = 100
            barcode.BarCodeHeight = 70
            barcode.DPI = 100

            Dim ThisBarcode As Image = barcode.generateBarcodeToBitmap
            Dim ThisName = Trim(reader("RFNAME") & " " & reader("RMNAME").ToString(0) & ". " & reader("RLNAME")).ToUpper
            Dim LetterBody = "          THIS IS TO CERTIFY that per records available in this office, " & ThisName & " is a bonafide resident of " & Trim(My.Settings.BAR).ToUpper & ", " & Trim(My.Settings.BCM103).ToUpper & ", " & Trim(My.Settings.BP102).ToUpper & ", Philippines. " & vbCrLf & vbCrLf & _
                          "         It is further certifies that above - named person has no pending and/or derogatory record filed against him/her as of this date." & vbCrLf & vbCrLf & _
                          "         This certification is issued upon the request of the above mentioned person to support him/her for " & Trim(Purpose).ToUpper & " and whatsoever legal purposes that may serve him/her best." & vbCrLf & vbCrLf & _
                          "         Issued this " & day & " day of " & month & ", " & year & " at Barangay Hall, " & Trim(My.Settings.BAR).ToUpper & ", " & Trim(My.Settings.BCM103).ToUpper & ", " & Trim(My.Settings.BP102).ToUpper & ", Philippines."

            e.Graphics.DrawString(LetterBody, RegFont, Brushes.Black, New Rectangle(75, y * 18, w - 150, 450), StringFormat.GenericDefault)

            'Document Identification
            e.Graphics.DrawImage(ThisBarcode, 70, 1000, 170, 45)
            'Printed on datetime
            e.Graphics.DrawString("Doc. No.:" & reader("DID"), RegFont, Brushes.Black, 78, 980)
        End While
        conn.Close()
        Dim HeaderText = "Republic of the Philippines" & vbCrLf & _
                             "Department of the Interior and Local Government" & vbCrLf & _
                             Trim(My.Settings.BP102).ToUpper & vbCrLf & _
                             Trim(My.Settings.BCM103).ToUpper

        Dim Brngy = Trim(My.Settings.BAR).ToUpper
        Dim PresideOffice = "OFFICE OF THE BARANGAY CHAIRMAN"
        Dim ToWhom = "TO WHOM IT MAY CONCERN:"
        Dim sss = "Certified By: " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                    "_____________________" & vbCrLf & _
                     " " & My.Settings.BChair.ToUpper
        'Align texts
        e.Graphics.DrawString(HeaderText, RegFont, Brushes.Black, New Rectangle(cx * 0.5, 20, w * 0.52, 100), sf)
        e.Graphics.DrawString(Brngy, Title, Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 80), sf)
        e.Graphics.DrawString(PresideOffice, Title, Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 200), sf)
        e.Graphics.DrawString(sss, BoldFont, Brushes.Black, 480, 700)
        e.Graphics.DrawString("CERTIFICATION", DocFont, Brushes.Black, New Rectangle(cx * 0.5, 100, w * 0.52, 300), sf)
        e.Graphics.DrawString("PERMANENT RESIDENT", BoldFont, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 120, w * 0.52, 300), sf)

        e.Graphics.DrawString(ToWhom, BoldFont, Brushes.Black, New Rectangle(75, y * 16, w - 150, 100))
        e.Graphics.DrawString("__________________", RegFont, Brushes.Black, 78, 920)
        e.Graphics.DrawString("Specimen Signature", RegFont, Brushes.Black, 78, 940)
        Dim Prinfo = "Printed on: " & DateTime.Today.ToShortDateString & " via " & System.Windows.Forms.Application.ProductName & " " & System.Windows.Forms.Application.ProductVersion
        e.Graphics.DrawLine(Pens.Black, 50, 150, 750, 150)
        e.Graphics.DrawString(Prinfo, New Font("Calibri", 8, FontStyle.Regular), Brushes.Gainsboro, New RectangleF(5, 50, 200, 200), rt)

        If ConfigDiag.CheckBox2.Checked = True Then
            Dim reso = "Any alteration on this Certificate may null and void. Provided under Barangay " & My.Settings.RESNO & ", this Barangay."
            e.Graphics.DrawString(reso, New Font("Calibri", 8, FontStyle.Italic), Brushes.Black, New Rectangle(75, 1050, w - 150, 50), sf)
        End If
    End Sub

    Private Sub TextBox42_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TableLayoutPanel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub
    Private RegVal1 As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser
    Private RegKey2 As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.ClassesRoot
    Private RegKey1 As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine

    Private Sub ReprintID_EndPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles ReprintID.EndPrint
        prtBackID.Print()

    End Sub

    Private Sub ReprintID_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles ReprintID.PrintPage

        Dim csz As New PaperSize("ID Size", 350, 230)
        prtID.PrinterSettings.DefaultPageSettings.PaperSize = csz

        'Head Colors
        Dim ThisColor As New SolidBrush(My.Settings.CIDHEADTEXTC)
        Dim ToThisColor As Color = New Pen(ThisColor).Color
        Dim ThisBrush = New Pen(ThisColor)
        Dim HeadTextColor = New SolidBrush(ThisBrush.Color)

        'Body Text Color
        Dim BThisColor As New SolidBrush(My.Settings.CIDBODYTEXTC)
        Dim BToThisColor As Color = New Pen(BThisColor).Color
        Dim BThisBrush = New Pen(BToThisColor)
        Dim BodyTextColor = New SolidBrush(BThisBrush.Color)

        'Text Formatting
        Dim sf As StringFormat = StringFormat.GenericTypographic
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word

        'Body Background Color
        Dim BcThisColor As New SolidBrush(My.Settings.CIDBACKC)
        Dim BcToThisColor As Color = New Pen(BcThisColor).Color
        Dim BcThisBrush = New Pen(BcToThisColor)
        Dim BodyBackColor = New SolidBrush(BcThisBrush.Color)
        If My.Settings.IDCheckBox = False Then
            Select Case My.Settings.IDCOMBOSELECT
                Case "Basic_1"
                    e.Graphics.DrawImage(My.Resources.Basic_1, 0, 0)
                Case "Basic_2"
                    e.Graphics.DrawImage(My.Resources.Basic_2, 0, 0)
                Case "Basic_3"
                    e.Graphics.DrawImage(My.Resources.Basic_3, 0, 0)
                Case "Classic_1"
                    e.Graphics.DrawImage(My.Resources.Classic_1, 0, 0)
                Case "Classic_2"
                    e.Graphics.DrawImage(My.Resources.Classic_2, 0, 0)
                Case "Classic_3"
                    e.Graphics.DrawImage(My.Resources.Classic_3, 0, 0)
                Case "Modern_1"
                    e.Graphics.DrawImage(My.Resources.Modern_1, 0, 0)
                Case "Modern_2"
                    e.Graphics.DrawImage(My.Resources.Modern_2, 0, 0)
                Case "Modern_3"
                    e.Graphics.DrawImage(My.Resources.Modern_3, 0, 0)
            End Select
        Else
            If Not My.Settings.IDBI = String.Empty Then
                e.Graphics.DrawImage(Image.FromFile(My.Settings.IDBI), 0, 0)

            ElseIf Not My.Settings.CIDBACKC = Color.Empty Then
                e.Graphics.FillRectangle(BodyBackColor, New Rectangle(0, 0, 350, 230))
            Else
                e.Graphics.DrawImage(My.Resources.bg3, 0, 0)
            End If
        End If

        Dim today = DateTime.Today
        Dim day = today.Day
        Dim month = today.Month.ToString("MMMM")
        Dim year = today.Year
        Dim pinF As New PointF(20, 20)
        e.Graphics.DrawRectangle(Pens.Gainsboro, 0, 0, 350, 230)
        'Logo
        If Not My.Settings.BRNGYLOGO = String.Empty Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.BRNGYLOGO), 5, 5, 50, 50)
        End If
        Try
            Dim conn As New SqlConnection(ConnStr)
            Dim cmd As New SqlCommand("SELECT ID, RLNAME, RFNAME, RMNAME, RDB, RCS, RSEX, VIDNO FROM RFS_PID WHERE ID=" & SearchBox.Text & "", conn)
            conn.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader()
            While reader.Read()
                'The QR Code
                Dim qrcode As KeepAutomation.Barcode.Bean.BarCode = New KeepAutomation.Barcode.Bean.BarCode
                qrcode.Symbology = KeepAutomation.Barcode.Symbology.QRCode
                qrcode.QRCodeVersion = KeepAutomation.Barcode.QRCodeVersion.V5
                qrcode.QRCodeECL = KeepAutomation.Barcode.QRCodeECL.H
                qrcode.QRCodeDataMode = KeepAutomation.Barcode.QRCodeDataMode.Auto
                qrcode.CodeToEncode = reader("ID")
                qrcode.BarcodeUnit = KeepAutomation.Barcode.BarcodeUnit.Cm
                qrcode.DPI = 100
                qrcode.X = 2
                qrcode.Y = 2
                qrcode.LeftMargin = 8
                qrcode.RightMargin = 8
                qrcode.TopMargin = 8
                qrcode.BottomMargin = 8
                Dim QRcodePrint As Image = qrcode.generateBarcodeToBitmap
                e.Graphics.DrawImage(QRcodePrint, 265, 145, 80, 80)

                Dim TheInfo = Trim(reader("RFNAME")).ToUpper & " " & Trim(reader("RMNAME").ToString(0)).ToUpper & ". " & Trim(reader("RLNAME")).ToUpper & vbCrLf & vbCrLf & _
                                  Trim(reader("RDB")).ToUpper & vbCrLf & vbCrLf & _
                                  Trim(reader("RSEX")).ToUpper
                e.Graphics.DrawString(TheInfo, New Font("Calibri", 9, FontStyle.Bold), BodyTextColor, 115, 90)
                e.Graphics.DrawString("R-" & reader("ID"), New Font("Calibri", 10, FontStyle.Regular), HeadTextColor, 222, 30)
                e.Graphics.DrawString(reader("VIDNO"), New Font("Calibri", 10, FontStyle.Bold), Brushes.Black, 5, 208, StringFormat.GenericDefault)
            End While
            conn.Close()
        Catch ex As Exception

        Finally
            Dim connn As New SqlConnection(ConnStr)
            Dim imgcmd As New SqlCommand("SELECT IMAGE FROM LOGGER WHERE ID=" & SearchBox.Text & "", connn)
            connn.Open()
            Dim sdr As SqlDataReader = imgcmd.ExecuteReader()
            While sdr.Read()
                Dim ImgStream As New IO.MemoryStream(CType(sdr("IMAGE"), Byte()))
                e.Graphics.DrawImage(Image.FromStream(ImgStream), 6, 81, 98, 120)
            End While
            connn.Close()
        End Try

        'ID Information 
        Dim HeaderText = "Republic of the Philippines" & vbCrLf & _
                          Trim(My.Settings.BP102).ToUpper & vbCrLf & _
                          Trim(My.Settings.BCM103).ToUpper
        Dim Brngy = Trim(My.Settings.BAR).ToUpper
        Dim InfoLabel = "Firstname/M.I/Lastname" & vbCrLf & vbCrLf & vbCrLf & _
                            "Birthdate" & vbCrLf & vbCrLf & vbCrLf & _
                            "Status"

        Dim TheSpice = "___________________" & vbCrLf & "Signature"
        e.Graphics.DrawString(TheSpice, New Font("Calibri", 8, FontStyle.Regular), Brushes.Black, 115, 175)

        Dim datetoday = DateTime.Now.ToString("MM/dd/yyyy")
        Dim datelast = DateTime.Now.AddYears(3).ToString("MM/dd/yyyy")

        e.Graphics.DrawString(datetoday, New Font("Calibri", 6, FontStyle.Bold), Brushes.Black, 320, 130, sf)
        e.Graphics.DrawString(datelast, New Font("Calibri", 6, FontStyle.Bold), Brushes.Black, 320, 140, sf)


        e.Graphics.DrawString("RESIDENT ID", New Font("Calibri", 14, FontStyle.Bold), HeadTextColor, 220, 12)

        e.Graphics.DrawRectangle(Pens.Black, 5, 80, 100, 120)

        e.Graphics.DrawString(HeaderText, New Font("Calibri", 6, FontStyle.Regular), HeadTextColor, 60, 5)
        e.Graphics.DrawString(Brngy, New Font("Calibri", 8, FontStyle.Bold), HeadTextColor, 59, 35)
        e.Graphics.DrawString(InfoLabel, New Font("Calibri", 6, FontStyle.Italic), BodyTextColor, 115, 80)

        ' e.Graphics.FillRectangle(Brushes.Black, 5, 80, 100, 120)
        'Execute texts
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
       
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        RePrintCert.Print()
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        CaptureCam.Show()

    End Sub

    Private Sub VoterCheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VoterCheckBox.CheckedChanged
        If VoterCheckBox.Checked = True Then
            TextBox43.Enabled = False
            TextBox43.Text = My.Settings.BAR & ", " & My.Settings.BCM103
        Else
            TextBox43.Enabled = True
        End If
    End Sub
   
    Private Sub TheResList_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles TheResList.PrintPage

        Dim zzzzzzzzzzzzz As New PaperSize("A4", 827, 1168)
        TheResList.PrinterSettings = New PrinterSettings
        TheResList.PrinterSettings.DefaultPageSettings.PaperSize = zzzzzzzzzzzzz

        Dim today = DateTime.Today
        Dim day = today.Day
        Dim month = today.Month.ToString("MMMM")
        Dim year = today.Year
        Dim pinF As New PointF(20, 20)

        'Fonts
        Dim RegFont As Font = New Font("Calibri", 12, FontStyle.Regular)
        Dim BoldFont As Font = New Font("Calibri", 12, FontStyle.Bold)
        Dim Title As Font = New Font("Calibri", 16, FontStyle.Bold)
        Dim DocFont As Font = New Font("Calibri", 20, FontStyle.Bold)

        'Text Alignment
        Dim sf As StringFormat = StringFormat.GenericTypographic
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word


        Dim rt As New StringFormat
        rt.FormatFlags = StringFormatFlags.DirectionVertical

        Dim cx, y, w, h, yy As Integer
        cx = 414.5
        y = 20
        yy = 20
        w = 827
        h = 1168

        If Trim(My.Settings.CMLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.CMLOGO), 650, 20, 120, 120)
        Else

        End If
        If Trim(My.Settings.BRNGYLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.BRNGYLOGO), 75, 20, 120, 120)
        Else
        End If

        Dim HeaderText = "Republic of the Philippines" & vbCrLf & _
                        "Department of the Interior and Local Government" & vbCrLf & _
                        Trim(My.Settings.BP102).ToUpper & vbCrLf & _
                        Trim(My.Settings.BCM103).ToUpper
        Dim Brngy = Trim(My.Settings.BAR).ToUpper

        Dim PresideOffice = "OFFICE OF THE BARANGAY CHAIRMAN"

        Dim sss = "Certified By: " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                    "_____________________" & vbCrLf & _
                     " " & My.Settings.BChair.ToUpper
        Dim Prinfo = "Printed on: " & DateTime.Today.ToShortDateString & " via " & System.Windows.Forms.Application.ProductName & " " & System.Windows.Forms.Application.ProductVersion
        e.Graphics.DrawLine(Pens.Black, 50, 150, 800, 150)
        e.Graphics.DrawString("LIST OF RESIDENTS", DocFont, Brushes.Black, New Rectangle(cx * 0.5, 90, w * 0.52, 280), sf)
        'Align texts
        e.Graphics.DrawString(HeaderText, RegFont, Brushes.Black, New Rectangle(cx * 0.5, 20, w * 0.52, 100), sf)
        e.Graphics.DrawString(Brngy, Title, Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 80), sf)
        e.Graphics.DrawString(PresideOffice, Title, Brushes.Black, New Rectangle(cx * 0.5, 70, w * 0.52, 200), sf)

        'Header Title
        Static PageNo As Single = 0
        Static i As Single = 0
        Dim conn As New SqlConnection(ConnStr)
        Dim dap As New SqlDataAdapter("SELECT RFNAME, RMNAME, RLNAME, RDB, RSEX, RCS FROM RFS_PID", conn)
        Dim ds As New DataSet
        Dim rect As Rectangle = New Rectangle(50, 50, 750, 600)
        dap.Fill(ds)
        i = i + 1
        For i = 1 To ds.Tables(0).Rows.Count - 1


            Dim n = Trim(ds.Tables(0).Rows(i).Item("RFNAME") & " " & ds.Tables(0).Rows(i).Item("RMNAME").ToString(0) & ". " & ds.Tables(0).Rows(i).Item("RLNAME") & ", " & Replace(ds.Tables(0).Rows(i).Item("RDB"), " ", "") & ", " & Replace(ds.Tables(0).Rows(i).Item("RSEX"), " ", "") & ", " & ds.Tables(0).Rows(i).Item("RCS")).ToUpper
            e.Graphics.DrawString(i & ". " & n, RegFont, Brushes.Black, New Rectangle(e.MarginBounds.X, y + 270, 750, y), StringFormat.GenericDefault)
            y += 20
            If y >= 800 Then
                e.HasMorePages = True
                PageNo = PageNo + 1
                e.Graphics.DrawString(PageNo & "/1", RegFont, Brushes.Black, 760, 1100)
                Exit Sub
            Else

            End If
            e.HasMorePages = False
            PageNo = 0
        Next

        'Printed on datetime
        e.Graphics.DrawString(Prinfo, New Font("Calibri", 8, FontStyle.Regular), Brushes.Gainsboro, New RectangleF(5, 50, 200, 200), rt)
        e.Graphics.DrawString("Total residents of this Barangay. Extracted from database as recorded in this office.", New Font("Calibri", 9, FontStyle.Italic), Brushes.Black, w * 0.5, 1130, sf)


    End Sub

    Function _ADDHOUSEHOLDS() As Single
        Using conn As New SqlConnection(ConnStr)
            Using cmd As New SqlCommand("INSERT INTO RFS_CHILD(ID,CFNAME,CMNAME,CLNAME,CEXTNAME,CPB,CRDB,CGENDER,CCSTAT,CCTZ,COCCUP) VALUES(@A, @B, @C, @D, @E, @F, @G, @H, @I, @J, @K)", conn)

            End Using
        End Using
    End Function

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim csz As New PaperSize("ID Size", 350, 230)
        Dim pd As New PrintPreviewDialog
        RePrintIDfromGv.PrinterSettings = New PrinterSettings
        RePrintIDfromGv.PrinterSettings.DefaultPageSettings.PaperSize = csz
        pd.Name = "Print Resident ID"
        RePrintIDfromGv.DocumentName = "Resident ID"
        pd.Document = RePrintIDfromGv
        pd.ShowDialog()


        With DataGridView1.CurrentRow

            Dim TheInfo = .Cells(4).Value & " " & .Cells(5).Value.ToString(0) & ". " & .Cells(3).Value
            With ListBox1.Items
                .Add("RePrint Resident ID: [" & DataGridView1.CurrentRow.Cells(1).Value & "]" & TheInfo)
                .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
            End With
        End With
    End Sub

    Private Sub RePrintIDfromGv_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles RePrintIDfromGv.PrintPage

        Dim csz As New PaperSize("ID Size", 350, 230)
        prtID.PrinterSettings.DefaultPageSettings.PaperSize = csz
        Dim i As Integer
        'Head Colors
        Dim ThisColor As New SolidBrush(My.Settings.CIDHEADTEXTC)
        Dim ToThisColor As Color = New Pen(ThisColor).Color
        Dim ThisBrush = New Pen(ThisColor)
        Dim HeadTextColor = New SolidBrush(ThisBrush.Color)

        'Body Text Color
        Dim BThisColor As New SolidBrush(My.Settings.CIDBODYTEXTC)
        Dim BToThisColor As Color = New Pen(BThisColor).Color
        Dim BThisBrush = New Pen(BToThisColor)
        Dim BodyTextColor = New SolidBrush(BThisBrush.Color)

        'Text Formatting
        Dim sf As StringFormat = StringFormat.GenericTypographic
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word

        'Body Background Color
        Dim BcThisColor As New SolidBrush(My.Settings.CIDBACKC)
        Dim BcToThisColor As Color = New Pen(BcThisColor).Color
        Dim BcThisBrush = New Pen(BcToThisColor)
        Dim BodyBackColor = New SolidBrush(BcThisBrush.Color)
        If My.Settings.IDCheckBox = False Then
            Select Case My.Settings.IDCOMBOSELECT
                Case "Basic_1"
                    e.Graphics.DrawImage(My.Resources.Basic_1, 0, 0)
                Case "Basic_2"
                    e.Graphics.DrawImage(My.Resources.Basic_2, 0, 0)
                Case "Basic_3"
                    e.Graphics.DrawImage(My.Resources.Basic_3, 0, 0)
                Case "Classic_1"
                    e.Graphics.DrawImage(My.Resources.Classic_1, 0, 0)
                Case "Classic_2"
                    e.Graphics.DrawImage(My.Resources.Classic_2, 0, 0)
                Case "Classic_3"
                    e.Graphics.DrawImage(My.Resources.Classic_3, 0, 0)
                Case "Modern_1"
                    e.Graphics.DrawImage(My.Resources.Modern_1, 0, 0)
                Case "Modern_2"
                    e.Graphics.DrawImage(My.Resources.Modern_2, 0, 0)
                Case "Modern_3"
                    e.Graphics.DrawImage(My.Resources.Modern_3, 0, 0)
            End Select
        Else
            If Not My.Settings.IDBI = String.Empty Then
                e.Graphics.DrawImage(Image.FromFile(My.Settings.IDBI), 0, 0)

            ElseIf Not My.Settings.CIDBACKC = Color.Empty Then
                e.Graphics.FillRectangle(BodyBackColor, New Rectangle(0, 0, 350, 230))
            Else
                e.Graphics.DrawImage(My.Resources.bg3, 0, 0)
            End If
        End If

        Dim today = DateTime.Today
        Dim day = today.Day
        Dim month = today.Month.ToString("MMMM")
        Dim year = today.Year
        Dim pinF As New PointF(20, 20)
        e.Graphics.DrawRectangle(Pens.Gainsboro, 0, 0, 350, 230)
        'Logo
        If Not My.Settings.BRNGYLOGO = String.Empty Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.BRNGYLOGO), 5, 5, 50, 50)
        End If
        Try
            Dim conn As New SqlConnection(ConnStr)

            With DataGridView1
                Dim cmd As New SqlCommand("SELECT ID, RLNAME, RFNAME, RMNAME, RDB, RCS, RSEX, VIDNO FROM RFS_PID WHERE ID=" & .CurrentRow.Cells(1).Value & "", conn)
                conn.Open()
                Dim reader As SqlDataReader = cmd.ExecuteReader()
                While reader.Read()
                    'The QR Code
                    Dim qrcode As KeepAutomation.Barcode.Bean.BarCode = New KeepAutomation.Barcode.Bean.BarCode
                    qrcode.Symbology = KeepAutomation.Barcode.Symbology.QRCode
                    qrcode.QRCodeVersion = KeepAutomation.Barcode.QRCodeVersion.V5
                    qrcode.QRCodeECL = KeepAutomation.Barcode.QRCodeECL.H
                    qrcode.QRCodeDataMode = KeepAutomation.Barcode.QRCodeDataMode.Auto
                    qrcode.CodeToEncode = reader("ID").ToString
                    qrcode.BarcodeUnit = KeepAutomation.Barcode.BarcodeUnit.Cm
                    qrcode.DPI = 100
                    qrcode.X = 2
                    qrcode.Y = 2
                    qrcode.LeftMargin = 8
                    qrcode.RightMargin = 8
                    qrcode.TopMargin = 8
                    qrcode.BottomMargin = 8
                    Dim QRcodePrint As Image = qrcode.generateBarcodeToBitmap
                    e.Graphics.DrawImage(QRcodePrint, 265, 145, 80, 80)
                    e.Graphics.DrawString(reader("VIDNO").ToString, New Font("Calibri", 10, FontStyle.Bold), Brushes.Black, 5, 208, StringFormat.GenericDefault)
                    Dim TheInfo = Trim(reader("RFNAME").ToString).ToUpper & " " & Trim(reader("RMNAME").ToString(0)).ToUpper & ". " & Trim(reader("RLNAME").ToString).ToUpper & vbCrLf & vbCrLf & _
                                      Trim(reader("RDB").ToString).ToUpper & vbCrLf & vbCrLf & _
                                      Trim(reader("RSEX").ToString).ToUpper
                    e.Graphics.DrawString(TheInfo, New Font("Calibri", 9, FontStyle.Bold), BodyTextColor, 115, 90)
                    e.Graphics.DrawString("R-" & reader("ID").ToString, New Font("Calibri", 10, FontStyle.Regular), HeadTextColor, 222, 30)

                End While
            End With
            conn.Close()
        Catch ex As Exception

        Finally
            Dim connn As New SqlConnection(ConnStr)
            For i = 0 To DataGridView1.Rows.Count - 1
                Dim imgcmd As New SqlCommand("SELECT IMAGE FROM LOGGER WHERE ID=" & DataGridView1.Rows(i).Cells(1).Value & "", connn)
                connn.Open()
                Dim sdr As SqlDataReader = imgcmd.ExecuteReader()
                While sdr.Read()
                    'e.Graphics.DrawImage(sdr("IMAGE"), 6, 81, 98, 120)
                End While
                connn.Close()
            Next


        End Try

        'ID Information 
        Dim HeaderText = "Republic of the Philippines" & vbCrLf & _
                          Trim(My.Settings.BP102).ToUpper & vbCrLf & _
                          Trim(My.Settings.BCM103).ToUpper
        Dim Brngy = Trim(My.Settings.BAR).ToUpper
        Dim InfoLabel = "Firstname/M.I/Lastname" & vbCrLf & vbCrLf & vbCrLf & _
                            "Birthdate" & vbCrLf & vbCrLf & vbCrLf & _
                            "Status"

        Dim TheSpice = "___________________" & vbCrLf & "Signature"
        e.Graphics.DrawString(TheSpice, New Font("Calibri", 8, FontStyle.Regular), Brushes.Black, 115, 175)

        Dim datetoday = DateTime.Now.ToString("MM/dd/yyyy")
        Dim datelast = DateTime.Now.AddYears(3).ToString("MM/dd/yyyy")

        e.Graphics.DrawString(datetoday, New Font("Calibri", 6, FontStyle.Bold), Brushes.Black, 320, 130, sf)
        e.Graphics.DrawString(datelast, New Font("Calibri", 6, FontStyle.Bold), Brushes.Black, 320, 140, sf)


        e.Graphics.DrawString("RESIDENT ID", New Font("Calibri", 14, FontStyle.Bold), HeadTextColor, 220, 12)

        e.Graphics.DrawRectangle(Pens.Black, 5, 80, 100, 120)

        e.Graphics.DrawString(HeaderText, New Font("Calibri", 6, FontStyle.Regular), HeadTextColor, 60, 5)
        e.Graphics.DrawString(Brngy, New Font("Calibri", 8, FontStyle.Bold), HeadTextColor, 59, 35)
        e.Graphics.DrawString(InfoLabel, New Font("Calibri", 6, FontStyle.Italic), BodyTextColor, 115, 80)

        ' e.Graphics.FillRectangle(Brushes.Black, 5, 80, 100, 120)
        'Execute texts
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Activator.Show()
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        _insertolog("Transferred Business to " & TextBox31.Text, TextBox38.Text)
    End Sub
    Private Sub Chart2_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Chart2.MouseClick
        Dim cm As New ContextMenuStrip
        Dim sfd As New System.Windows.Forms.SaveFileDialog
        Chart2.ContextMenuStrip = cm
        cm.Items.Add("Save Image")

        With cm
            AddHandler .Items(0).Click, AddressOf sfd.ShowDialog
            With sfd
                .Title = "Save Chart"
                .Filter = "PNG (*.png)|*.png"
                .FileName = "Chart"
                .RestoreDirectory = True
                If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
                    Chart2.SaveImage(sfd.FileName, DataVisualization.Charting.ChartImageFormat.Png)
                End If
            End With
        End With
    End Sub

    Private Sub ResidencyChart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResidencyChart.Click

    End Sub

    Private Sub ResidencyChart_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ResidencyChart.MouseClick
        Dim cm As New ContextMenuStrip
        Dim sfd As New System.Windows.Forms.SaveFileDialog
        ResidencyChart.ContextMenuStrip = cm
        cm.Items.Add("Save Image")

        With cm
            AddHandler .Items(0).Click, AddressOf sfd.ShowDialog
            With sfd
                .Title = "Save Chart"
                .Filter = "PNG (*.png)|*.png"
                .FileName = "Chart"
                .RestoreDirectory = True
                If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
                    ResidencyChart.SaveImage(sfd.FileName, DataVisualization.Charting.ChartImageFormat.Png)
                End If
            End With
        End With
    End Sub

    Private Sub Label79_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label82_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click

    End Sub



    Private Sub IdentificationCardToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim csz As New PaperSize("ID Size", 350, 230)
        Dim pd As New PrintPreviewDialog
        RePrintIDfromGv.PrinterSettings = New PrinterSettings
        RePrintIDfromGv.PrinterSettings.DefaultPageSettings.PaperSize = csz
        pd.Name = "Print Resident ID"
        RePrintIDfromGv.DocumentName = "Resident ID"
        pd.Document = RePrintIDfromGv
        pd.ShowDialog()
        Button32.Show()


        Dim r = DataGridView1.Rows
        Dim TheInfo = r(0).Cells(4).Value & " " & r(0).Cells(5).Value.ToString(0) & ". " & r(0).Cells(3).Value
        With ListBox1.Items

            .Add("RePrint Resident ID: [" & r(0).Cells(1).Value & "]" & TheInfo)
            .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
        End With
    End Sub

    Private Sub ExcelxlsxToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExcelxlsxToolStripMenuItem.Click
        Dim sfd As New System.Windows.Forms.SaveFileDialog
        sfd.Filter = "Microsoft Excel (*.xlsx) |*xlsx"
        sfd.RestoreDirectory = True
        sfd.FileName = "ListOfResidents_" & DateTime.Now.ToString("MMMM_dd_yyyy")
        If sfd.ShowDialog = Forms.DialogResult.OK Then
            _exportDGV(Me.DataGridView1, sfd.FileName)
        End If
    End Sub

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        Dim sfd As New System.Windows.Forms.SaveFileDialog
        sfd.Filter = "Microsoft Excel (*.xlsx)|*.xlsx"
        sfd.RestoreDirectory = True
        If sfd.ShowDialog = Forms.DialogResult.OK Then
            _exportDGV(Me.DataGridView3, sfd.FileName)
        End If

    End Sub

    Private Sub Button14_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim sfd As New System.Windows.Forms.SaveFileDialog
        sfd.Filter = "Microsoft Excel (*.xlsx)|*.xlsx"
        sfd.RestoreDirectory = True
        If sfd.ShowDialog = Forms.DialogResult.OK Then
            _exportDGV(Me.BUSDGriview, sfd.FileName)
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Not ComboBox1.Text = "SINGLE" Then
            If MsgBox("Do you want to add siblings?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                HhChildren.Text = IDno.Text & " " & TextBox2.Text.Substring(0) & " " & TextBox3.Text & " " & TextBox1.Text
                SplitContainer2.Panel1.Enabled = False
                HhChildren.TextBox1.Text = IDno.Text
                HhChildren.Show()
            End If
        Else

        End If
    End Sub

    Private Sub PrintStatistics_EndPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintStatistics.EndPrint
        _insertolog("Demographic printed", GenerateRandomNumber(1234568, 99999999))
    End Sub

    Private Sub PrintStatistics_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintStatistics.PrintPage
        Static PageNo As Integer
        Dim sssss As New PaperSize("A4", 827, 1168)
        PrintStatistics.PrinterSettings = New PrinterSettings
        PrintStatistics.PrinterSettings.DefaultPageSettings.PaperSize = sssss

        Static startAt As Integer = 1

        Dim today = DateTime.Today
        Dim day = today.Day
        Dim month = today.Month.ToString("MMMM")
        Dim year = today.Year
        Dim pinF As New PointF(20, 20)
        'Fonts
        Dim RegFont As Font = New Font("Calibri", 12, FontStyle.Regular)
        Dim BoldFont As Font = New Font("Calibri", 12, FontStyle.Bold)
        Dim Title As Font = New Font("Calibri", 14, FontStyle.Bold)
        Dim DocFont As Font = New Font("Calibri", 16, FontStyle.Bold)
        'Text Alignment
        Dim sf As StringFormat = StringFormat.GenericTypographic
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word
        Dim rt As New StringFormat
        rt.FormatFlags = StringFormatFlags.DirectionVertical
        Dim cx, y, w, h As Integer
        cx = 414.5
        y = 20
        w = 827
        h = 1168

        Dim HeaderText = "Republic of the Philippines" & vbCrLf & _
                        "Department of the Interior and Local Government" & vbCrLf & _
                        Trim(My.Settings.BP102).ToUpper & vbCrLf & _
                        Trim(My.Settings.BCM103).ToUpper
        Dim Brngy = Trim(My.Settings.BAR).ToUpper

        Dim PresideOffice = "OFFICE OF THE BARANGAY CHAIRMAN"

        PageNo += 1
        'Has more pages
        Select Case PageNo
            'First Page
            Case 1
                If Trim(My.Settings.CMLOGO).Length > 1 Then
                    e.Graphics.DrawImage(Image.FromFile(My.Settings.CMLOGO), 650, 20, 120, 120)
                Else
                End If
                If Trim(My.Settings.BRNGYLOGO).Length > 1 Then
                    e.Graphics.DrawImage(Image.FromFile(My.Settings.BRNGYLOGO), 75, 20, 120, 120)
                Else
                End If

                'Align Header Texts
                e.Graphics.DrawLine(Pens.Black, 50, 150, 750, 150)
                e.Graphics.DrawString(HeaderText, RegFont, Brushes.Black, New Rectangle(cx * 0.5, 20, w * 0.52, 100), sf)
                e.Graphics.DrawString(Brngy, Title, Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 80), sf)

                e.Graphics.DrawString(PresideOffice, Title, Brushes.Black, w * 0.5, 170, sf)
                e.Graphics.DrawString("DEMOGRAPHIC FEATURES", DocFont, Brushes.Black, w * 0.5, 220, sf)
                'Print from Database
                Using conn As New SqlConnection(ConnStr)
                    Dim xx = 140
                    Dim yy = 400
                    Using cmd As New SqlCommand("SELECT COUNT(*) AS Nums, SUM(CASE WHEN DENTRY<=DATEADD(YEAR,-5,GETDATE()) THEN 1 ELSE 0 END) AS NEWPOP, SUM(CASE WHEN DENTRY<DATEADD(YEAR,-4, GETDATE()) THEN 1 ELSE 0 END) AS LY, SUM(CASE WHEN RSEX='MALE' THEN 1 ELSE 0 END) AS M, SUM(CASE WHEN RSEX='FEMALE' THEN 1 ELSE 0 END) AS F FROM RFS_PID", conn)
                        conn.Open()
                        Dim reader As SqlDataReader = cmd.ExecuteReader()
                        While reader.Read

                            'Demographic Features

                            ' Dim m = Chart2.Series("Population").Points(0).YValues(0) + Chart2.Series("Population").Points(1).YValues(0) + Chart2.Series("Population").Points(2).YValues(0) + Chart2.Series("Population").Points(3).YValues(0) + Chart2.Series("Population").Points(4).YValues(0) + Chart2.Series("Population").Points(5).YValues(0)
                            Dim p = (reader("Nums") / reader("LY")) ^ (1 / 5) - 1
                            'Dim vp = p / reader("NEWPOP")
                            Dim mean = p * 100
                            ' Dim final = mean / 5

                            e.Graphics.DrawString("1. Total Population", Title, Brushes.Black, 70, 270, StringFormat.GenericDefault)
                            Dim TotalPop = My.Settings.BAR & " has a total population of " & reader("Nums") & " as of the latest record for the year " & DateTime.Now.ToString("yyyy") & ", of which " & reader("M") & " are men while " & reader("F") & " are women. This indicated that the population raised to " & Math.Round(mean, 0, MidpointRounding.ToEven) & "% compared to the last " & DateTime.Now.AddYears(-5).ToString("yyyy") & " indicated in Figure 2 herebelow, which is " & reader("LY") & "."
                            e.Graphics.DrawString(TotalPop, RegFont, Brushes.Black, New Rectangle(90, 300, 680, 80), StringFormat.GenericDefault)
                            e.Graphics.DrawString("Figure 1. Distribution of Population by Age", BoldFont, Brushes.Black, w * 0.5, 380, sf)

                            'The Charts 1
                            e.Graphics.DrawRectangle(Pens.Black, New Rectangle(xx - 5, yy - 5, 540, 330))
                            Chart1.Printing.PrintPaint(e.Graphics, New Rectangle(xx, yy, 530, 320))

                            'The Chart 2
                            e.Graphics.DrawString("Figure 2. Average Annual (Compound) Growth Rates", BoldFont, Brushes.Black, w * 0.5, yy + 360, sf)
                            e.Graphics.DrawRectangle(Pens.Black, New Rectangle(xx - 5, yy + 375, 540, 330))
                            Chart2.Printing.PrintPaint(e.Graphics, New Rectangle(xx, yy + 380, 530, 320))
                        End While
                        conn.Close()
                    End Using
                End Using
                e.Graphics.DrawString(PageNo & "/2", RegFont, Brushes.Black, 780, 1130)

                e.HasMorePages = True

            Case 2
                'Print from Database
                Using conn As New SqlConnection(ConnStr)
                    Dim xx = 200
                    Dim yy = 200
                    Using cmd As New SqlCommand("SELECT SUM(CASE WHEN RCS!='SINGLE' THEN 1 ELSE 0 END) AS HH, SUM(CASE WHEN RCS='MARRIED' THEN 1 ELSE 0 END) AS MARRIED, SUM(CASE WHEN RCS='WIDOW' THEN 1 ELSE 0 END) AS WIDOW, SUM(CASE WHEN RCS='DIVORCED' THEN 1 ELSE 0 END) AS DIVORCED, SUM(CASE WHEN RCS='LIVE-IN' THEN 1 ELSE 0 END) AS LIVEIN, SUM(CASE WHEN RCS='SINGLE' THEN 1 ELSE 0 END) AS SINGLE FROM RFS_PID", conn)
                        conn.Open()
                        Dim reader As SqlDataReader = cmd.ExecuteReader()
                        While reader.Read

                            'Demographic Features
                            e.Graphics.DrawString("2. Total Households", Title, Brushes.Black, 70, 70, StringFormat.GenericDefault)
                            Dim TotalPop = "The total number of households in " & My.Settings.BAR & " is " & reader("HH") & ". These households were counted in the barangay. There are " & reader("MARRIED") & " married, " & reader("DIVORCED") & " divorced, " & reader("WIDOW") & " windower, and " & reader("LIVEIN") & " live-in couples."
                            e.Graphics.DrawString(TotalPop, RegFont, Brushes.Black, New Rectangle(90, 100, 680, 80), StringFormat.GenericDefault)
                            'Chart title
                            e.Graphics.DrawString("Figure 3. Marital Status", BoldFont, Brushes.Black, w * 0.5, yy - 20, sf)

                            'The Charts 1 - Marital Status
                            e.Graphics.DrawRectangle(Pens.Black, New Rectangle(xx - 5, yy - 5, 440, 230))
                            MaritalChart.Printing.PrintPaint(e.Graphics, New Rectangle(xx, yy, 430, 220))

                        End While
                        conn.Close()
                    End Using
                    Using cmd As New SqlCommand("SELECT COUNT(HOMAT) AS NUMS, SUM(CASE WHEN HOMAT='Concrete' THEN 1 ELSE 0 END) AS C, SUM(CASE WHEN HOMAT='Semi-Concrete' THEN 1 ELSE 0 END) AS SC, SUM(CASE WHEN HOMAT='Light' THEN 1 ELSE 0 END) AS L FROM RFS_HHM", conn)
                        conn.Open()
                        Dim reader As SqlDataReader = cmd.ExecuteReader()
                        While reader.Read
                            e.Graphics.DrawString("3. Housing Materials", Title, Brushes.Black, 70, yy + 260, StringFormat.GenericDefault)
                            Dim n = reader("C") + reader("SC") + reader("L")
                            Dim concrete = FormatPercent(reader("C") / n, -1)
                            Dim sc = FormatPercent(reader("SC") / n, -1)
                            Dim L = FormatPercent(reader("L") / n, -1)

                            Dim TotalHouseholds = "A total of " & reader("NUMS") & " households are counted in the barangay. The total figure of " & concrete & " are made up of concrete materials. There has " & sc & " that are made up of semi-concrete materials while " & L & " are made up of light materials. These households are situated in Figure 5."
                            e.Graphics.DrawString(TotalHouseholds, RegFont, Brushes.Black, New Rectangle(90, yy + 290, 680, 80), StringFormat.GenericDefault)

                            'The Chart 2 - Housing Materials
                            e.Graphics.DrawString("Figure 4. Housing Materials", BoldFont, Brushes.Black, w * 0.5, yy + 370, sf)
                            e.Graphics.DrawRectangle(Pens.Black, New Rectangle(xx - 5, yy + 385, 440, 230))
                            HousingChart.Printing.PrintPaint(e.Graphics, New Rectangle(xx, yy + 390, 430, 220))

                            'Residency Chart
                            e.Graphics.DrawString("Figure 5. Shelter of Residents", BoldFont, Brushes.Black, w * 0.5, yy + 640, sf)
                            e.Graphics.DrawRectangle(Pens.Black, New Rectangle(xx - 5, yy + 655, 440, 230))
                            ResidencyChart.Printing.PrintPaint(e.Graphics, New Rectangle(xx, yy + 660, 430, 220))

                        End While
                        conn.Close()
                        e.Graphics.DrawString(PageNo & "/2", RegFont, Brushes.Black, 780, 1130)
                        e.Graphics.DrawString("Printed on " & DateTime.Now.ToString("MMMM dd, yyyy") & " via Rinfosys.", RegFont, Brushes.Black, 70, 1130)
                    End Using
                End Using
                e.HasMorePages = False
                PageNo = 0
        End Select

    End Sub

    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        Dim pd As New PrintPreviewDialog
        pd.Name = "Poplation Report"

        PrintStatistics.DocumentName = "Popluation Report"
        pd.Document = PrintStatistics
        If pd.ShowDialog() = Forms.DialogResult.OK Then

            With ListBox1.Items
                .Add("Print Statistics")
                .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
            End With
        End If

    End Sub

    Private Sub prtProfile_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles prtProfile.PrintPage
        Static j As String = SearchBox.Text
        Dim today = DateTime.Today
        Dim day = today.Day
        Dim month = today.ToString("MMMM")
        Dim year = today.Year
        Dim pinF As New PointF(20, 20)
        'Fonts
        Dim RegFont As Font = New Font("Calibri", 12, FontStyle.Regular)
        Dim BoldFont As Font = New Font("Calibri", 12, FontStyle.Bold)
        Dim Title As Font = New Font("Calibri", 16, FontStyle.Bold)
        Dim DocFont As Font = New Font("Verdana", 20, FontStyle.Bold)
        'Text Alignment
        Dim sf As StringFormat = StringFormat.GenericTypographic
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word
        Dim rt As New StringFormat
        rt.FormatFlags = StringFormatFlags.DirectionVertical
        Dim cx, y, w, h As Integer
        cx = 425
        y = 20
        w = 850
        h = 1300
        If Trim(My.Settings.CMLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.CMLOGO), 650, 20, 120, 120)
        Else
        End If
        If Trim(My.Settings.BRNGYLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.BRNGYLOGO), 75, 20, 120, 120)
        Else
        End If
        Dim HeaderText = "Republic of the Philippines" & vbCrLf & _
                        "Department of the Interior and Local Government" & vbCrLf & _
                        Trim(My.Settings.BP102).ToUpper & vbCrLf & _
                        Trim(My.Settings.BCM103).ToUpper
        Dim Brngy = Trim(My.Settings.BAR).ToUpper

        Dim PresideOffice = "OFFICE OF THE BARANGAY CHAIRMAN"
        'Align texts
        e.Graphics.DrawLine(Pens.Black, 100, 150, 750, 150)
        e.Graphics.DrawString(HeaderText, RegFont, Brushes.Black, New Rectangle(cx * 0.5, 20, w * 0.52, 100), sf)
        e.Graphics.DrawString(Brngy, Title, Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 80), sf)
        e.Graphics.DrawString(PresideOffice, Title, Brushes.Black, New Rectangle(cx * 0.5, 70, w * 0.52, 200), sf)

        e.Graphics.DrawString(Trim(My.Settings.RESPROF).ToUpper, DocFont, Brushes.Black, New Rectangle(cx * 0.5, 120, w * 0.52, 200), sf)
        'The Tables
        Dim t As Integer = 270
        Dim x As Integer = 85

        e.Graphics.DrawRectangle(Pens.Black, New Rectangle(x, t, 530, 167))
        e.Graphics.DrawRectangle(Pens.Black, New Rectangle(x, t + 210, 680, 450))
        e.Graphics.DrawString("FAMILY COMPOSITION", BoldFont, Brushes.Black, New Rectangle(x, y + 420, 687, 40), sf)
        'VL1
        e.Graphics.DrawLine(Pens.Black, x + 50, t + 210, x + 50, y + 910)
        'VL2
        e.Graphics.DrawLine(Pens.Black, x + 380, t + 210, x + 380, y + 910)
        'VL3
        e.Graphics.DrawLine(Pens.Black, x + 430, t + 210, x + 430, y + 910)
        'VL4
        e.Graphics.DrawLine(Pens.Black, x + 550, t + 210, x + 550, y + 910)
        'HL1
        e.Graphics.DrawLine(Pens.Black, x, t + 240, x + 680, t + 240)
        'Table HeaderTexts
        e.Graphics.DrawString("NO.", BoldFont, Brushes.Black, x + 5, t + 215, StringFormat.GenericDefault)
        e.Graphics.DrawString("NAME", BoldFont, Brushes.Black, x + 195, t + 215, StringFormat.GenericDefault)
        e.Graphics.DrawString("AGE", BoldFont, Brushes.Black, x + 390, t + 215, StringFormat.GenericDefault)
        e.Graphics.DrawString("RELATIONSHIP", BoldFont, Brushes.Black, x + 435, t + 215, StringFormat.GenericDefault)
        e.Graphics.DrawString("OCCUPATION", BoldFont, Brushes.Black, x + 560, t + 215, StringFormat.GenericDefault)

        'The Profile
        Dim LabelInfo As System.String = "NAME: " & vbCrLf & _
                           "BIRTHDATE: " & vbCrLf & _
                           "CIVIL STATUS: " & vbCrLf & _
                           "GENDER: " & vbCrLf & _
                           "OCCUPATION: " & vbCrLf & _
                           "ADDRESS: " & vbCrLf & _
                           "RESIDENTIAL: "
        e.Graphics.DrawString(LabelInfo, RegFont, Brushes.Black, New Rectangle(x + 5, 275, 780, 150), StringFormat.GenericDefault)
        e.Graphics.DrawString(Trim("CTRL NO. " & j), BoldFont, Brushes.Black, New Rectangle(cx * 0.5, 145, w * 0.52, 200), sf)
        Using conn As New SqlConnection(ConnStr)
            Using cmd As New SqlCommand("SELECT RLNAME, RFNAME, RMNAME, RDB, RCS, RSEX, RSADD, ROCCU, RBT FROM RFS_PID WHERE ID='" & SearchBox.Text & "'", conn)
                conn.Open()
                Dim reader As SqlDataReader = cmd.ExecuteReader()
                While reader.Read
                    Dim Info = Trim(reader("RFNAME") & " " & reader("RMNAME").ToString(0) & ". " & reader("RLNAME") & vbCrLf & _
                           reader("RDB") & vbCrLf & _
                           reader("RCS") & vbCrLf & _
                           reader("RSEX") & vbCrLf & _
                           reader("ROCCU") & vbCrLf & _
                           reader("RSADD") & vbCrLf & _
                           reader("RBT")).ToUpper

                    e.Graphics.DrawString(Info, BoldFont, Brushes.Black, New Rectangle(x + 105, 275, 780, 150), StringFormat.GenericDefault)

                    e.Graphics.DrawRectangle(Pens.Black, New Rectangle(627, 270, 136, 167))
                End While
               
                conn.Close()
            End Using
            Using cmd As New SqlCommand("SELECT IMAGE FROM LOGGER WHERE ID='" & SearchBox.Text & "'", conn)
                conn.Open()
                Dim reader As SqlDataReader = cmd.ExecuteReader()
                While reader.Read
                    Dim ImgStream As New IO.MemoryStream(CType(reader("IMAGE"), Byte()))
                    e.Graphics.DrawImage(Image.FromStream(ImgStream), New Rectangle(627, 270, 136, 167))
                End While
                conn.Close()
            End Using

        End Using
    End Sub

    Private Sub AsOfDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AsOfDate.ValueChanged
        __POPULATION()
    End Sub

    Private Sub TableLayoutPanel3_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles TableLayoutPanel3.Paint

    End Sub

    Private Sub Button36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button36.Click
        Dim csz As New PaperSize("US Letter Size", 850, 1300)
        prtProfile.PrinterSettings = New PrinterSettings
        prtProfile.PrinterSettings.DefaultPageSettings.PaperSize = csz

        Dim pd As New PrintPreviewDialog
        pd.Name = "Printing Profile"
        prtProfile.DocumentName = "Barangay Resident's Profile"
        pd.Document = prtProfile
        pd.ShowDialog()
        Button32.Show()
        With ListBox1.Items
            .Add("Print Resident's Profile")
            .Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"))
        End With
    End Sub

    Private Sub PrintEconomy_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintEconomy.PrintPage
        PrintEconomy.Dispose()
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim sfd As New System.Windows.Forms.SaveFileDialog

        sfd.Filter = "Microsoft Excel (*.xlsx) |*xlsx"
        sfd.RestoreDirectory = True

        sfd.FileName = "ListOfResidents_" & DateTime.Now.ToString("MMMM_dd_yyyy")
        If sfd.ShowDialog = Forms.DialogResult.OK Then
            _exportDGV(Me.DataGridView1, sfd.FileName)
        End If
    End Sub


    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button34.Click
        _TRUNCATE("RFS_BUS")
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        LogData.Show()
    End Sub

    Private Sub DeceasedToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeceasedToolStripMenuItem.Click
        Dim rip = DataGridView1.SelectedRows(0).Cells(4).Value & " " & DataGridView1.SelectedRows(0).Cells(5).Value & " " & DataGridView1.SelectedRows(0).Cells(3).Value

        IsDeceasedForm.TextBox1.Text = DataGridView1.SelectedRows(0).Cells(1).Value
        IsDeceasedForm.TextBox2.Text = rip
        IsDeceasedForm.Show()
    End Sub

    Private Sub PrintNewBorn_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintNewBorn.PrintPage
        Dim today = DateTime.Today
        Dim day = today.Day
        Dim month = today.ToString("MMMM")
        Dim year = today.Year
        Dim pinF As New PointF(20, 20)

        'Fonts
        Dim RegFont As Font = New Font("Calibri", 12, FontStyle.Regular)
        Dim BoldFont As Font = New Font("Verdana", 12, FontStyle.Bold)
        Dim Title As Font = New Font("Calibri", 16, FontStyle.Bold)
        Dim DocFont As Font = New Font("Verdana", 20, FontStyle.Bold)

        'Page area
        'Dim rect As New Rectangle()

        'Barcode
        Dim barcode As KeepAutomation.Barcode.Bean.BarCode = New KeepAutomation.Barcode.Bean.BarCode
        barcode.Symbology = KeepAutomation.Barcode.Symbology.Codabar
        barcode.DisplayText = True
        barcode.BarcodeUnit = KeepAutomation.Barcode.BarcodeUnit.Cm
        barcode.CodeToEncode = DIDno.Text
        barcode.X = 2
        barcode.Y = 40
        barcode.BarCodeWidth = 100
        barcode.BarCodeHeight = 70
        barcode.DPI = 100

        Dim ThisBarcode As Image = barcode.generateBarcodeToBitmap

        'Text Alignment
        Dim sf As StringFormat = StringFormat.GenericTypographic
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word

        Dim rt As New StringFormat
        rt.FormatFlags = StringFormatFlags.DirectionVertical

        Dim cx, y, w, h As Integer
        cx = 414.5
        y = 20
        w = 827
        h = 1168
        If Trim(My.Settings.CMLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.CMLOGO), 650, 20, 120, 120)
        Else

        End If
        If Trim(My.Settings.BRNGYLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.BRNGYLOGO), 75, 20, 120, 120)

        Else

        End If

        Dim Purpose = TextBox19.Text
        Select Case CheckBox2.Checked
            Case True
                Purpose = TextBox19.Text.ToUpper
            Case False
                Purpose = "APPLICATION"
        End Select

        Dim LiveBirth = "DATE OF BIRTH: " & vbCrLf & _
                "PLACE OF BIRTH: " & vbCrLf & _
                ""

        Dim ThisName = Trim(TextBox2.Text & " " & TextBox3.Text.Substring(0) & " " & TextBox1.Text).ToUpper
        Dim LetterBody = "          THIS IS TO CERTIFY that " & ThisName & ", " & ComboBox2.Text & ", " & ComboBox3.Text & " was born on " & RDB1.Text & " at " & Trim(My.Settings.BAR).ToUpper & ", " & Trim(My.Settings.BCM103).ToUpper & ", " & Trim(My.Settings.BP102).ToUpper & ", Philippines. " & vbCrLf & vbCrLf & _
                      "         It is further certifies that the above-mentioned person is a citizen of " & ComboBox4.Text & ", a resident of this Barangay." & vbCrLf & vbCrLf & _
                      "         This certification is issued to support above-named child for " & Trim(Purpose).ToUpper & " and whatsoever legal purposes that may serve him/her best." & vbCrLf & vbCrLf & _
                      "         Done this " & day & " day of " & month & ", " & year & " at Barangay Hall, " & Trim(My.Settings.BAR).ToUpper & ", " & Trim(My.Settings.BCM103).ToUpper & ", " & Trim(My.Settings.BP102).ToUpper & ", Philippines."
        Dim HeaderText = "Republic of the Philippines" & vbCrLf & _
                         "Philippine Statistics Authority" & vbCrLf & _
                         Trim(My.Settings.BP102).ToUpper & vbCrLf & _
                         Trim(My.Settings.BCM103).ToUpper

        Dim Brngy = Trim(My.Settings.BAR).ToUpper
        Dim PresideOffice = "OFFICE OF THE BARANGAY CHAIRMAN"
        Dim ToWhom = "TO WHOM IT MAY CONCERN:"
        Dim sss = "Certified By: " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                    "_____________________" & vbCrLf & _
                     " " & My.Settings.BChair.ToUpper
        Dim Prinfo = "Printed on: " & DateTime.Today.ToShortDateString & " via " & System.Windows.Forms.Application.ProductName & " " & System.Windows.Forms.Application.ProductVersion
        e.Graphics.DrawLine(Pens.Black, 50, 150, 750, 150)

        'Align texts
        e.Graphics.DrawString(HeaderText, RegFont, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 20, w * 0.52, 100), sf)
        e.Graphics.DrawString(Brngy, Title, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 80), sf)
        e.Graphics.DrawString(PresideOffice, Title, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 200), sf)

        e.Graphics.DrawString("BIRTH CERTIFICATION", DocFont, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 100, w * 0.52, 300), sf)
        e.Graphics.DrawString(ToWhom, BoldFont, System.Drawing.Brushes.Black, New Rectangle(75, y * 16, w - 150, 100))
        e.Graphics.DrawString(LetterBody, RegFont, System.Drawing.Brushes.Black, New Rectangle(75, y * 18, w - 150, 450), StringFormat.GenericDefault)

        'Officials' Certification
        e.Graphics.DrawString(sss, BoldFont, System.Drawing.Brushes.Black, 480, 700)

        'Document Identification
        e.Graphics.DrawString("__________________", RegFont, System.Drawing.Brushes.Black, 78, 920)
        e.Graphics.DrawString("Specimen Signature", RegFont, System.Drawing.Brushes.Black, 78, 940)
        e.Graphics.DrawImage(ThisBarcode, 70, 1000, 170, 45)
        'Printed on datetime
        e.Graphics.DrawString(Prinfo, New Font("Calibri", 8, FontStyle.Regular), System.Drawing.Brushes.Gainsboro, New RectangleF(5, 50, 200, 200), rt)
        e.Graphics.DrawString("Doc. No.:" & DIDno.Text, RegFont, System.Drawing.Brushes.Black, 78, 980)
        If ConfigDiag.CheckBox2.Checked = True Then
            Dim reso = "Any alteration on this Certificate may null and void. Provided under Barangay " & My.Settings.RESNO & ", this Barangay."
            e.Graphics.DrawString(reso, New Font("Calibri", 8, FontStyle.Italic), System.Drawing.Brushes.Black, New Rectangle(75, 1050, w - 150, 50), sf)
        End If
    End Sub

    Private Sub NewBornCheck_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewBornCheck.CheckedChanged
        If NewBornCheck.Checked = True Then
            Panel1.Enabled = False
            ComboBox1.Text = "SINGLE"
            CheckBox1.Enabled = False

        Else
            Panel1.Enabled = True
            ComboBox1.Text = ""
            CheckBox1.Enabled = True

        End If
    End Sub

    Private Sub RinfosysContextMenu_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles RinfosysContextMenu.Opening

    End Sub

    Private Sub DateTimePicker3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker3.ValueChanged
        If DateTime.Now = DateTimePicker3.Value Then
            __GrowthRates()
        Else

        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

    End Sub

    Private Sub PrintSoloParent_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintSoloParent.PrintPage
        Dim today = DateTime.Today
        Dim day = today.Day
        Dim month = today.ToString("MMMM")
        Dim year = today.Year
        Dim pinF As New PointF(20, 20)

        'Fonts
        Dim RegFont As Font = New Font("Calibri", 12, FontStyle.Regular)
        Dim BoldFont As Font = New Font("Verdana", 12, FontStyle.Bold)
        Dim Title As Font = New Font("Calibri", 16, FontStyle.Bold)
        Dim DocFont As Font = New Font("Verdana", 20, FontStyle.Bold)

        'Page area
        'Dim rect As New Rectangle()

        'Barcode
        Dim barcode As KeepAutomation.Barcode.Bean.BarCode = New KeepAutomation.Barcode.Bean.BarCode
        barcode.Symbology = KeepAutomation.Barcode.Symbology.Codabar
        barcode.DisplayText = True
        barcode.BarcodeUnit = KeepAutomation.Barcode.BarcodeUnit.Cm
        barcode.CodeToEncode = DIDno.Text
        barcode.X = 2
        barcode.Y = 40
        barcode.BarCodeWidth = 100
        barcode.BarCodeHeight = 70
        barcode.DPI = 100

        Dim ThisBarcode As Image = barcode.generateBarcodeToBitmap

        'Text Alignment
        Dim sf As StringFormat = StringFormat.GenericTypographic
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word

        Dim rt As New StringFormat
        rt.FormatFlags = StringFormatFlags.DirectionVertical

        Dim cx, y, w, h As Integer
        cx = 414.5
        y = 20
        w = 827
        h = 1168
        If Trim(My.Settings.CMLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.CMLOGO), 650, 20, 120, 120)
        Else

        End If
        If Trim(My.Settings.BRNGYLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.BRNGYLOGO), 75, 20, 120, 120)

        Else

        End If

        Dim Purpose = TextBox19.Text
        Select Case CheckBox2.Checked
            Case True
                Purpose = TextBox19.Text.ToUpper
            Case False
                Purpose = "APPLICATION"
        End Select

        Dim ThisName = Trim(TextBox2.Text & " " & TextBox3.Text.Substring(0) & " " & TextBox1.Text).ToUpper
        Dim LetterBody = "          THIS IS TO CERTIFY that per records available in this office, " & ThisName & " is a SOLO PARENT and a bonafide resident of " & Trim(My.Settings.BAR).ToUpper & ", " & Trim(My.Settings.BCM103).ToUpper & ", " & Trim(My.Settings.BP102).ToUpper & ", Philippines. " & vbCrLf & vbCrLf & _
                      "         It is further certifies that above - named person is known to me and has no pending and/or derogatory record filed against him/her as of this date." & vbCrLf & vbCrLf & _
                      "         This certification is issued for whatsoever legal purposes it may serve him/her best." & vbCrLf & vbCrLf & _
                      "         Done this " & day & " day of " & month & ", " & year & " at Barangay Hall, " & Trim(My.Settings.BAR).ToUpper & ", " & Trim(My.Settings.BCM103).ToUpper & ", " & Trim(My.Settings.BP102).ToUpper & ", Philippines."
        Dim HeaderText = "Republic of the Philippines" & vbCrLf & _
                         "Department of the Interior and Local Government" & vbCrLf & _
                         Trim(My.Settings.BP102).ToUpper & vbCrLf & _
                         Trim(My.Settings.BCM103).ToUpper

        Dim Brngy = Trim(My.Settings.BAR).ToUpper
        Dim PresideOffice = "OFFICE OF THE BARANGAY CHAIRMAN"
        Dim ToWhom = "TO WHOM IT MAY CONCERN:"
        Dim sss = "Certified By: " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                    "_____________________" & vbCrLf & _
                     " " & My.Settings.BChair.ToUpper
        Dim Prinfo = "Printed on: " & DateTime.Today.ToShortDateString & " via " & System.Windows.Forms.Application.ProductName & " " & System.Windows.Forms.Application.ProductVersion
        e.Graphics.DrawLine(Pens.Black, 50, 150, 750, 150)

        'Align texts
        e.Graphics.DrawString(HeaderText, RegFont, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 20, w * 0.52, 100), sf)
        e.Graphics.DrawString(Brngy, Title, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 80), sf)
        e.Graphics.DrawString(PresideOffice, Title, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 200), sf)

        e.Graphics.DrawString("C E R T I F I C A T I O N", DocFont, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 100, w * 0.52, 300), sf)
        e.Graphics.DrawString("SOLO PARENT (RA 8972)", BoldFont, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 120, w * 0.52, 300), sf)

        e.Graphics.DrawString(ToWhom, BoldFont, System.Drawing.Brushes.Black, New Rectangle(75, y * 16, w - 150, 100))
        e.Graphics.DrawString(LetterBody, RegFont, System.Drawing.Brushes.Black, New Rectangle(75, y * 18, w - 150, 450), StringFormat.GenericDefault)

        'Officials' Certification
        e.Graphics.DrawString(sss, BoldFont, System.Drawing.Brushes.Black, 480, 700)

        'Document Identification
        e.Graphics.DrawString("__________________", RegFont, System.Drawing.Brushes.Black, 78, 920)
        e.Graphics.DrawString("Specimen Signature", RegFont, System.Drawing.Brushes.Black, 78, 940)
        e.Graphics.DrawImage(ThisBarcode, 70, 1000, 170, 45)
        'Printed on datetime
        e.Graphics.DrawString(Prinfo, New Font("Calibri", 8, FontStyle.Regular), System.Drawing.Brushes.Gainsboro, New RectangleF(5, 50, 200, 200), rt)
        e.Graphics.DrawString("Doc. No.:" & DIDno.Text, RegFont, System.Drawing.Brushes.Black, 78, 980)
        If ConfigDiag.CheckBox2.Checked = True Then
            Dim reso = "Any alteration on this Certificate may null and void. Provided under Barangay " & My.Settings.RESNO & ", this Barangay."
            e.Graphics.DrawString(reso, New Font("Calibri", 8, FontStyle.Italic), System.Drawing.Brushes.Black, New Rectangle(75, 1050, w - 150, 50), sf)
        End If
    End Sub

    Private Sub PrintHomeSharer_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintHomeSharer.PrintPage
        Dim today = DateTime.Today
        Dim day = today.Day
        Dim month = today.ToString("MMMM")
        Dim year = today.Year
        Dim pinF As New PointF(20, 20)

        'Fonts
        Dim RegFont As Font = New Font("Calibri", 12, FontStyle.Regular)
        Dim BoldFont As Font = New Font("Verdana", 12, FontStyle.Bold)
        Dim Title As Font = New Font("Calibri", 16, FontStyle.Bold)
        Dim DocFont As Font = New Font("Verdana", 20, FontStyle.Bold)

        'Page area
        'Dim rect As New Rectangle()

        'Barcode
        Dim barcode As KeepAutomation.Barcode.Bean.BarCode = New KeepAutomation.Barcode.Bean.BarCode
        barcode.Symbology = KeepAutomation.Barcode.Symbology.Codabar
        barcode.DisplayText = True
        barcode.BarcodeUnit = KeepAutomation.Barcode.BarcodeUnit.Cm
        barcode.CodeToEncode = DIDno.Text
        barcode.X = 2
        barcode.Y = 40
        barcode.BarCodeWidth = 100
        barcode.BarCodeHeight = 70
        barcode.DPI = 100

        Dim ThisBarcode As Image = barcode.generateBarcodeToBitmap

        'Text Alignment
        Dim sf As StringFormat = StringFormat.GenericTypographic
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word

        Dim rt As New StringFormat
        rt.FormatFlags = StringFormatFlags.DirectionVertical

        Dim cx, y, w, h As Integer
        cx = 414.5
        y = 20
        w = 827
        h = 1168
        If Trim(My.Settings.CMLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.CMLOGO), 650, 20, 120, 120)
        Else

        End If
        If Trim(My.Settings.BRNGYLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.BRNGYLOGO), 75, 20, 120, 120)

        Else

        End If

        Dim Purpose = TextBox19.Text
        Select Case CheckBox2.Checked
            Case True
                Purpose = TextBox19.Text.ToUpper
            Case False
                Purpose = "APPLICATION"
        End Select

        Dim ThisName = Trim(TextBox2.Text & " " & TextBox3.Text.Substring(0) & " " & TextBox1.Text).ToUpper
        Dim LetterBody = "          THIS IS TO CERTIFY that per records available in this office, " & ThisName & " is a bonafide resident of " & Trim(My.Settings.BAR).ToUpper & ", " & Trim(My.Settings.BCM103).ToUpper & ", " & Trim(My.Settings.BP102).ToUpper & ", Philippines. " & vbCrLf & vbCrLf & _
                      "         It is further certifies that above - named person has no pending and/or derogatory record filed against him/her as of this date." & vbCrLf & vbCrLf & _
                      "         This certification is issued upon the request of the above mentioned person to support him/her for " & Trim(Purpose).ToUpper & " and whatsoever legal purposes that may serve him/her best." & vbCrLf & vbCrLf & _
                      "         Issued this " & day & " day of " & month & ", " & year & " at Barangay Hall, " & Trim(My.Settings.BAR).ToUpper & ", " & Trim(My.Settings.BCM103).ToUpper & ", " & Trim(My.Settings.BP102).ToUpper & ", Philippines."
        Dim HeaderText = "Republic of the Philippines" & vbCrLf & _
                         "Department of the Interior and Local Government" & vbCrLf & _
                         Trim(My.Settings.BP102).ToUpper & vbCrLf & _
                         Trim(My.Settings.BCM103).ToUpper

        Dim Brngy = Trim(My.Settings.BAR).ToUpper
        Dim PresideOffice = "OFFICE OF THE BARANGAY CHAIRMAN"
        Dim ToWhom = "TO WHOM IT MAY CONCERN:"
        Dim sss = "Certified By: " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                    "_____________________" & vbCrLf & _
                     " " & My.Settings.BChair.ToUpper
        Dim Prinfo = "Printed on: " & DateTime.Today.ToShortDateString & " via " & System.Windows.Forms.Application.ProductName & " " & System.Windows.Forms.Application.ProductVersion
        e.Graphics.DrawLine(Pens.Black, 50, 150, 750, 150)

        'Align texts
        e.Graphics.DrawString(HeaderText, RegFont, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 20, w * 0.52, 100), sf)
        e.Graphics.DrawString(Brngy, Title, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 80), sf)
        e.Graphics.DrawString(PresideOffice, Title, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 200), sf)

        e.Graphics.DrawString("CERTIFICATION", DocFont, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 100, w * 0.52, 300), sf)
        e.Graphics.DrawString("HOME SHARER", BoldFont, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 130, w * 0.52, 300), sf)
        e.Graphics.DrawString(ToWhom, BoldFont, System.Drawing.Brushes.Black, New Rectangle(75, y * 16, w - 150, 100))
        e.Graphics.DrawString(LetterBody, RegFont, System.Drawing.Brushes.Black, New Rectangle(75, y * 18, w - 150, 450), StringFormat.GenericDefault)

        'Officials' Certification
        e.Graphics.DrawString(sss, BoldFont, System.Drawing.Brushes.Black, 480, 700)

        'Document Identification
        e.Graphics.DrawString("__________________", RegFont, System.Drawing.Brushes.Black, 78, 920)
        e.Graphics.DrawString("Specimen Signature", RegFont, System.Drawing.Brushes.Black, 78, 940)
        e.Graphics.DrawImage(ThisBarcode, 70, 1000, 170, 45)
        'Printed on datetime
        e.Graphics.DrawString(Prinfo, New Font("Calibri", 8, FontStyle.Regular), System.Drawing.Brushes.Gainsboro, New RectangleF(5, 50, 200, 200), rt)
        e.Graphics.DrawString("Doc. No.:" & DIDno.Text, RegFont, System.Drawing.Brushes.Black, 78, 980)
        If ConfigDiag.CheckBox2.Checked = True Then
            Dim reso = "Any alteration on this Certificate may null and void. Provided under Barangay " & My.Settings.RESNO & ", this Barangay."
            e.Graphics.DrawString(reso, New Font("Calibri", 8, FontStyle.Italic), System.Drawing.Brushes.Black, New Rectangle(75, 1050, w - 150, 50), sf)
        End If
    End Sub

    Private Sub PrintIndigentPeople_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintIndigentPeople.PrintPage
        Dim today = DateTime.Today
        Dim day = today.Day
        Dim month = today.ToString("MMMM")
        Dim year = today.Year
        Dim pinF As New PointF(20, 20)

        'Fonts
        Dim RegFont As Font = New Font("Calibri", 12, FontStyle.Regular)
        Dim BoldFont As Font = New Font("Verdana", 12, FontStyle.Bold)
        Dim Title As Font = New Font("Calibri", 16, FontStyle.Bold)
        Dim DocFont As Font = New Font("Verdana", 20, FontStyle.Bold)

        'Page area
        'Dim rect As New Rectangle()

        'Barcode
        Dim barcode As KeepAutomation.Barcode.Bean.BarCode = New KeepAutomation.Barcode.Bean.BarCode
        barcode.Symbology = KeepAutomation.Barcode.Symbology.Codabar
        barcode.DisplayText = True
        barcode.BarcodeUnit = KeepAutomation.Barcode.BarcodeUnit.Cm
        barcode.CodeToEncode = DIDno.Text
        barcode.X = 2
        barcode.Y = 40
        barcode.BarCodeWidth = 100
        barcode.BarCodeHeight = 70
        barcode.DPI = 100

        Dim ThisBarcode As Image = barcode.generateBarcodeToBitmap

        'Text Alignment
        Dim sf As StringFormat = StringFormat.GenericTypographic
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word

        Dim rt As New StringFormat
        rt.FormatFlags = StringFormatFlags.DirectionVertical

        Dim cx, y, w, h As Integer
        cx = 414.5
        y = 20
        w = 827
        h = 1168
        If Trim(My.Settings.CMLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.CMLOGO), 650, 20, 120, 120)
        Else

        End If
        If Trim(My.Settings.BRNGYLOGO).Length > 1 Then
            e.Graphics.DrawImage(Image.FromFile(My.Settings.BRNGYLOGO), 75, 20, 120, 120)

        Else

        End If

        Dim Purpose = TextBox19.Text
        Select Case CheckBox2.Checked
            Case True
                Purpose = TextBox19.Text.ToUpper
            Case False
                Purpose = "APPLICATION"
        End Select

        Dim ThisName = Trim(TextBox2.Text & " " & TextBox3.Text.Substring(0) & " " & TextBox1.Text).ToUpper
        Dim LetterBody = "          THIS IS TO CERTIFY that " & ThisName & " of legal age, is a bonafide resident of " & Trim(My.Settings.BAR).ToUpper & ", " & Trim(My.Settings.BCM103).ToUpper & ", " & Trim(My.Settings.BP102).ToUpper & ", Philippines. " & vbCrLf & vbCrLf & _
                      "         The said person is known to me a good moral character and active member of the community who is belong to the low-income family." & vbCrLf & vbCrLf & _
                      "         This certification is issued upon the request of the said person for whatsoever legal purposes it may serve him/her best." & vbCrLf & vbCrLf & _
                      "         Done this " & day & " day of " & month & ", " & year & " at Barangay Hall, " & Trim(My.Settings.BAR).ToUpper & ", " & Trim(My.Settings.BCM103).ToUpper & ", " & Trim(My.Settings.BP102).ToUpper & ", Philippines."
        Dim HeaderText = "Republic of the Philippines" & vbCrLf & _
                         "Department of the Interior and Local Government" & vbCrLf & _
                         Trim(My.Settings.BP102).ToUpper & vbCrLf & _
                         Trim(My.Settings.BCM103).ToUpper

        Dim Brngy = Trim(My.Settings.BAR).ToUpper
        Dim PresideOffice = "OFFICE OF THE BARANGAY CHAIRMAN"
        Dim ToWhom = "TO WHOM IT MAY CONCERN:"
        Dim sss = "Certified By: " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                    "_____________________" & vbCrLf & _
                     " " & My.Settings.BChair.ToUpper
        Dim Prinfo = "Printed on: " & DateTime.Today.ToShortDateString & " via " & System.Windows.Forms.Application.ProductName & " " & System.Windows.Forms.Application.ProductVersion
        e.Graphics.DrawLine(Pens.Black, 50, 150, 750, 150)

        'Align texts
        e.Graphics.DrawString(HeaderText, RegFont, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 20, w * 0.52, 100), sf)
        e.Graphics.DrawString(Brngy, Title, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 80), sf)
        e.Graphics.DrawString(PresideOffice, Title, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 80, w * 0.52, 200), sf)

        e.Graphics.DrawString("CERTIFICATE", DocFont, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 100, w * 0.52, 300), sf)
        e.Graphics.DrawString("INDIGENIOUS PEOPLE", BoldFont, System.Drawing.Brushes.Black, New Rectangle(cx * 0.5, 120, w * 0.52, 300), sf)
        e.Graphics.DrawString(ToWhom, BoldFont, System.Drawing.Brushes.Black, New Rectangle(75, y * 16, w - 150, 100))
        e.Graphics.DrawString(LetterBody, RegFont, System.Drawing.Brushes.Black, New Rectangle(75, y * 18, w - 150, 450), StringFormat.GenericDefault)

        'Officials' Certification
        e.Graphics.DrawString(sss, BoldFont, System.Drawing.Brushes.Black, 480, 700)

        'Document Identification
        e.Graphics.DrawString("__________________", RegFont, System.Drawing.Brushes.Black, 78, 920)
        e.Graphics.DrawString("Specimen Signature", RegFont, System.Drawing.Brushes.Black, 78, 940)
        e.Graphics.DrawImage(ThisBarcode, 70, 1000, 170, 45)
        'Printed on datetime
        e.Graphics.DrawString(Prinfo, New Font("Calibri", 8, FontStyle.Regular), System.Drawing.Brushes.Gainsboro, New RectangleF(5, 50, 200, 200), rt)
        e.Graphics.DrawString("Doc. No.:" & DIDno.Text, RegFont, System.Drawing.Brushes.Black, 78, 980)
        If ConfigDiag.CheckBox2.Checked = True Then
            Dim reso = "Any alteration on this Certificate may null and void. Provided under Barangay " & My.Settings.RESNO & ", this Barangay."
            e.Graphics.DrawString(reso, New Font("Calibri", 8, FontStyle.Italic), System.Drawing.Brushes.Black, New Rectangle(75, 1050, w - 150, 50), sf)
        End If
    End Sub
End Class
