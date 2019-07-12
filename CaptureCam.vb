Imports WebCam_Capture
Imports CG_Module
Imports UT_Module
Imports InstrumentDriverInterop
Imports ImageProcessing

Public Class CaptureCam
    Private IsFormBeingDragged As Boolean = False
    Private MouseDownX As Integer
    Private MouseDownY As Integer
    Private theimg As New PictureBox
    'The WebCamera
    Private webcam As WebCam

    Dim x As Integer
    Dim y As Integer
    Dim w As Integer
    Dim h As Integer

    'create a pen object
    Public cropPen As System.Drawing.Pen

    'select a default crop line size
    Public cropPenSize As Integer = 1 '2

    'will contain the dashStyle of the pen
    Public cropDashStyle As Drawing2D.DashStyle = Drawing2D.DashStyle.Solid

    'set a default crop line color
    Public cropPenColor As System.Drawing.Color = System.Drawing.Color.Aquamarine
    Dim cropBitmap As Bitmap

    Private Sub CaptureCam_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        x = Convert.ToInt32(5)
        y = Convert.ToInt32(5)
        w = Convert.ToInt32(138)
        h = Convert.ToInt32(169)
        Dim g As Graphics = theimg.CreateGraphics
        g.DrawRectangle(Pens.Black, x, y, w, h)

        webcam = New WebCam()
        theimg.Size = New System.Drawing.Size(138, 169)
        theimg = PictureBox1

        PictureBox1.Size = New System.Drawing.Size(138, 169)
        webcam.InitializeWebCam(theimg)
        webcam.Start()
        Button3.Text = "Disconnect"
        Button3.BackColor = System.Drawing.Color.Coral
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If x < 1 Then
            MsgBox("Invalid Size")
        End If
        Dim rect As System.Drawing.Rectangle = New System.Drawing.Rectangle(x, y, w, h)
        Dim bit As Bitmap = New Bitmap(theimg.Image, theimg.Width, theimg.Height)
        cropBitmap = New Bitmap(w, h)
        Dim g As Graphics = Graphics.FromImage(cropBitmap)
        g.DrawImage(bit, 0, 0, rect, GraphicsUnit.Pixel)
        cropBitmap.Save(DateTime.Now.ToString("MMddyyyy_hms"), System.Drawing.Imaging.ImageFormat.Png)
        'Rinfosys.PictureBox2.Image = cropBitmap
        'Rinfosys.PictureBox2.BackgroundImage = cropBitmap
        ' Rinfosys.ImgURL.Text = cropBitmap.FromFile()
        If Not File.Exists(System.Environment.CurrentDirectory + "/temp.jpg") Then
            cropBitmap.Save(System.Environment.CurrentDirectory + "/temp.jpg")
            Rinfosys.ImgURL.Text = System.Environment.CurrentDirectory + "/temp.jpg"
            Rinfosys.PictureBox2.ImageLocation = Rinfosys.ImgURL.Text
        Else

        End If
        


    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Select Case Button3.Text
            Case "Connect"
                webcam = New WebCam()
                theimg = PictureBox1
                PictureBox1.Size = New System.Drawing.Size(138, 170)
                webcam.InitializeWebCam(PictureBox1)
                webcam.Start()
                Button3.Text = "Disconnect"
                Button3.BackColor = System.Drawing.Color.Coral
            Case "Disconnect"
                webcam.Stop()
                Button3.BackColor = System.Drawing.Color.GreenYellow
                Button3.Text = "Connect"
        End Select
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        webcam.Continue()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        theimg.Size = New System.Drawing.Size(138, 170)
        Helper.SaveImageCapture(theimg.Image)
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
        webcam.Stop()
        Me.Dispose()

    End Sub

    'Move borderless Window
    Private Sub CaptureCam_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        If e.Button = MouseButtons.Left Then
            IsFormBeingDragged = True
            MouseDownX = e.X
            MouseDownY = e.Y
        End If
    End Sub

    Private Sub CaptureCam_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        If IsFormBeingDragged Then
            Dim temp As System.Drawing.Point = New System.Drawing.Point()
            temp.X = Me.Location.X + (e.X - MouseDownX)
            temp.Y = Me.Location.Y + (e.Y - MouseDownY)
            Me.Location = temp
            temp = Nothing
        End If
    End Sub

    Private Sub CaptureCam_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        If e.Button = MouseButtons.Left Then
            IsFormBeingDragged = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' theimg.Scale(New SizeF(138, 170))
        ' Rinfosys.PictureBox2.Image = theimg.Image
        webcam.Stop()
        WebCamCapture.CloseClipboard()
        Me.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then

            Me.TransparencyKey = System.Drawing.SystemColors.Highlight
        Else
            Me.TransparencyKey = System.Drawing.SystemColors.Window
            PictureBox1.BackColor = System.Drawing.SystemColors.Highlight
        End If
    End Sub
End Class