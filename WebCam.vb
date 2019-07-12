Imports System
Imports System.IO
Imports System.Linq
Imports System.Text
Imports WebCam_Capture
Imports System.Collections.Generic

'Design by Pongsakorn Poosankam
Class WebCam
    Private webcam As WebCamCapture
    Private _FrameImage As New PictureBox
    Private FrameNumber As Integer = 30
    Public Sub InitializeWebCam(ByRef ImageControl As PictureBox)
        webcam = New WebCamCapture()
        webcam.Size = New System.Drawing.Size(137, 168)
        _FrameImage.Size = New System.Drawing.Size(137, 168)
        webcam.FrameNumber = CULng((0))
        webcam.TimeToCapture_milliseconds = FrameNumber
        AddHandler webcam.ImageCaptured, AddressOf webcam_ImageCaptured
        _FrameImage = ImageControl

    End Sub

    Private Sub webcam_ImageCaptured(ByVal source As Object, ByVal e As WebcamEventArgs)
        _FrameImage.Image = e.WebCamImage
    End Sub

    Public Sub Start()
        webcam.TimeToCapture_milliseconds = FrameNumber
        webcam.Start(0)
    End Sub

    Public Sub [Stop]()
        webcam.[Stop]()
        webcam.Dispose()

    End Sub

    Public Sub [Continue]()
        ' change the capture time frame
        webcam.TimeToCapture_milliseconds = FrameNumber

        ' resume the video capture from the stop
        webcam.Start(Me.webcam.FrameNumber)
    End Sub
End Class