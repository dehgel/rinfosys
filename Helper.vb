Imports System
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Windows.Forms

Class Helper

    Public Shared Sub SaveImageCapture(ByVal image As System.Drawing.Image)

        Dim s As New SaveFileDialog()
        s.FileName = "Image"
        s.DefaultExt = ".Jpg"
        s.Filter = "Image (.jpg)|*.jpg"
        If s.ShowDialog() = DialogResult.OK Then
            Dim filename As String = s.FileName
            Dim fstream As New FileStream(filename, FileMode.Create)
            image.Save(fstream, System.Drawing.Imaging.ImageFormat.Jpeg)
            fstream.Close()
        End If
    End Sub
  
End Class