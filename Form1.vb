Public Class Form1

    Private Sub Form1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If

    End Sub
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Form1_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        Form2.Show()
        Form2.Location = Cursor.Position
        Form2.Location = Form2.Location
    End Sub

    Private Sub Form1_MouseMove(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        Form2.Size = Cursor.Position - Form2.Location
    End Sub

    Private Sub Form1_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        Form2.Hide()
        Me.Hide()
        Dim bounds As Form2
        Dim screenshot As System.Drawing.Bitmap
        Dim graph As Graphics
        bounds = Form2
        screenshot = New System.Drawing.Bitmap(bounds.Width, bounds.Height, System.Drawing.Imaging.PixelFormat.Format32bppRgb)
        graph = Graphics.FromImage(screenshot)
        graph.CopyFromScreen(Form2.Bounds.X, Form2.Bounds.Y, 0, 0, bounds.Size, CopyPixelOperation.SourceCopy)
        Form2.BackgroundImage = screenshot
        Dim sPath As New SaveFileDialog
        sPath.Filter = "Image (*.png)|*,*"
        sPath.ShowDialog()
        Dim bmp As Bitmap
        Try
            bmp = Form2.BackgroundImage
            bmp.Save(sPath.FileName + ".png")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Me.Close()
        My.Forms.Form2.Close()
    End Sub
End Class