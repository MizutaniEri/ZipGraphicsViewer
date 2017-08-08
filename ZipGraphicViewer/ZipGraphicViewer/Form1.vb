Imports System.IO.Compression
Imports System.IO

Public Class Form1
    Private ActiveExt() As String = {".jpg", ".bmp", ".png", "gif"}
    Private zipList As List(Of ZipArchiveEntry)
    Private index As Integer = -1
    Private zipFileName As String
    Private imageFIleName As String

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Dim openDialog As New OpenFileDialog()
        openDialog.Filter = "ZipFile|*.zip"
        If openDialog.ShowDialog() = DialogResult.OK Then
            ZipFileLoader(openDialog.FileName)
        End If
    End Sub

    ''' <summary>
    ''' ZIPファイルから画像ファイルの一覧を作成し、最初の画像を表示する
    ''' </summary>
    ''' <param name="FileName">ZIPファイル</param>
    ''' <remarks></remarks>
    Private Sub ZipFileLoader(FileName As String)
        Using zipArc As ZipArchive = ZipFile.OpenRead(FileName)
            zipFileName = FileName
            zipList = zipArc.Entries.Where(Function(entity) ActiveExt.Contains(Path.GetExtension(entity.Name).ToLower())).ToList()
            index += 1
            ZipView(index)
        End Using
    End Sub

    ''' <summary>
    ''' ZIPファイルから指定のインデックスのファイルの画像を表示する
    ''' </summary>
    ''' <param name="index"></param>
    ''' <remarks></remarks>
    Private Sub ZipView(index As Integer)
        imageFIleName = zipList(index).FullName
        ' ZIP書庫を開く()
        Using a As ZipArchive = ZipFile.OpenRead(zipFileName)
            ' ZipArchiveEntryを取得する
            Dim e As ZipArchiveEntry = a.GetEntry(imageFIleName)
            Dim stream As Stream = e.Open()
            PictureBox1.Image = Image.FromStream(stream, False, False)
        End Using
        If (FitScreenSizeToolStripMenuItem.Checked) Then
            ScreenFitZoom()
        ElseIf (WidthFitZoomToolStripMenuItem.Checked) Then
            ScreenFitWidthZoom()
        End If
        Text = zipFileName & " (" & (index + 1) & "/" & zipList.Count & ") - " & zipList(index).Name
    End Sub

    Private Sub NextToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NextToolStripMenuItem.Click
        NextListImage()
    End Sub

    Private Sub NextListImage()
        If zipList.Count <= (index + 1) Then
            index = 0
        Else
            index += 1
        End If
        ZipView(index)
    End Sub

    Private Sub BeforeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BeforeToolStripMenuItem.Click
        BeforeListImage()
    End Sub

    Private Sub BeforeListImage()
        If 0 > (index - 1) Then
            index = zipList.Count - 1
        Else
            index -= 1
        End If
        ZipView(index)
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim format As System.Drawing.Imaging.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg
            Select Case Path.GetExtension(SaveFileDialog1.FileName).ToLower()
                Case ".jpg"
                    format = System.Drawing.Imaging.ImageFormat.Jpeg
                Case ".bmp"
                    format = System.Drawing.Imaging.ImageFormat.Bmp
                Case ".png"
                    format = System.Drawing.Imaging.ImageFormat.Png
            End Select

            PictureBox1.Image.Save(SaveFileDialog1.FileName, format)
        End If
    End Sub

    Private Function ZoomGraphic(newSize As Size)
        '描画先とするImageオブジェクトを作成する
        Dim canvas As New Bitmap(newSize.Width, newSize.Height)
        'ImageオブジェクトのGraphicsオブジェクトを作成する
        Dim g As Graphics = Graphics.FromImage(canvas)

        '画像ファイルを読み込んで、Imageオブジェクトとして取得する
        Dim img As Image = PictureBox1.Image
        '画像のサイズを2倍にしてcanvasに描画する
        g.DrawImage(img, 0, 0, newSize.Width, newSize.Height)
        'Imageオブジェクトのリソースを解放する
        img.Dispose()

        'Graphicsオブジェクトのリソースを解放する
        g.Dispose()
        Return canvas
    End Function

    Private Sub GetScreenFitSize(ByRef imageSize As Size)
        ' スクリーンサイズの取得
        Dim Rect As Rectangle = Screen.GetWorkingArea(New Point(0, 0))
        ' スクリーンサイズの幅と高さ
        Dim screenX As Integer = Rect.Size.Width
        Dim screenY As Integer = Rect.Size.Height
        ' 画像の幅と高さ
        Dim imageX As Integer = imageSize.Width
        Dim imageY As Integer = imageSize.Height
        Dim newX As Integer = 0
        Dim newY As Integer = 0
        Dim RX As Integer = 0
        Dim RY As Integer = 0

        ' 画像の比率に沿った幅と高さ計算
        RX = imageX * screenY / imageY
        RY = imageY * screenX / imageX

        If ((RX < screenX) And (RY > screenY)) Then
            newX = RX
            newY = screenY
        Else
            newX = screenX
            newY = RY
        End If
        imageSize.Width = newX
        imageSize.Height = newY
    End Sub

    Private Sub GetScreenWidthFitSize(ByRef imageSize As Size)
        ' スクリーンサイズの取得
        Dim Rect As Rectangle = Screen.GetWorkingArea(New Point(0, 0))
        ' スクリーンサイズの幅と高さ
        Dim screenX As Integer = Rect.Size.Width
        Dim screenY As Integer = Rect.Size.Height
        ' 画像の幅と高さ
        Dim imageX As Integer = imageSize.Width
        Dim imageY As Integer = imageSize.Height
        Dim newX As Integer = 0
        Dim newY As Integer = 0
        Dim RX As Integer = 0
        Dim RY As Integer = 0

        ' 画像の比率に沿った幅と高さ計算
        RX = imageX * screenY / imageY
        RY = imageY * screenX / imageX

        'If ((RX < screenX) And (RY > screenY)) Then
        '    newX = RX
        '    newY = screenY
        'Else
        newX = screenX
        newY = RY
        'End If
        imageSize.Width = newX
        imageSize.Height = newY
    End Sub

    Private Sub FitScreenSizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FitScreenSizeToolStripMenuItem.Click
        If (FitScreenSizeToolStripMenuItem.Checked And WidthFitZoomToolStripMenuItem.Checked) Then
            WidthFitZoomToolStripMenuItem.Checked = False
        End If
        If (FitScreenSizeToolStripMenuItem.Checked) Then
            ScreenFitZoom()
        Else
            ZipView(index)
        End If
    End Sub

    Private Sub WidthFitZoomToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WidthFitZoomToolStripMenuItem.Click
        If (FitScreenSizeToolStripMenuItem.Checked And WidthFitZoomToolStripMenuItem.Checked) Then
            FitScreenSizeToolStripMenuItem.Checked = False
        End If
        If (WidthFitZoomToolStripMenuItem.Checked) Then
            ScreenFitWidthZoom()
        Else
            ZipView(index)
        End If
    End Sub

    Private Sub ScreenFitZoom()
        Dim imageSize As Drawing.Size
        imageSize.Width = PictureBox1.Image.Width
        imageSize.Height = PictureBox1.Image.Height
        GetScreenFitSize(imageSize)
        PictureBox1.Image = ZoomGraphic(imageSize)
    End Sub

    Private Sub ScreenFitWidthZoom()
        Dim imageSize As Drawing.Size
        imageSize.Width = PictureBox1.Image.Width
        imageSize.Height = PictureBox1.Image.Height
        GetScreenWidthFitSize(imageSize)
        PictureBox1.Image = ZoomGraphic(imageSize)
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        '左カーソルキーで次の画像表示
        If e.KeyCode = Keys.Left Then
            BeforeListImage()
            '右カーソル or Backspaceキーで前の画像表示
        ElseIf e.KeyCode = Keys.Right Or e.KeyCode = Keys.Back Then
            NextListImage()
        ElseIf e.KeyCode = Keys.Escape Then
            Close()
        End If
    End Sub

End Class
