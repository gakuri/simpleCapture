Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Windows.Forms

Module ScreenCapture

    Sub Main()
        ' すべてのスクリーンのサイズを合計して、キャプチャ用ビットマップを作成
        Dim totalWidth As Integer = 0
        Dim totalHeight As Integer = 0

        ' すべてのスクリーンの幅と高さを計算
        For Each scr As Screen In Screen.AllScreens
            totalWidth = Math.Max(totalWidth, scr.Bounds.X + scr.Bounds.Width)
            totalHeight = Math.Max(totalHeight, scr.Bounds.Y + scr.Bounds.Height)
        Next

        ' 全モニタをカバーするビットマップを作成
        Dim bmpScreenCapture As New Bitmap(totalWidth, totalHeight)
        Dim g As Graphics = Graphics.FromImage(bmpScreenCapture)

        ' 各モニタのスクリーンをキャプチャ
        For Each scr As Screen In Screen.AllScreens
            g.CopyFromScreen(scr.Bounds.X, scr.Bounds.Y, scr.Bounds.X, scr.Bounds.Y, scr.Bounds.Size)
        Next

        ' 日時形式を設定（yyyymmdd_hhnn）
        Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmm")

        ' ファイルパスを設定（例: C:\Screenshots\20241013_1130.jpg）
        Dim savePath As String = Path.Combine("C:\Screenshots", timestamp & ".jpg")

        ' フォルダが存在しない場合は作成
        If Not Directory.Exists("C:\Screenshots") Then
            Directory.CreateDirectory("C:\Screenshots")
        End If

        ' 画像を保存
        bmpScreenCapture.Save(savePath, ImageFormat.Jpeg)

        ' リソースを解放
        g.Dispose()
        bmpScreenCapture.Dispose()

        Console.WriteLine("スクリーンキャプチャが保存されました: " & savePath)
    End Sub

End Module
