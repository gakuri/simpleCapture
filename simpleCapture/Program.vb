Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Runtime.CompilerServices
Imports System.Text

Module ScreenCapture
    <DllImport("user32.dll")>
    Public Function SetProcessDPIAware() As Boolean
    End Function

    ' Windows API: GetDpiForWindowをインポート
    <DllImport("user32.dll")>
    Private Function GetDpiForWindow(ByVal hWnd As IntPtr) As UInteger
    End Function

    Private Options As New Dictionary(Of String, String) From {
        {"-f", ""},
        {"-format", ""},
        {"-q", ""},
        {"-quality", ""},
        {"-p", ""},
        {"-path", ""},
        {"-c", ""},
        {"-color", ""}
    }

    Private compatibleFormat As String() = {"jpg"}
    Private formatType As String = "jpg"
    Private jpgQuality As Integer = 90
    Private saveDirectory As String = System.IO.Directory.GetCurrentDirectory()
    Private compatibleColorFormat As String() = {"full", "gray"}
    Private colorFormat As String = "full"

    Private Function setOptions() As Boolean
        Dim result As Boolean = True

        For Each optionValue As String In Environment.CommandLine.Split(" -").Skip(1)
            If InStr(optionValue, " ") <= 0 Then
                Console.WriteLine("オプション -" & optionValue & " に値が設定されていません")
                result = False
                Exit For
            End If

            If optionValue.Split(" ").Count > 2 Then
                Console.WriteLine("オプション -" & optionValue.Split(" ")(0) & " に１つ以上の値が設定されています")
                result = False
                Exit For
            End If

            If Options.ContainsKey("-" & optionValue.Split(" ")(0)) = False Then
                Console.WriteLine("オプション -" & optionValue.Split(" ")(0) & " というキーは想定されていません")
                result = False
                Exit For
            End If

            Options("-" & optionValue.Split(" ")(0)) = optionValue.Split(" ")(1)
        Next

        Return result

    End Function

    Private Function setFormatType() As Boolean

        Dim key As String
        Dim value As String

        If Options("-format") <> "" Then
            key = "-format"
            value = Options(key)
        ElseIf Options("-f") <> "" Then
            key = "-f"
            value = Options(key)
        Else
            'ファイルタイプの指定がないので、完了で返す
            Return True
        End If

        If Not compatibleFormat.Contains(Options(key)) Then
            Console.WriteLine($"{key}で指定された画像形式 {Options(key)}には対応していません。")
            Return False
        Else
            formatType = Options(key)
            Return True
        End If

    End Function
    Private Function setJpegQuality() As Boolean

        Dim quality As Integer

        Dim key As String
        Dim value As String

        If Options("-quality") <> "" Then
            key = "-quality"
            value = Options(key)
        ElseIf Options("-q") <> "" Then
            key = "-q"
            value = Options(key)
        Else
            'ファイルタイプの指定がないので、完了で返す
            Return True
        End If

        If Integer.TryParse(Options(key), quality) = False Then
            Console.WriteLine($"{key}で指定された画像の品質 {Options(key)}は、0-100の整数ではありません。")
            Return False
        ElseIf quality < 0 OrElse quality > 100 Then
            Console.WriteLine($"{key}で指定された画像の品質 {Options(key)}は、0-100の整数ではありません。")
            Return False
        Else
            jpgQuality = quality
            Return True
        End If

    End Function

    Private Function setSaveDirectory() As Boolean

        Dim key As String
        Dim value As String

        If Options("-path") <> "" Then
            key = "-path"
            value = Options(key)
        ElseIf Options("-p") <> "" Then
            key = "-p"
            value = Options(key)
        Else
            'パスの指定がないので、完了で返す
            Return True
        End If

        If IsValidPath(Options(key)) = False Then
            Console.WriteLine($"{key}で指定された保存先 {Options(key)}は、パスとして正しくありません")
            Return False
        Else
            saveDirectory = Options(key)
            Return True
        End If

    End Function

    ' パスが正しいかを判断する関数
    Function IsValidPath(ByVal path As String) As Boolean
        Try
            ' 無効な文字が含まれているかをチェック
            Dim invalidChars As Char() = System.IO.Path.GetInvalidPathChars()
            If path.IndexOfAny(invalidChars) >= 0 Then
                Return False
            End If

            ' Path クラスを使ってパスの妥当性を確認（例: フォーマットが正しいか）
            Dim fullPath As String = System.IO.Path.GetFullPath(path)

            ' パスが null や空文字列でないことを確認
            If String.IsNullOrWhiteSpace(path) Then
                Return False
            End If

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function setColorFormat() As Boolean

        Dim key As String
        Dim value As String

        If Options("-color") <> "" Then
            key = "-color"
            value = Options(key)
        ElseIf Options("-c") <> "" Then
            key = "-c"
            value = Options(key)
        Else
            'ファイルタイプの指定がないので、完了で返す
            Return True
        End If

        If Not compatibleColorFormat.Contains(Options(key)) Then
            Console.WriteLine($"{key}で指定された色指定 {Options(key)}には対応していません。")
            Return False
        Else
            colorFormat = Options(key)
            Return True
        End If

    End Function

    Sub Main()

        If setOptions() = False Then Environment.Exit(1)
        If setFormatType() = False Then Environment.Exit(1)
        If setJpegQuality() = False Then Environment.Exit(1)
        If setSaveDirectory() = False Then Environment.Exit(1)
        If setColorFormat() = False Then Environment.Exit(1)

        ' DPI対応を無効化（スケーリングを無視）
        SetProcessDPIAware()

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
        Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")

        ' ファイルパスを設定（例: C:\Screenshots\20241013_1130.jpg）
        Dim savePath As String = Path.Combine(saveDirectory, timestamp & ".jpg")

        ' フォルダが存在しない場合は作成
        If Not Directory.Exists(saveDirectory) Then
            Directory.CreateDirectory(saveDirectory)
        End If

        ' JPEGエンコーダを取得
        Dim jpgEncoder As ImageCodecInfo = GetEncoder(ImageFormat.Jpeg)

        ' エンコーダパラメータ（品質の指定）
        Dim encoderParams As New EncoderParameters(1)
        Dim qualityParam As New EncoderParameter(System.Drawing.Imaging.Encoder.Quality, jpgQuality) ' 85%の品質に設定
        encoderParams.Param(0) = qualityParam

        ' 画像を保存
        Select Case colorFormat
            Case "full"
                bmpScreenCapture.Save(savePath, jpgEncoder, encoderParams)
            Case "gray"
                ConvertToGrayscale(bmpScreenCapture).Save(savePath, jpgEncoder, encoderParams)
        End Select

        ' リソースを解放
        g.Dispose()
        bmpScreenCapture.Dispose()

        Console.WriteLine("スクリーンキャプチャが保存されました: " & savePath)
    End Sub

    ' 指定したイメージフォーマットのエンコーダを取得する関数
    Private Function GetEncoder(ByVal format As ImageFormat) As ImageCodecInfo
        Dim codecs As ImageCodecInfo() = ImageCodecInfo.GetImageDecoders()
        For Each codec As ImageCodecInfo In codecs
            If codec.FormatID = format.Guid Then
                Return codec
            End If
        Next
        Return Nothing
    End Function

    ' カラー画像をグレースケールに変換する関数
    Function ConvertToGrayscale(ByVal original As Bitmap) As Bitmap
        Dim grayscale As New Bitmap(original.Width, original.Height)

        For y As Integer = 0 To original.Height - 1
            For x As Integer = 0 To original.Width - 1
                ' ピクセルの色を取得
                Dim originalColor As Color = original.GetPixel(x, y)

                ' グレースケールの計算（輝度法）
                Dim gray As Integer = CInt(originalColor.R * 0.3 + originalColor.G * 0.59 + originalColor.B * 0.11)

                ' グレースケール値で新しい色を作成
                Dim grayColor As Color = Color.FromArgb(gray, gray, gray)

                ' 新しいピクセルにセット
                grayscale.SetPixel(x, y, grayColor)
            Next
        Next

        Return grayscale
    End Function

End Module
