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

    ' Windows API: GetDpiForWindow���C���|�[�g
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
                Console.WriteLine("�I�v�V���� -" & optionValue & " �ɒl���ݒ肳��Ă��܂���")
                result = False
                Exit For
            End If

            If optionValue.Split(" ").Count > 2 Then
                Console.WriteLine("�I�v�V���� -" & optionValue.Split(" ")(0) & " �ɂP�ȏ�̒l���ݒ肳��Ă��܂�")
                result = False
                Exit For
            End If

            If Options.ContainsKey("-" & optionValue.Split(" ")(0)) = False Then
                Console.WriteLine("�I�v�V���� -" & optionValue.Split(" ")(0) & " �Ƃ����L�[�͑z�肳��Ă��܂���")
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
            '�t�@�C���^�C�v�̎w�肪�Ȃ��̂ŁA�����ŕԂ�
            Return True
        End If

        If Not compatibleFormat.Contains(Options(key)) Then
            Console.WriteLine($"{key}�Ŏw�肳�ꂽ�摜�`�� {Options(key)}�ɂ͑Ή����Ă��܂���B")
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
            '�t�@�C���^�C�v�̎w�肪�Ȃ��̂ŁA�����ŕԂ�
            Return True
        End If

        If Integer.TryParse(Options(key), quality) = False Then
            Console.WriteLine($"{key}�Ŏw�肳�ꂽ�摜�̕i�� {Options(key)}�́A0-100�̐����ł͂���܂���B")
            Return False
        ElseIf quality < 0 OrElse quality > 100 Then
            Console.WriteLine($"{key}�Ŏw�肳�ꂽ�摜�̕i�� {Options(key)}�́A0-100�̐����ł͂���܂���B")
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
            '�p�X�̎w�肪�Ȃ��̂ŁA�����ŕԂ�
            Return True
        End If

        If IsValidPath(Options(key)) = False Then
            Console.WriteLine($"{key}�Ŏw�肳�ꂽ�ۑ��� {Options(key)}�́A�p�X�Ƃ��Đ���������܂���")
            Return False
        Else
            saveDirectory = Options(key)
            Return True
        End If

    End Function

    ' �p�X�����������𔻒f����֐�
    Function IsValidPath(ByVal path As String) As Boolean
        Try
            ' �����ȕ������܂܂�Ă��邩���`�F�b�N
            Dim invalidChars As Char() = System.IO.Path.GetInvalidPathChars()
            If path.IndexOfAny(invalidChars) >= 0 Then
                Return False
            End If

            ' Path �N���X���g���ăp�X�̑Ó������m�F�i��: �t�H�[�}�b�g�����������j
            Dim fullPath As String = System.IO.Path.GetFullPath(path)

            ' �p�X�� null ��󕶎���łȂ����Ƃ��m�F
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
            '�t�@�C���^�C�v�̎w�肪�Ȃ��̂ŁA�����ŕԂ�
            Return True
        End If

        If Not compatibleColorFormat.Contains(Options(key)) Then
            Console.WriteLine($"{key}�Ŏw�肳�ꂽ�F�w�� {Options(key)}�ɂ͑Ή����Ă��܂���B")
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

        ' DPI�Ή��𖳌����i�X�P�[�����O�𖳎��j
        SetProcessDPIAware()

        ' ���ׂẴX�N���[���̃T�C�Y�����v���āA�L���v�`���p�r�b�g�}�b�v���쐬
        Dim totalWidth As Integer = 0
        Dim totalHeight As Integer = 0

        ' ���ׂẴX�N���[���̕��ƍ������v�Z
        For Each scr As Screen In Screen.AllScreens
            totalWidth = Math.Max(totalWidth, scr.Bounds.X + scr.Bounds.Width)
            totalHeight = Math.Max(totalHeight, scr.Bounds.Y + scr.Bounds.Height)
        Next

        ' �S���j�^���J�o�[����r�b�g�}�b�v���쐬
        Dim bmpScreenCapture As New Bitmap(totalWidth, totalHeight)
        Dim g As Graphics = Graphics.FromImage(bmpScreenCapture)

        ' �e���j�^�̃X�N���[�����L���v�`��
        For Each scr As Screen In Screen.AllScreens
            g.CopyFromScreen(scr.Bounds.X, scr.Bounds.Y, scr.Bounds.X, scr.Bounds.Y, scr.Bounds.Size)
        Next

        ' �����`����ݒ�iyyyymmdd_hhnn�j
        Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")

        ' �t�@�C���p�X��ݒ�i��: C:\Screenshots\20241013_1130.jpg�j
        Dim savePath As String = Path.Combine(saveDirectory, timestamp & ".jpg")

        ' �t�H���_�����݂��Ȃ��ꍇ�͍쐬
        If Not Directory.Exists(saveDirectory) Then
            Directory.CreateDirectory(saveDirectory)
        End If

        ' JPEG�G���R�[�_���擾
        Dim jpgEncoder As ImageCodecInfo = GetEncoder(ImageFormat.Jpeg)

        ' �G���R�[�_�p�����[�^�i�i���̎w��j
        Dim encoderParams As New EncoderParameters(1)
        Dim qualityParam As New EncoderParameter(System.Drawing.Imaging.Encoder.Quality, jpgQuality) ' 85%�̕i���ɐݒ�
        encoderParams.Param(0) = qualityParam

        ' �摜��ۑ�
        Select Case colorFormat
            Case "full"
                bmpScreenCapture.Save(savePath, jpgEncoder, encoderParams)
            Case "gray"
                ConvertToGrayscale(bmpScreenCapture).Save(savePath, jpgEncoder, encoderParams)
        End Select

        ' ���\�[�X�����
        g.Dispose()
        bmpScreenCapture.Dispose()

        Console.WriteLine("�X�N���[���L���v�`�����ۑ�����܂���: " & savePath)
    End Sub

    ' �w�肵���C���[�W�t�H�[�}�b�g�̃G���R�[�_���擾����֐�
    Private Function GetEncoder(ByVal format As ImageFormat) As ImageCodecInfo
        Dim codecs As ImageCodecInfo() = ImageCodecInfo.GetImageDecoders()
        For Each codec As ImageCodecInfo In codecs
            If codec.FormatID = format.Guid Then
                Return codec
            End If
        Next
        Return Nothing
    End Function

    ' �J���[�摜���O���[�X�P�[���ɕϊ�����֐�
    Function ConvertToGrayscale(ByVal original As Bitmap) As Bitmap
        Dim grayscale As New Bitmap(original.Width, original.Height)

        For y As Integer = 0 To original.Height - 1
            For x As Integer = 0 To original.Width - 1
                ' �s�N�Z���̐F���擾
                Dim originalColor As Color = original.GetPixel(x, y)

                ' �O���[�X�P�[���̌v�Z�i�P�x�@�j
                Dim gray As Integer = CInt(originalColor.R * 0.3 + originalColor.G * 0.59 + originalColor.B * 0.11)

                ' �O���[�X�P�[���l�ŐV�����F���쐬
                Dim grayColor As Color = Color.FromArgb(gray, gray, gray)

                ' �V�����s�N�Z���ɃZ�b�g
                grayscale.SetPixel(x, y, grayColor)
            Next
        Next

        Return grayscale
    End Function

End Module
