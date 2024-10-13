Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Windows.Forms

Module ScreenCapture

    Sub Main()
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
        Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmm")

        ' �t�@�C���p�X��ݒ�i��: C:\Screenshots\20241013_1130.jpg�j
        Dim savePath As String = Path.Combine("C:\Screenshots", timestamp & ".jpg")

        ' �t�H���_�����݂��Ȃ��ꍇ�͍쐬
        If Not Directory.Exists("C:\Screenshots") Then
            Directory.CreateDirectory("C:\Screenshots")
        End If

        ' �摜��ۑ�
        bmpScreenCapture.Save(savePath, ImageFormat.Jpeg)

        ' ���\�[�X�����
        g.Dispose()
        bmpScreenCapture.Dispose()

        Console.WriteLine("�X�N���[���L���v�`�����ۑ�����܂���: " & savePath)
    End Sub

End Module
