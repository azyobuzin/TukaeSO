Public Class Form1

    Private Sub PictureBox1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        Dim g As Graphics = e.Graphics
        '既定のアプリケーションアイコン(WIN32: IDI_APPLICATION)
        g.DrawIcon(SystemIcons.Application, 0, 0)
        'システムのアスタリスクアイコン(WIN32: IDI_ASTERISK)
        g.DrawIcon(SystemIcons.Asterisk, 40, 0)
        'システムのエラーアイコン(WIN32: IDI_ERROR)
        g.DrawIcon(SystemIcons.Error, 80, 0)
        'システムの感嘆符アイコン(WIN32: IDI_EXCLAMATION)
        g.DrawIcon(SystemIcons.Exclamation, 120, 0)
        'システムの手の形のアイコン(WIN32: IDI_HAND)
        g.DrawIcon(SystemIcons.Hand, 160, 0)
        'システムの情報アイコン(WIN32: IDI_INFORMATION)
        g.DrawIcon(SystemIcons.Information, 200, 0)
        'システムの疑問符アイコン(WIN32: IDI_QUESTION)
        g.DrawIcon(SystemIcons.Question, 240, 0)
        'システムの警告アイコン(WIN32: IDI_WARNING)
        g.DrawIcon(SystemIcons.Warning, 280, 0)
        'Windowsのロゴアイコン(WIN32: IDI_WINLOGO)
        g.DrawIcon(SystemIcons.WinLogo, 320, 0)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
       
    End Sub
End Class