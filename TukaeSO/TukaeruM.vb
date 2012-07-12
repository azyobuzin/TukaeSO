Namespace TukaeruModuleAndClass
    ''' <summary>ブール型をYesとNoに変えたものです。</summary>
    ''' <remarks>逆にめんどくさくなるかもしれません。</remarks>
    Public Module YesNo
        ''' <summary>ブール型のTrueです。</summary>
        Public Const Yes As Boolean = True
        ''' <summary>ブール型のFalseです。</summary>
        Public Const No As Boolean = False
    End Module
    ''' <summary>何種類かの文字一覧があります。</summary>
    Public NotInheritable Class Chars
        Private Sub New()
        End Sub

        ''' <summary>大文字のアルファベットです。</summary>
        Public Shared ReadOnly Property BigAlphabet() As Char()
            Get
                Dim re(25) As Char
                For i = 0 To 25
                    re(i) = Chr(65 + i)
                Next
                Return re
            End Get
        End Property

        ''' <summary>小文字のアルファベットです。</summary>
        Public Shared ReadOnly Property SmallAlphabet() As Char()
            Get
                Dim re(25) As Char
                For i = 0 To 25
                    re(i) = Chr(97 + i)
                Next
                Return re
            End Get
        End Property

        ''' <summary>0～9までの数字です。</summary>
        Public Shared ReadOnly Property Nomber() As Char()
            Get
                Return New Char() {"0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c}
            End Get
        End Property
    End Class
    ''' <summary>分類するほどでもないメソッドを集めました。</summary>
    Public Module TukaeruModule
        ''' <summary>ファイルやフォルダが存在するかを調べます。</summary>
        ''' <param name="Path">調べるファイル名やフォルダ名を指定します。</param>
        ''' <returns>存在するならTrue、そうでないならFalseを返します。</returns>
        Public Function FileAndFolderExists(ByVal Path As String) As Boolean
            If IO.File.Exists(Path) Then
                Return True
            Else
                Return IO.Directory.Exists(Path)
            End If
        End Function
        ''' <summary>バイト単位をメガバイト単位などにわかりやすく変換します。</summary>
        ''' <param name="FileSize">ファイルの大きさをバイト単位で指定してください。</param>
        ''' <remarks>エクサバイト(EB)まで対応しています。</remarks>
        <Extension()> Public Function FileSizeConvert(ByVal FileSize As Long) As String
            If FileSize < 1024 Then
                Return CStr(FileSize) & " B"
            ElseIf FileSize < 1048576 Then
                Return Format(FileSize \ 1024, "0 KB")
            ElseIf FileSize < 1073741824 Then
                Return Format(FileSize / 1048576, "0.0 MB")
            ElseIf FileSize < 1099511627776 Then
                Return Format(FileSize / 1073741824, "0.0 GB")
            ElseIf FileSize < 1125899906842624 Then
                Return Format(FileSize / 1099511627776, "0.0 TB")
            ElseIf FileSize < 1152921504606846976 Then
                Return Format(FileSize / 1125899906842624, "0.0 PB")
            Else
                Return Format(FileSize / 1152921504606846976, "0.0 EB")
            End If
        End Function
        ''' <summary>バイト単位をメガバイト単位などにわかりやすく変換します。</summary>
        ''' <param name="FileSize">ファイルの大きさをバイト単位で指定してください。</param>
        ''' <remarks>こちらはDecimalを使用しているので重いです。通常はLong版を使ってください。ヨタバイト(YB)まで対応しています。</remarks>
        Public Function FileSizeConvert(ByVal FileSize As Object) As String
            Dim DecFileSize As Decimal
            If FileSize.GetType.Name = "String" Then
                Try
                    DecFileSize = StringToDecimal(DirectCast(FileSize, String))
                Catch ex As ArgumentException
                    DecFileSize = 0D
                Catch ex As OverflowException
                    Throw New ArgumentException("Decimalに変換できませんでした。" & vbNewLine & "引数が大きすぎます。", ex)
                End Try
            Else
                Try
                    DecFileSize = CDec(FileSize)
                Catch ex As InvalidCastException
                    Throw New ArgumentException("Decimalに変換できない型です。" & vbNewLine & "数字を指定してください。", ex)
                End Try
            End If
            If DecFileSize < 1024D Then
                Return CStr(DecFileSize) & " B"
            ElseIf DecFileSize < 1048576 Then
                Return Format(DecFileSize \ 1024, "0 KB")
            ElseIf DecFileSize < 1073741824D Then
                Return Format(DecFileSize / 1048576, "0.0 MB")
            ElseIf DecFileSize < 1099511627776D Then
                Return Format(DecFileSize / 1073741824, "0.0 GB")
            ElseIf DecFileSize < 1125899906842624D Then
                Return Format(DecFileSize / 1099511627776, "0.0 TB")
            ElseIf DecFileSize < 1152921504606846976D Then
                Return Format(DecFileSize / 1125899906842624, "0.0 PB")
            ElseIf DecFileSize < 1180591620717411303424D Then
                Return Format(DecFileSize / 1152921504606846976, "0.0 EB")
            ElseIf DecFileSize < 1208925819614629174706176D Then
                Return Format(DecFileSize / 1180591620717411303424D, "0.0 ZB")
            Else
                Return Format(DecFileSize / 1208925819614629174706176D, "0.0 YB")
            End If
        End Function
        ''' <summary>Val関数と似ていて、数字だけを抽出して十進型で返します。</summary>
        ''' <param name="Str">数字が含まれた文字を指定します。</param>
        <Extension()> Public Function StringToDecimal(ByVal Str As String) As Decimal
            If Len(Str) = 0 Then
                Throw New ArgumentNullException("引数が空です。")
            End If
            Dim NomberList() As Char = {"0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c, "."c}
            Dim NomberStr As String = ""
            If Str(0) = "-" Then
                NomberStr &= "-"
            End If
            Dim dotCount As New UInteger
            For Each c As Char In Str
                If c = NomberList(10) Then
                    dotCount += 1
                End If
                For Each s As Char In NomberList
                    If c = s Then
                        NomberStr &= c.ToString
                        Exit For
                    End If
                Next
            Next
            If NomberStr = "" OrElse NomberStr = "-" Then
                Return 0D
            End If
            Dim ReDec As Decimal
            Try
                ReDec = CDec(NomberStr)
            Catch ex As OverflowException
                Throw New OverflowException("引数の数字が大きすぎます。", ex)
            End Try
            Return ReDec
        End Function
        ''' <summary>Val関数と似ていて、数字だけを抽出して返します。</summary>
        ''' <param name="Str">数字が含まれた文字を指定します。</param>
        <Extension()> Public Function StringToInt32(ByVal Str As String) As Integer
            If Len(Str) = 0 Then
                Throw New ArgumentException("引数が空です。")
            End If
            Dim NomberList() As Char = {"0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c}
            Dim NomberStr As String = ""
            If Str(0) = "-" Then
                NomberStr &= "-"
            End If
            For Each c As Char In Str
                For Each s As Char In NomberList
                    If c = s Then
                        NomberStr &= c.ToString
                        Exit For
                    End If
                Next
            Next
            If NomberStr = "" OrElse NomberStr = "-" Then
                Throw New ArgumentException("数字が１つもありません。")
            End If
            Dim ReInt As Integer
            Try
                ReInt = CInt(NomberStr)
            Catch ex As OverflowException
                Throw New OverflowException("引数の数字が大きすぎます。", ex)
            End Try
            Return ReInt
        End Function
        ''' <summary>Val関数と似ていて、数字だけを抽出して返します。</summary>
        ''' <param name="Str">数字が含まれた文字を指定します。</param>
        <Extension()> Public Function StringToInt64(ByVal Str As String) As Long
            If Len(Str) = 0 Then
                Throw New ArgumentNullException("Str", "引数が空です。")
            End If
            Dim NomberList() As Char = {"0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c}
            Dim NomberStr As String = ""
            If Str(0) = "-" Then
                NomberStr &= "-"
            End If
            For Each c As Char In Str
                For Each s As Char In NomberList
                    If c = s Then
                        NomberStr &= c.ToString
                        Exit For
                    End If
                Next
            Next
            If NomberStr = "" OrElse NomberStr = "-" Then
                Throw New ArgumentException("数字が１つもありません。", Str)
            End If
            Dim ReInt As Long
            Try
                ReInt = CLng(NomberStr)
            Catch ex As OverflowException
                Throw New OverflowException("引数の数字が大きすぎます。", ex)
            End Try
            Return ReInt
        End Function
#Region "ウィンドウ移動"
        Private Declare Function SendMessage Lib "User32.dll" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Long
        Private Declare Sub ReleaseCapture Lib "User32.dll" ()
        Const WM_NCLBUTTONDOWN = &HA1
        Const HTCAPTION = 2
        ''' <summary>ウィンドウをタイトルバーを使わずに移動させる方法を提供します。</summary>
        Public Sub WindowMove(ByVal hWnd As IntPtr)
            Call ReleaseCapture()
            Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        End Sub
#End Region
        ''' <summary>メモ帳などでよく使われるShift-JISを取得します。</summary>
        Public Function GetShift_JISEncoding() As System.Text.Encoding
            Return System.Text.Encoding.GetEncoding("Shift_JIS")
        End Function
        ''' <summary>コレクションをコンソールに書き込みます。</summary>
        ''' <param name="c">対象となるコレクションを指定します。</param>
        ''' <param name="Width">1行に書き込む要素の数を指定します。</param>
        ''' <param name="SleepTime">1行ごとに休む時間を指定します。0以下にすると休みません。</param>
        ''' <param name="SplitString">要素と要素との区切り目の文字を指定します。</param>
        <Extension()> _
        Public Sub CollectionPrint(ByVal c As ICollection, ByVal Width As UInteger, Optional ByVal SleepTime As Integer = 0, Optional ByVal SplitString As String = ", ")
            If c Is Nothing OrElse c.Count = 0 Then
                Throw New ArgumentNullException("c", "引数cが空です。")
            End If
            If Width = 0 Then
                Throw New ArgumentException("Widthは1以上にしてください。", "Width")
            End If

            For i As Integer = 0 To c.Count Step Width
                If i >= c.Count Then
                    Exit For
                End If

                Dim WriteStr As String = ""
                For WidthInt As UInteger = 1 To Width
                    If WidthInt = Width OrElse WidthInt + i >= c.Count - 1 Then
                        WriteStr &= CStr(c(i + WidthInt))
                    Else
                        WriteStr &= CStr(c(i + WidthInt)) & SplitString
                    End If
                Next
                Console.WriteLine(WriteStr)
                If 0 < SleepTime Then
                    Threading.Thread.Sleep(SleepTime)
                End If
            Next
        End Sub
        Dim MainRandomer As Random
        ''' <summary>ランダムな数字を取得します。</summary>
        ''' <param name="MaxValue">最大値を指定します。MinValueより大きい必要があります。</param>
        ''' <param name="MinValue">最小値を指定します。</param>
        Public Function GetRandomNumber(ByVal MinValue As Integer, ByVal MaxValue As Integer) As Integer
            If MainRandomer Is Nothing Then
                MainRandomer = New Random
            End If
            Return MainRandomer.Next(MinValue, MaxValue)
        End Function
        ''' <summary>進行状況をプログレスバーにセットします。</summary>
        ''' <param name="total">トータルを指定します。</param>
        ''' <param name="progress">現在の進行状況を指定します。</param>
        ''' <param name="progressBar">セットされるプログレスバーを指定します。</param>
        Public Sub ProgressSet(ByVal total As Long, ByVal progress As Long, ByVal progressBar As ProgressBar)
            Dim intTotal As Integer
            Dim intProgress As Integer
            Dim 割る数 As UInteger = 1
            Do
                If Integer.TryParse(total / 割る数, intTotal) Then
                    intProgress = progress / 割る数
                    Exit Do
                End If
                割る数 += 1
            Loop
            progressBar.Maximum = intTotal
            progressBar.Value = intProgress
        End Sub
        ''' <summary>進行状況をプログレスバーにセットします。</summary>
        ''' <param name="eventArgs">進行状況が変わったときに発生するイベントの引数「e」をセットします。</param>
        ''' <param name="progressBar">セットされるプログレスバーを指定します。</param>
        <Extension()> Public Sub ProgressSet(ByVal eventArgs As Deployment.Application.DeploymentProgressChangedEventArgs, ByVal progressBar As ProgressBar)
            ProgressSet(eventArgs.BytesTotal, eventArgs.BytesCompleted, progressBar)
        End Sub
        ''' <summary>進行状況をプログレスバーにセットします。</summary>
        ''' <param name="eventArgs">進行状況が変わったときに発生するイベントの引数「e」をセットします。</param>
        ''' <param name="progressBar">セットされるプログレスバーを指定します。</param>
        <Extension()> Public Sub ProgressSet(ByVal eventArgs As Net.UploadProgressChangedEventArgs, ByVal progressBar As ProgressBar)
            ProgressSet(eventArgs.TotalBytesToSend, eventArgs.BytesSent, progressBar)
        End Sub
        ''' <summary>進行状況をプログレスバーにセットします。</summary>
        ''' <param name="eventArgs">進行状況が変わったときに発生するイベントの引数「e」をセットします。</param>
        ''' <param name="progressBar">セットされるプログレスバーを指定します。</param>
        <Extension()> Public Sub ProgressSet(ByVal eventArgs As Net.DownloadProgressChangedEventArgs, ByVal progressBar As ProgressBar)
            ProgressSet(eventArgs.TotalBytesToReceive, eventArgs.BytesReceived, progressBar)
        End Sub
        ''' <summary>簡単なパスワードを作成します。</summary>
        ''' <param name="length">文字数を指定します。</param>
        Public Function CreatePassword(ByVal length As Integer) As String
            Dim sb As New System.Text.StringBuilder(length)
            For i = 1 To length
                Dim chrs As Char()
                Select Case GetRandomNumber(0, 3)
                    Case 0
                        chrs = Chars.BigAlphabet
                    Case 1
                        chrs = Chars.SmallAlphabet
                    Case Else
                        chrs = Chars.Nomber
                End Select
                sb.Append(chrs.Random)
            Next
            Return sb.ToString
        End Function
    End Module


    ''' <summary>設定した時間になったことをお知らせします。</summary>
    <System.ComponentModel.DefaultEvent("ItWasTime"), System.Serializable()> _
    <System.ComponentModel.DefaultProperty("SettingDateTime")> _
    Public Class DateTimer
        Inherits System.ComponentModel.Component
        Protected Timer As New Timers.Timer(50.0R)
        ''' <param name="SetingDateTime">設定時刻を指定します。</param>
        Public Sub New(ByVal SetingDateTime As Date)
            AddHandler Timer.Elapsed, AddressOf TimeCheck
            Me.SettingDateTime = SetingDateTime
            Timer.Enabled = True
        End Sub
        Public Sub New()
            AddHandler Timer.Elapsed, AddressOf TimeCheck
        End Sub
        ''' <summary>設定した時間になった時に発生します。</summary>
        ''' <param name="Enabled">時間になった時にEnabledプロパティを変更できるように参照型でプロパティを送ります。</param>
        <System.ComponentModel.Description("設定した時間になった時に発生します。")> _
        Public Event ItWasTime(ByVal sender As Object, ByVal e As EventArgs, ByRef Enabled As Boolean)
        Dim _settingDateTime As Date
        ''' <summary>設定時刻を指定します。</summary>
        <System.ComponentModel.Category("動作"), System.ComponentModel.Description("設定時刻を指定します。")> _
        Public Property SettingDateTime() As Date
            Get
                Return _settingDateTime
            End Get
            Set(ByVal value As Date)
                _settingDateTime = value
            End Set
        End Property
        Protected Overridable Sub TimeCheck(ByVal sender As Object, ByVal e As Timers.ElapsedEventArgs)
            If Now >= SettingDateTime Then
                Me.Enabled = False
                OnItWasTime(New EventArgs)
            End If
        End Sub
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            RemoveHandler Timer.Elapsed, AddressOf TimeCheck
            MyBase.Dispose(disposing)
        End Sub
        Protected Overridable Sub OnItWasTime(ByVal e As EventArgs)
            RaiseEvent ItWasTime(Me, e, Timer.Enabled)
        End Sub
        Protected Overrides Sub Finalize()
            RemoveHandler Timer.Elapsed, AddressOf TimeCheck
            MyBase.Finalize()
        End Sub
        ''' <summary>確認を行うかを指定します。</summary>
        <System.ComponentModel.Description("確認を行うかを指定します。"), System.ComponentModel.DefaultValue(False)> _
        Public Property Enabled() As Boolean
            Get
                Return Timer.Enabled
            End Get
            Set(ByVal value As Boolean)
                Timer.Enabled = value
            End Set
        End Property
    End Class

    ''' <summary>改行のタイプです。</summary>
    <Flags()> Public Enum NewLineType
        Cr = 13
        Lf = 10
        CrLf = 23
    End Enum

End Namespace