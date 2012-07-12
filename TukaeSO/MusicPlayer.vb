Imports System.ComponentModel
Namespace MyStyle
    Public Class MusicPlayer
        Inherits System.ComponentModel.Component

        Public Const [Alias] As String = "TukaeSOMusicPlayer"

        ''' <summary>AutoStopがTrueで、自動的に停止された時に発生します。</summary>
        <Description("AutoStopがTrueで、自動的に停止された時に発生します。")> _
        Public Event AutoStopped As EventHandler
        Protected Sub OnAutoStopped(ByVal e As EventArgs)
            RaiseEvent AutoStopped(Me, e)
        End Sub

        Public Sub New()
            MyBase.New()
        End Sub
        Public Sub New(ByVal FileName As String)
            MyBase.New()
            Me.FileName = FileName
        End Sub

        Dim _FileName As String
        ''' <summary>再生するファイルを指定します。</summary>
        <Category("ファイル"), Description("再生するファイルを指定します。")> _
        Public Property FileName() As String
            Get
                Return _FileName
            End Get
            Set(ByVal value As String)
                If Not value = vbNullString Then
                    If Not IO.File.Exists(value) Then
                        Throw New ArgumentException("FileNameが無効です。ファイルが存在しません。")
                    End If
                End If
                _FileName = value
                Try
                    [Stop]()
                Catch ex As PlayStateNotRightException
                    '何もしない
                End Try
            End Set
        End Property

        Dim _Repeat As Boolean
        ''' <summary>繰り返し再生するかを指定します。</summary>
        <Category("動作"), Description("繰り返し再生するかを指定します。"), DefaultValue(False)> _
        Public Property Repeat() As Boolean
            Get
                Return _Repeat
            End Get
            Set(ByVal value As Boolean)
                _Repeat = value
            End Set
        End Property

        Protected WithEvents AutoStopTimer As New Timers.Timer(500)
        Dim _AutoStop As Boolean = True
        ''' <summary>再生が終わった時に自動的にStopメソッドを実行するかを指定します。</summary>
        <Category("動作"), Description("再生が終わった時に自動的にStopメソッドを実行するかを指定します。")> _
        <DefaultValue(True)> Public Overridable Property AutoStop() As Boolean
            Get
                Return _AutoStop
            End Get
            Set(ByVal value As Boolean)
                _AutoStop = value
                If _AutoStop = False Then
                    AutoStopTimer.Enabled = False
                End If
                If PlayState = EPlayState.Playing Then
                    If _AutoStop Then
                        AutoStopTimer.Enabled = True
                    End If
                End If
            End Set
        End Property

        ''' <summary>再生を開始します。</summary>
        Public Overridable Overloads Sub Play()
            If FileName = vbNullString Then
                Throw New NullReferenceException("FileNameが空です。" & vbNewLine & "再生するファイルを指定してください。")
            End If
            Dim Result As Integer
            Result = SendString("open """ & FileName & """ alias " & [Alias])
            If Result <> 0 Then
                Dim ErrorMessage As String = GetErrorString(Result)
                Throw New MciException(ErrorMessage, Result)
            End If
            SendString("setaudio " & [Alias] & " volume to " & Volume)
            Result = Nothing
            If Repeat Then
                Result = SendString("play " & [Alias] & " repeat")
                If Result <> 0 Then
                    Dim ErrorMessage As String = GetErrorString(Result)
                    Throw New MciException(ErrorMessage, Result)
                End If
            Else
                Result = SendString("play " & [Alias])
                If Result <> 0 Then
                    Dim ErrorMessage As String = GetErrorString(Result)
                    Throw New MciException(ErrorMessage, Result)
                End If
            End If
            If AutoStop Then
                AutoStopTimer.Enabled = True
            End If
        End Sub

        ''' <summary>指定したファイルの再生を開始します。</summary>
        ''' <param name="FileName">再生するファイルを指定します。</param>
        ''' <param name="EndToWait">再生が終わるまで待つかを指定します。
        ''' 待つならTrue そうでないならFalseを入力します。</param>
        Public Overloads Shared Sub Play(ByVal FileName As FileName, ByVal EndToWait As Boolean)
            Dim Player As New MusicPlayer
            Player.AutoStop = False
            Try
                Player.FileName = FileName
            Catch ex As IO.FileNotFoundException
                Throw ex
            End Try
            Try
                Player.Play()
            Catch ex As MciException
                Throw ex
            End Try
            If EndToWait Then
                Do Until Player.PlayState = EPlayState.Stopped
                    Application.DoEvents()
                Loop
            End If
        End Sub

        ''' <summary>停止させます。</summary>
        ''' <remarks>再生中でないとエラーになります。</remarks>
        Public Overridable Sub [Stop]()
            If PlayState = EPlayState.NoOpen Then
                Throw New PlayStateNotRightException("MCIで開かれていません。" & vbNewLine & "Playを行ってから実行してください。", EPlayState.NoOpen)
            End If
            AutoStopTimer.Enabled = False
            Dim Result As Integer = SendString("close " & [Alias])
            If Result <> 0 Then
                Dim ErrorMessage As String = GetErrorString(Result)
                Throw New MciException(ErrorMessage, Result)
            End If
        End Sub

        Protected Friend Declare Ansi Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer
        ''' <summary>MCIにコマンドを送ります。</summary>
        ''' <param name="cmdString">コマンドを指定します。</param>
        Protected Friend Function SendString(ByVal cmdString As String) As Integer
            SendString = mciSendString(cmdString, "", 0, 0)
        End Function

        Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal fdwError As Integer, ByVal lpszErrorText As String, ByVal cchErrorText As Integer) As Integer
        ''' <summary>mciSendString関数の戻り値からエラーメッセージを取得します。</summary>
        ''' <param name="APIResult">mciSendString関数の戻り値を指定します。。</param>
        ''' <returns>APIResultに対応するエラーメッセージを返します。</returns>
        Protected Friend Function GetErrorString(ByVal APIResult As Integer) As String
            Dim Buffer As String = New String(Chr(0), 255)
            Dim ErrorMessage As String
            Call mciGetErrorString(APIResult, Buffer, Len(Buffer))
            ErrorMessage = Replace(Replace(Buffer, Chr(0), ""), " ", "")
            Return ErrorMessage
        End Function

        ''' <summary>再生中の曲の長さを取得します。</summary>
        ''' <remarks>再生中でないとエラーになります。</remarks>
        <Browsable(False)> Public Overridable ReadOnly Property Length() As Long
            Get
                If PlayState = EPlayState.NoOpen Then
                    Throw New PlayStateNotRightException("MCIで開かれていません。" & vbNewLine & "Playを行ってから実行してください。", EPlayState.NoOpen)
                End If
                Dim Buffer As String = New String(Chr(0), 255)
                Dim _Length As Long
                Dim Result As Integer
                Result = mciSendString("status " & [Alias] & " length", Buffer, Len(Buffer), 0)
                _Length = TukaeruModuleAndClass.StringToInt64(Buffer)
                If Result <> 0 Then
                    Dim ErrorMessage As String = GetErrorString(Result)
                    Throw New MciException(ErrorMessage, Result)
                End If
                Return _Length
            End Get
        End Property

        ''' <summary>一時停止します。</summary>
        ''' <remarks>再生中でないとエラーになります。</remarks>
        Public Sub Pause()
            Select Case PlayState
                Case EPlayState.NoOpen
                    Throw New PlayStateNotRightException("MCIで開かれていません。" & vbNewLine & "Playを行ってから実行してください。", EPlayState.NoOpen)
                Case EPlayState.Paused
                    Throw New PlayStateNotRightException("すでに一時停止中です。", EPlayState.Paused)
                Case EPlayState.Stopped
                    Throw New PlayStateNotRightException("停止中です。" & vbNewLine & "Playを行ってから実行してください。", EPlayState.Stopped)
            End Select
            Dim Result As Integer = SendString("pause " & [Alias])
            If Result <> 0 Then
                Dim ErrorMessage As String = GetErrorString(Result)
                Throw New MciException(ErrorMessage, Result)
            End If
        End Sub

        ''' <summary>一時停止から復帰します。</summary>
        ''' <remarks>一時停止中でないとエラーになります。</remarks>
        Public Sub [Resume]()
            Select Case PlayState
                Case EPlayState.NoOpen
                    Throw New PlayStateNotRightException("MCIで開かれていません。", EPlayState.NoOpen)
                Case EPlayState.Stopped
                    Throw New PlayStateNotRightException("停止中です。", EPlayState.Stopped)
                Case EPlayState.Playing
                    Throw New PlayStateNotRightException("すでに再生中です。", EPlayState.Playing)
            End Select
            Dim Result As Integer = SendString("resume " & [Alias])
            If Result <> 0 Then
                Dim ErrorMessage As String = GetErrorString(Result)
                Throw New MciException(ErrorMessage, Result)
            End If
        End Sub

        ''' <summary>再生中なら一時停止、一時停止中なら再開します。</summary>
        ''' <remarks>再生中でないとエラーになります。</remarks>
        Public Sub PauseOrResume()
            Select Case PlayState
                Case EPlayState.Playing
                    Pause()
                Case EPlayState.Paused
                    [Resume]()
                Case Else
                    Throw New PlayStateNotRightException("PauseOrResumeは、再生中か一時停止中でないと利用できません。", PlayState)
            End Select
        End Sub

        ''' <summary>再生状況を取得します。</summary>
        <Browsable(False)> Public ReadOnly Property PlayState() As EPlayState
            Get
                Dim Buffer As String = New String(Chr(0), 255)
                Dim Mode As String
                Dim Result As Integer
                Result = mciSendString("status " & [Alias] & " mode", Buffer, Len(Buffer), 0)
                Mode = Replace(Buffer, Chr(0), "")
                If Result = 263 Then
                    Return EPlayState.NoOpen
                ElseIf Result <> 0 Then
                    Dim ErrorMessage As String = GetErrorString(Result)
                    Throw New MciException(ErrorMessage, Result)
                End If
                Select Case Mode
                    Case "playing"
                        Return EPlayState.Playing
                    Case "stopped"
                        Return EPlayState.Stopped
                    Case "paused"
                        Return EPlayState.Paused
                End Select
            End Get
        End Property

        Dim _Volume As Integer = 500
        ''' <summary>音量を指定します。</summary>
        <Category("動作"), Description("音量を指定します。"), DefaultValue(500S)> _
        Public Property Volume() As Integer
            Get
                If Not PlayState = EPlayState.NoOpen Then
                    Dim Buffer As String = New String(Chr(0), 255)
                    Call mciSendString("status " & [Alias] & " volume", Buffer, Len(Buffer), 0)
                    Try
                        _Volume = CInt(Replace(Buffer, Chr(0), ""))
                    Catch ex As Exception
                        _Volume = 500
                    End Try
                End If
                Return _Volume
            End Get
            Set(ByVal value As Integer)
                If value > 1000 OrElse value < 0 Then
                    Throw New ArgumentException("セットする数字が不正です。" & vbNewLine & "音量は0以上 1000以下である必要があります。")
                End If
                _Volume = value
                If Not PlayState = EPlayState.NoOpen Then
                    Dim Result As Integer = SendString("setaudio " & [Alias] & " volume to " & CStr(_Volume))
                    If Result <> 0 Then
                        Throw New MciException(GetErrorString(Result), Result)
                    End If
                End If
            End Set
        End Property

        ''' <summary>再生中の位置を取得します。</summary>
        ''' <returns>現在地をミリ秒単位で返します。</returns>
        ''' <remarks>再生中でないとエラーになります。</remarks>
        ''' <value>移動したい位置をミリ秒単位で指定します。</value>
        <Browsable(False)> Public Property Location() As Long
            Get
                If PlayState = EPlayState.NoOpen Then
                    Throw New PlayStateNotRightException("再生位置を取得するには、MCIで開かれている必要があります。", EPlayState.NoOpen)
                End If
                Dim Buffer As String = New String(Chr(0), 255)
                Dim Result As Integer = mciSendString("status " & [Alias] & " position", Buffer, Len(Buffer), 0)
                If Result <> 0 Then
                    Dim ErrorMessage As String = GetErrorString(Result)
                    Throw New MciException(ErrorMessage, Result)
                End If
                Return TukaeruModuleAndClass.StringToInt64(Buffer)
            End Get
            Set(ByVal value As Long)
                If PlayState = EPlayState.NoOpen Then
                    Throw New PlayStateNotRightException("再生位置をセットするには、MCIで開かれている必要があります。", EPlayState.NoOpen)
                End If
                If value < 0 OrElse value > Length Then
                    Throw New ArgumentException("セットする数字が不正です。" & vbNewLine & "位置は0以上 長さ以下である必要があります。")
                End If
                Dim Result As Integer = SendString("seek " & [Alias] & " to " & CStr(value))
                If Result <> 0 Then
                    Dim ErrorMessage As String = GetErrorString(Result)
                    Throw New MciException(ErrorMessage, Result)
                End If
                SendString("play " & [Alias])
            End Set
        End Property

        Private Sub AutoStopTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles AutoStopTimer.Elapsed
            If PlayState = EPlayState.Stopped OrElse PlayState = EPlayState.NoOpen Then
                AutoStopTimer.Enabled = False
                [Stop]()
                OnAutoStopped(New EventArgs)
            End If
        End Sub

        Protected Overrides Sub Finalize()
            On Error Resume Next
            Me.Stop()
            MyBase.Finalize()
        End Sub

        ''' <summary>MCIで発生した時の例外クラスです。</summary>
        Public Class MciException
            Inherits ApplicationException
            Public Sub New(ByVal Message As String, ByVal ErrorNomber As Integer)
                MyBase.New(Message)
                en = ErrorNomber
            End Sub
            Dim en As Integer
            ''' <summary>MCIから返された数字です。</summary>
            Public ReadOnly Property ErrorNomber() As Integer
                Get
                    Return en
                End Get
            End Property
        End Class
        ''' <summary>PlayStateが実行する作業に適していない時にスローされる例外クラスです。</summary>
        Public Class PlayStateNotRightException
            Inherits ApplicationException
            Public Sub New(ByVal Message As String, ByVal NowPlayState As EPlayState)
                MyBase.New(Message)
                _nowPlayState = NowPlayState
            End Sub
            Dim _nowPlayState As EPlayState
            Public ReadOnly Property NowPlayState() As EPlayState
                Get
                    Return _nowPlayState
                End Get
            End Property
        End Class
        ''' <summary>再生状況を表します。</summary>
        Public Enum EPlayState
            Playing = 1
            Stopped = -1
            Paused = 0
            NoOpen = -2
        End Enum
    End Class
End Namespace

