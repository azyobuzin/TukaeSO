Namespace TukaeruModuleAndClass
    ''' <summary>特殊ディレクトリを取得します。</summary>
    Public MustInherit Class SpecialFolders
        ''' <summary>WshShellを使用してオブジェクト特殊ディレクトリを取得します。</summary>
        ''' <param name="name">取得する特定フォルダの名前です。詳しくは、http://msdn.microsoft.com/ja-jp/library/cc364490.aspx
        ''' </param>
        Public Shared Function GetFolderPath(ByVal name As String) As String
            Dim o As Object = CreateObject("WScript.Shell")
            GetFolderPath = CType(o.SpecialFolders(name), String)
            Runtime.InteropServices.Marshal.ReleaseComObject(o)
        End Function
        ''' <summary>ログインしているユーザーの特殊ディレクトリを取得します。</summary>
        Public MustInherit Class UserDirectory
            '''<summary>
            ''' アプリケーションフォルダを取得します。
            ''' （注）ProgramFilesフォルダではありません。
            ''' </summary>
            Public Shared ReadOnly Property ApplicationData() As String
                Get
                    Dim AppDataFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
                    Return AppDataFolder
                End Get
            End Property
            ''' <summary>クッキーフォルダを取得します。</summary>
            Public Shared ReadOnly Property Cookies() As String
                Get
                    Dim Cookie As String = Environment.GetFolderPath(Environment.SpecialFolder.Cookies)
                    Return Cookie
                End Get
            End Property
            '''<summary>デスクトップフォルダを取得します。</summary>
            Public Shared ReadOnly Property Desktop() As String
                Get
                    Dim DesktopFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    Return DesktopFolder
                End Get
            End Property
            ''' <summary>お気に入りフォルダを取得します。</summary>
            Public Shared ReadOnly Property Favorites() As String
                Get
                    Dim okiniiri As String = Environment.GetFolderPath(Environment.SpecialFolder.Favorites)
                    Return okiniiri
                End Get
            End Property
            '''<summary>履歴フォルダを取得します。</summary>
            Public Shared ReadOnly Property History() As String
                Get
                    Dim rireki As String = Environment.GetFolderPath(Environment.SpecialFolder.History)
                    Return rireki
                End Get
            End Property
            '''<summary>インターネット一時フォルダを取得します。</summary>
            Public Shared ReadOnly Property InternetCache() As String
                Get
                    Dim InternetTemp As String = Environment.GetFolderPath(Environment.SpecialFolder.InternetCache)
                    Return InternetTemp
                End Get
            End Property
            '''<summary>
            ''' アプリケーションフォルダを取得します。
            ''' （注）ProgramFilesフォルダではありません。
            ''' </summary>
            Public Shared ReadOnly Property LocalApplicationData() As String
                Get
                    Dim AppDataFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                    Return AppDataFolder
                End Get
            End Property
            '''<summary>マイドキュメントフォルダを取得します。</summary>
            Public Shared ReadOnly Property MyDocuments() As String
                Get
                    Dim Documents As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    Return Documents
                End Get
            End Property
            '''<summary>マイミュージックフォルダを取得します。</summary>
            Public Shared ReadOnly Property MyMusic() As String
                Get
                    Dim Music As String = Environment.GetFolderPath(Environment.SpecialFolder.MyMusic)
                    Return Music
                End Get
            End Property
            '''<summary>マイピクチャーフォルダを取得します。</summary>
            Public Shared ReadOnly Property MyPictures() As String
                Get
                    Dim PicFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
                    Return PicFolder
                End Get
            End Property
            '''<summary>Personalフォルダを取得します。</summary>
            Public Shared ReadOnly Property Personal() As String
                Get
                    Dim Person As String = Environment.GetFolderPath(Environment.SpecialFolder.Personal)
                    Return Person
                End Get
            End Property
            '''<summary>スタートメニューのプログラムフォルダを取得します。</summary>
            Public Shared ReadOnly Property Programs() As String
                Get
                    Dim Program As String = Environment.GetFolderPath(Environment.SpecialFolder.Programs)
                    Return Program
                End Get
            End Property
            '''<summary>最近使用したファイルを格納しているフォルダを取得します。</summary>
            Public Shared ReadOnly Property Recent() As String
                Get
                    Dim rireki As String = Environment.GetFolderPath(Environment.SpecialFolder.Recent)
                    Return rireki
                End Get
            End Property
            '''<summary>「送る」フォルダを取得します。</summary>
            Public Shared ReadOnly Property SendTo() As String
                Get
                    Dim ST As String = Environment.GetFolderPath(Environment.SpecialFolder.SendTo)
                    Return ST
                End Get
            End Property
            '''<summary>スタートメニューフォルダを取得します。</summary>
            Public Shared ReadOnly Property StartMenu() As String
                Get
                    Dim Start As String = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu)
                    Return Start
                End Get
            End Property
            '''<summary>スタートアップフォルダを取得します。</summary>
            Public Shared ReadOnly Property Startup() As String
                Get
                    Dim StartupFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.Startup)
                    Return StartupFolder
                End Get
            End Property
            '''<summary>テンプレートフォルダを取得します。</summary>
            Public Shared ReadOnly Property Templates() As String
                Get
                    Dim Temp As String = Environment.GetFolderPath(Environment.SpecialFolder.Templates)
                    Return Temp
                End Get
            End Property
            '''<summary>一時フォルダを取得します。</summary>
            Public Shared ReadOnly Property TempFolder() As String
                Get
                    Dim TempFileName As String = IO.Path.GetTempPath
                    Return TempFileName.Substring(0, TempFileName.Length - 1)
                End Get
            End Property
        End Class
        ''' <summary>すべてのユーザーの特殊ディレクトリを取得します。</summary>
        Public MustInherit Class AllUsers
            ''' <summary>共通アプリケーションデータフォルダを取得します。</summary>
            Public Shared ReadOnly Property CommonApplicationData() As String
                Get
                    Dim CommonAppData As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)
                    Return CommonAppData
                End Get
            End Property
            '''<summary>
            ''' CommonProgramFilesフォルダを取得します。
            ''' （例）C:\ProgramFiles\Common Files
            ''' </summary>
            Public Shared ReadOnly Property CommonProgramFiles() As String
                Get
                    Dim CommonPF As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles)
                    Return CommonPF
                End Get
            End Property
            '''<summary>ProgramFilesフォルダを取得します。</summary>
            Public Shared ReadOnly Property ProgramFiles() As String
                Get
                    Dim pf As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
                    Return pf
                End Get
            End Property
            '''<summary>
            ''' システムフォルダを取得します。
            '''（例）C:\Windows\System32 
            '''</summary>
            Public Shared ReadOnly Property System() As String
                Get
                    Dim SystemDir As String = Environment.GetFolderPath(Environment.SpecialFolder.System)
                    Return SystemDir
                End Get
            End Property
            '''<summary>Windowsフォルダを取得します。</summary>
            Public Shared ReadOnly Property WindowsFolder() As String
                Get
                    Dim WindowsDir As String = Environment.GetEnvironmentVariable("WINDIR")
                    Return WindowsDir
                End Get
            End Property
            ''' <summary>すべてのユーザーのデスクトップを取得します。</summary>
            Public Shared ReadOnly Property Desktop() As String
                Get
                    Return GetFolderPath("AllUsersDesktop")
                End Get
            End Property
            ''' <summary>すべてのユーザーのスタートメニューフォルダを取得します。</summary>
            Public Shared ReadOnly Property StartMenu() As String
                Get
                    Return GetFolderPath("AllUsersStartMenu")
                End Get
            End Property
        End Class
    End Class

    Public MustInherit Class FileAndFolderDialog
        '''<summary>フォルダ選択ダイアログを表示します。</summary>
        Public Shared Function FolderDialog(Optional ByVal 説明 As String = "フォルダを選択してください", Optional ByVal NewFolder As Boolean = True) As String
            Try
                Dim 返す内容 As String
                Dim FolderBrowserDialog1 As New FolderBrowserDialog
                FolderBrowserDialog1.Description = 説明
                FolderBrowserDialog1.ShowNewFolderButton = NewFolder
                If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
                    返す内容 = FolderBrowserDialog1.SelectedPath
                Else
                    返す内容 = "Cancel"
                End If
                Return 返す内容
            Catch ex As System.Exception
                Throw New ApplicationException(ex.Message)
            End Try
        End Function
        ''' <summary>ファイル選択ダイアログを表示します。</summary>
        ''' <param name="BoxTitle">ダイアログのタイトルです。</param>
        ''' <param name="Filter">
        ''' 開くことのできるファイルの拡張子を指定してください。
        ''' （書き方例）"MP3|*.mp3"
        '''</param>
        ''' <returns>指定されたファイル名を返します。</returns>
        Public Shared Function FileOpenDialog(Optional ByVal BoxTitle As String = "開く", Optional ByVal Filter As String = "すべてのファイル|*.*") As String
            Try
                Dim Dialog As New OpenFileDialog
                Dialog.Title = BoxTitle
                Try
                    Dialog.Filter = Filter
                Catch ex As System.Exception
                    Throw New ArgumentException("Filterが正しくありません。")
                End Try
                If Dialog.ShowDialog = DialogResult.OK Then
                    Return Dialog.FileName
                Else
                    Return "Cancel"
                End If
            Catch ex As System.Exception
                Throw New ApplicationException(ex.Message)
            End Try
        End Function
        ''' <summary>複数選択できるファイル選択ダイアログを表示します。</summary>
        ''' <param name="BoxTitle">ダイアログのタイトルです。</param>
        ''' <param name="Filter">
        ''' 開くことのできるファイルの拡張子を指定してください。
        ''' （書き方例）"MP3|*.mp3"
        '''</param>
        ''' <returns>指定されたファイル名を配列で返します。</returns>
        Public Shared Function FilesOpenDialog(Optional ByVal BoxTitle As String = "開く", Optional ByVal Filter As String = "すべてのファイル|*.*") As String()
            Try
                Dim Dialog As New OpenFileDialog
                Dialog.Title = BoxTitle
                Try
                    Dialog.Filter = Filter
                Catch ex As System.Exception
                    Throw New ArgumentException("Filterが正しくありません。")
                End Try
                If Dialog.ShowDialog = DialogResult.OK Then
                    Return Dialog.FileNames
                Else
                    Return New String() {"Cancel"}
                End If
            Catch ex As System.Exception
                Throw New ApplicationException(ex.Message)
            End Try
        End Function

        ''' <summary>保存ダイアログを表示します。</summary>
        ''' <param name="BoxTitle">ダイアログのタイトルです。</param>
        ''' <param name="Filter">
        ''' 保存することのできるファイルの拡張子を指定してください。
        ''' （書き方例）"MP3|*.mp3"
        ''' </param>
        ''' <returns>指定されたファイル名を返します。</returns>
        Public Shared Function FileSaveDialog(Optional ByVal BoxTitle As String = "保存", Optional ByVal Filter As String = "テキストファイル|*.txt|すべてのファイル|*.*") As String
            Try
                Dim Dialog As New SaveFileDialog
                Dialog.Title = BoxTitle
                Try
                    Dialog.Filter = Filter
                Catch ex As System.Exception
                    Throw New ArgumentException("Filterが正しくありません。")
                End Try
                If Dialog.ShowDialog = DialogResult.OK Then
                    Return Dialog.FileName
                Else
                    Return "Cancel"
                End If
            Catch ex As System.Exception
                Throw New ApplicationException(ex.Message)
            End Try
        End Function
    End Class
End Namespace
