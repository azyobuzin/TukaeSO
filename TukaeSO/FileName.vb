Imports System.IO
Namespace MyStyle
    ''' <summary>ファイル名を管理する型です。</summary>
    ''' <remarks>FileInfoクラスと似ていますが、String型との変換が簡単になっています。</remarks>
    <System.Serializable()> Public Structure FileName
        Private _FullPath As String
        ''' <param name="FullName">ファイルのフルパスを指定します。</param>
        Public Sub New(ByVal FullName As String)
            Me.FullName = FullName
        End Sub
        ''' <summary>ファイルのフルパスを指定します。</summary>
        Public Property FullName() As String
            Get
                Return _FullPath
            End Get
            Set(ByVal value As String)
                If Len(value) = 0 Then
                    _FullPath = vbNullString
                End If
                Select Case IsFileName(value)
                    Case PathType.File
                        _FullPath = value
                    Case PathType.FileURL
                        _FullPath = value.Substring(9)
                    Case Else
                        Throw New ArgumentException("パスが正しくありません。")
                End Select
            End Set
        End Property
        ''' <summary>ファイル名だけを取得します。</summary>
        Public ReadOnly Property FileName() As String
            Get
                If FullName = vbNullString Then
                    Throw New NullReferenceException("フルパスが空です。")
                End If
                Return Path.GetFileName(FullName)
            End Get
        End Property
        ''' <summary>フォルダ名を取得します。</summary>
        Public ReadOnly Property DirectoryName() As String
            Get
                If FullName = vbNullString Then
                    Throw New NullReferenceException("フルパスが空です。")
                End If
                Return Path.GetDirectoryName(FullName)
            End Get
        End Property
        ''' <summary>ファイル名を取得します。</summary>
        Public Overrides Function ToString() As String
            Return FullName
        End Function
        ''' <summary>ファイルを読み込むためのStreamReaderを取得します。</summary>
        Public Function GetReader() As StreamReader
            Try
                If Not IsExist() Then
                    Throw New FileNotFoundException("ファイルが存在しません。")
                End If
            Catch ex As NullReferenceException
                Throw ex
            End Try
            Return New StreamReader(FullName)
        End Function
        ''' <summary>ファイルに書き込むためのStreamWriterを取得します。</summary>
        Public Function GetWriter() As StreamWriter
            Try
                If Not IsExist() Then
                    Throw New FileNotFoundException("ファイルが存在しません。")
                End If
            Catch ex As NullReferenceException
                Throw ex
            End Try
            Return New StreamWriter(FullName)
        End Function
        ''' <summary>ファイルが存在するかを確認します。</summary>
        Public Function IsExist() As Boolean
            If FullName = vbNullString Then
                Throw New NullReferenceException("フルパスが空です。")
            End If
            Return File.Exists(FullName)
        End Function
        ''' <summary>ファイルを作成します。</summary>
        Public Function Create() As FileStream
            Try
                Return File.Create(FullName)
            Catch ex As Exception
                Throw ex
            End Try
        End Function
        ''' <summary>ファイルを削除します。</summary>
        Public Sub Delete()
            Try
                If Not IsExist() Then
                    Throw New FileNotFoundException("ファイルが存在しません。")
                End If
                File.Delete(FullName)
            Catch ex As Exception
                Throw ex
            End Try
        End Sub
        ''' <summary>拡張子を取得します。</summary>
        Public ReadOnly Property Extension() As String
            Get
                Return Path.GetExtension(FullName)
            End Get
        End Property
        ''' <summary>FileInfo型に変換します。</summary>
        Public Function ToFileInfo() As FileInfo
            Return New FileInfo(FullName)
        End Function

        Public Overloads Shared Narrowing Operator CType(ByVal Str As String) As FileName
            Return New FileName(Str)
        End Operator
        Public Overloads Shared Widening Operator CType(ByVal FileName As FileName) As String
            Return FileName.ToString
        End Operator
        Public Overloads Shared Widening Operator CType(ByVal fileName As FileName) As FileInfo
            Return fileName.ToFileInfo
        End Operator
        Public Overloads Shared Widening Operator CType(ByVal info As FileInfo) As FileName
            Return New FileName(info.FullName)
        End Operator

        ''' <summary>指定された文字列がパスかどうかを調べます。</summary>
        Public Shared Function IsFileName(ByVal path As String) As PathType
            If Len(path) = 0 Then
                Throw New ArgumentNullException(path)
            ElseIf Directory.Exists(path) Then
                Return PathType.Directory
            ElseIf path.Substring(0, 8).ToLower = "file:///" Then
                Return PathType.FileURL
            ElseIf File.Exists(path) Then
                Return PathType.File
            ElseIf path(path.Length - 1) = "\"c Then
                Return PathType.Directory
            ElseIf path.Substring(0, 3) Like "?:\" OrElse path.Substring(0, 2) Like "\\" Then
                Return FileOrFolder(path)
            End If
            Return PathType.NotFileName
        End Function
        Private Shared Function FileOrFolder(ByVal path As String) As PathType
            Dim 最後の円マークから後 = path.Substring(path.LastIndexOf("\"c))
            If path.Contains("."c) Then
                Return PathType.File
            End If
            Return PathType.Directory
        End Function

        Public Enum PathType
            ''' <summary>パスがファイルであることを表します。</summary>
            File = 1
            ''' <summary>パスがフォルダであることを表します。</summary>
            Directory
            ''' <summary>パスが「file:///」から始まるファイルのURLであることを表します。</summary>
            FileURL
            ''' <summary>ファイルを表すパスではないことを表します。</summary>
            NotFileName = 0
        End Enum
    End Structure
End Namespace