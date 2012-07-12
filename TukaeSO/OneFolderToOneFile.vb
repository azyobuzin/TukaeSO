Imports System.Xml
Imports System.IO

Namespace MyStyle
    ''' <summary>たくさんのファイルを1つにまとめます。
    ''' 実際に使うにはインスタンス化する必要があります。</summary>
    ''' <remarks>pepetaroFolderでは、この後圧縮を行います。</remarks>
    Public Class OneFolderToOneFile
        ''' <summary>処理を開始します。</summary>
        ''' <param name="savePath">保存先のファイル名を指定します。</param>
        ''' <param name="files">格納するファイルを指定します。</param>
        Public Overloads Sub ToOneFile(ByVal savePath As FileName, ByVal files() As FileName)
            Dim writer = XmlWriter.Create(savePath)
            With writer
                .WriteStartDocument()

                'ファイル書き込み
                .WriteStartElement("Files")
                For Each FileName In files
                    If FileName.IsExist Then
                        Dim stream As New FileStream(FileName, FileMode.Open, FileAccess.Read)
                        Dim bs(stream.Length - 1) As Byte
                        stream.Read(bs, 0, stream.Length)
                        .WriteStartElement(FileName.FileName)
                        .WriteBase64(bs, 0, stream.Length)
                        .WriteEndElement()
                        stream.Close()
                        stream.Dispose()
                    End If
                Next
                .WriteEndElement()

                .WriteEndDocument()
                .Close()
            End With
        End Sub

        ''' <summary>フォルダのファイルを取得して処理を行います。</summary>
        ''' <param name="savePath">保存先のファイル名を指定します。</param>
        ''' <param name="folder">まとめたいファイルがあるフォルダを指定します。サブフォルダは検索されません。</param>
        Public Overloads Sub ToOneFile(ByVal savePath As FileName, ByVal folder As DirectoryInfo)
            Dim files = From file In folder.GetFiles Select FullName = New FileName(file.FullName)
            ToOneFile(savePath, files:=files)
        End Sub

        ''' <summary>OneFile化したファイルをフォルダにコピーします。</summary>
        ''' <param name="file">OneFile化されたファイルを指定します。</param>
        ''' <param name="outFolder">出力先を指定します。</param>
        Public Sub FromOneFile(ByVal file As FileName, ByVal outFolder As DirectoryInfo)
            If Not file.IsExist Then
                Throw New FileNotFoundException("ファイルが存在しません。", file.ToString)
            ElseIf Not outFolder.Exists Then
                outFolder.Create()
            End If

            Dim reader = XmlReader.Create(file)
            With reader
                While .Read
                    '判断。Filesはパス
                    If .NodeType = XmlNodeType.Element AndAlso .LocalName <> "Files" Then
                        Dim stream As New FileStream(outFolder.FullName & "\" & _
                                                     .LocalName, FileMode.Create, FileAccess.Write)
                        Dim bs = Convert.FromBase64String(.ReadString)
                        stream.Write(bs, 0, bs.Length)
                        stream.Close()
                        stream.Dispose()
                    End If
                End While
                .Close()
            End With
        End Sub
    End Class
End Namespace