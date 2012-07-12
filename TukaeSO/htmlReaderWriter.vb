Imports System.Text
Imports System.IO
Namespace MyStyle
    ''' <summary>Htmlを読み込みます。</summary>
    Public Class HtmlReader
        Inherits TextReader
        Dim reader As StreamReader
        Dim fileName As FileName
        Dim encoding As Encoding

#Region "New"
        ''' <summary>新しいHtmlReaderを作成します。
        ''' エンコードはShift-JISになります。</summary>
        ''' <param name="fileName">読み込むファイルを指定します。</param>
        Public Sub New(ByVal fileName As FileName)
            If Not fileName.IsExist Then
                Throw New FileNotFoundException("ファイルが存在しません。", fileName.FullName)
            End If

            Me.fileName = fileName
            Me.encoding = GetShift_JISEncoding()

            reader = New StreamReader(fileName, Me.encoding)
        End Sub

        ''' <summary>新しいHtmlReaderを作成します。</summary>
        ''' <param name="fileName">読み込むファイルを指定します。</param>
        ''' <param name="encoding">エンコードを指定します。</param>
        Public Sub New(ByVal fileName As FileName, ByVal encoding As Encoding)
            If Not fileName.IsExist Then
                Throw New FileNotFoundException("ファイルが存在しません。", fileName.FullName)
            End If

            Me.fileName = fileName
            Me.encoding = encoding

            reader = New StreamReader(fileName, encoding)
        End Sub
#End Region

#Region "Read"
        ''' <summary>次の文字を取得します。カーソルは進みません。</summary>
        ''' <returns>次の文字を表す整数を返します。
        ''' 使用できない場合は-1を返します。</returns>
        Public Overrides Function Peek() As Integer
            Return reader.Peek()
        End Function

        ''' <summary>次の文字を取得し、カーソルを進めます。</summary>
        ''' <returns>次の文字を表す整数を返します。
        ''' 使用できない場合は-1を返します。</returns>
        Public Overrides Function Read() As Integer
            Return reader.Read()
        End Function

        ''' <returns>読み取った文字数を返します。</returns>
        ''' <summary>現在のストリームの index で示された位置から、count で指定された最大文字数を buffer に読み込みます。</summary>
        Public Overrides Function ReadBlock(<Runtime.InteropServices.Out()> ByVal buffer() As Char, _
                                            ByVal index As Integer, ByVal count As Integer) As Integer
            Return reader.ReadBlock(buffer, index, count)
        End Function

        ''' <summary>現在のストリームの index で示された位置から、count で指定された最大文字数を buffer に読み込みます。</summary>
        ''' <returns>読み取った文字数を返します。</returns>
        Public Overrides Function Read(<Runtime.InteropServices.Out()> ByVal buffer() As Char, _
                                       ByVal index As Integer, ByVal count As Integer) As Integer
            Return reader.Read(buffer, index, count)
        End Function

        ''' <summary>1行分を読み取ります。</summary>
        Public Overrides Function ReadLine() As String
            Return reader.ReadLine()
        End Function

        ''' <summary>現在の位置から最後まで読み取ります。</summary>
        Public Overrides Function ReadToEnd() As String
            Return reader.ReadToEnd()
        End Function

        ''' <summary>HtmlのBody部分を読み取ります。</summary>
        Public Function ReadBody() As String
            Dim str As String
            Using bodyReader As New StreamReader(fileName, encoding)
                str = bodyReader.ReadToEnd
                bodyReader.Close()
            End Using

            Dim startPosion = str.ToLower.IndexOf("<body")
            Dim endPosion = str.ToLower.IndexOf("</body")
            Do
                endPosion += 1
                If str(endPosion) = ">"c Then
                    Exit Do
                End If
            Loop
            endPosion += 1

            Return str.Substring(startPosion, endPosion - startPosion)
        End Function

        ''' <summary>タイトルを読み取ります。</summary>
        Public Function ReadTitle() As String
            Dim lower As String
            Dim str As String
            Using bodyReader As New StreamReader(fileName, encoding)
                str = bodyReader.ReadToEnd
                bodyReader.Close()
            End Using
            lower = str.ToLower

            Dim start = lower.IndexOf("<title>") + 7
            Dim endPos = lower.IndexOf("</title>", start)
            Return str.Substring(start, endPos)
        End Function
#End Region

        ''' <summary>リーダーを終了します。</summary>
        Public Overrides Sub Close()
            reader.Close()
            MyBase.Close()
        End Sub

    End Class
End Namespace