<Assembly: Extension()> 
Namespace MyStyle
    ''' <summary>拡張メソッドの集まりです。</summary>
    <Extension()> Public Module MyExtensions
        ''' <summary>ブールのTrueとFalseを入れ替えます。</summary>
        ''' <param name="b">入れ替えるオブジェクトを指定します。</param>
        ''' <returns>入れ替えられた値を返します。</returns>
        <Extension()> Public Function Change(ByVal b As Boolean) As Boolean
            Select Case b
                Case True
                    Return False
                Case False
                    Return True
                Case Else
                    Throw New ArgumentException("引数が正しくありません。TrueまたはFalseのみ指定できます。")
            End Select
        End Function

        ''' <summary>Color型からXml用にチューニングされたColorForXml型に変換します。</summary>
        ''' <param name="C">元のColor構造体を指定します。</param>
        <Extension()> Public Function ToForXml(ByVal C As Color) As MyStyle.ColorForXml
            Return New MyStyle.ColorForXml(C)
        End Function

        ''' <summary>TimeSpan型を「00分:00秒」のような簡単なString型に変換します。</summary>
        <Extension()> Public Function ToSimpleString(ByVal span As TimeSpan) As String
            Dim Minuts As Integer = span.Minutes
            Dim Hours As Integer = span.Hours
            If Not span.Days = 0 Then
                Hours += span.Days * 24
            End If
            If Not Hours = 0 Then
                Minuts += Hours * 60
            End If
            Return Minuts.ToString("00") & ":" & span.Seconds.ToString("00")
        End Function

        ''' <summary>キーと値のコレクションを値のコレクションに変換します。</summary>
        <Extension()> Public Function ToArrayList(ByVal d As IDictionary) As ArrayList
            Return New ArrayList(d.Values)
        End Function

        ''' <summary>ディレクトリ内のファイルの名前だけを取得します。</summary>
        <Extension()> Public Function GetFileNames(ByVal dirInfo As IO.DirectoryInfo) As String()
            Return From file In dirInfo.GetFiles Select file.Name
        End Function

        ''' <summary>ディレクトリ内のファイルのフルパスを取得します。</summary>
        <Extension()> Public Function GetFileFullNames(ByVal dirInfo As IO.DirectoryInfo) As FileName()
            Return From file In dirInfo.GetFiles Select New FileName(file.FullName)
        End Function

        ''' <summary>ランダムに１つ取り出します。</summary>
        <Extension()> Public Function Random(ByVal source As IEnumerable) As Object
            Return source(GetRandomNumber(0, Aggregate o In source Into Count()))
        End Function

        ''' <summary>ランダムに１つ取り出します。</summary>
        <Extension()> Public Function Random(Of TSource)(ByVal source As IEnumerable(Of TSource)) As TSource
            Return source(GetRandomNumber(0, Aggregate o In source Into Count()))
        End Function
    End Module
End Namespace