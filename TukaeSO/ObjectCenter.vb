Namespace MyStyle
    ''' <summary>破棄可能なオブジェクトの管理を行います。</summary>
    <DebuggerDisplay("Count = {Count}")> _
    Public NotInheritable Class ObjectCenter
        Inherits ReadOnlyCollectionBase
        Implements IDisposable

        ''' <summary>ObjectCenterを初期化します。</summary>
        ''' <param name="objects">管理するオブジェクトを指定します。</param>
        Public Sub New(ByVal ParamArray objects() As ObjectEntry)
            MyBase.New()
            MyBase.InnerList.AddRange(objects)
        End Sub

        ''' <summary>アイテムを取得します。読み取り専用なのでメンバーのみ操作できます。</summary>
        Default ReadOnly Property Item(ByVal key As String) As Object
            Get
                Return Aggregate o As ObjectEntry In InnerList Where o.Key = key Into First()
            End Get
        End Property

        Private disposedValue As Boolean = False        ' 重複する呼び出しを検出するには

        ' IDisposable
        Protected Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: 他の状態を解放します (マネージ オブジェクト)。
                    For Each o As ObjectEntry In InnerList
                        o.Value.Dispose()
                    Next
                End If

                ' TODO: ユーザー独自の状態を解放します (アンマネージ オブジェクト)。
                ' TODO: 大きなフィールドを null に設定します。
                MyBase.InnerList.Clear()
            End If
            Me.disposedValue = True
        End Sub

#Region " IDisposable Support "
        ' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
        ''' <summary>このクラスで管理されているすべてのオブジェクトを破棄します。</summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

        ''' <summary>ObjectCenterの要素となる型です。</summary>
        Public Class ObjectEntry
            ''' <summary>新しくエントリーを作成します。</summary>
            ''' <param name="key">キーを指定します。
            ''' 他の値のキーと重なると最初に見つけた方が返されます。</param>
            ''' <param name="object">登録するオブジェクトを指定します。</param>
            Public Sub New(ByVal key As String, ByVal [object] As IDisposable)
                Me.Key = key
                Me.Value = [object]
            End Sub

            Public Key As String
            Public Value As IDisposable
        End Class

    End Class
End Namespace