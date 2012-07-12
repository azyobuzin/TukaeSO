Namespace MyStyle
    ''' <summary>例外が発生したことにより発生するイベントの引数です。</summary>
    Public Class ExceptionEventArgs
        Inherits EventArgs
        ''' <param name="innerException">内部で発生した例外を指定します。</param>
        Public Sub New(ByVal innerException As Exception)
            MyBase.New()
            _innerException = innerException
        End Sub
        Public Sub New()
            MyBase.new()
        End Sub

        Dim _innerException As Exception
        ''' <summary>内部で発生した例外です。</summary>
        Public ReadOnly Property InnerException() As Exception
            Get
                Return _innerException
            End Get
        End Property
    End Class
    Public Delegate Sub ExceptionEventHandler(ByVal sender As Object, ByVal e As ExceptionEventArgs)
End Namespace
