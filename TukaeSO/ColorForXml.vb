Namespace MyStyle
    ''' <summary>Color構造体をXml用にチューニングしました。</summary>
    <Serializable()> Public Structure ColorForXml
        Private _Color As Color

#Region "コンストラクタ"
        Public Sub New(ByVal Color As Color)
            _Color = Color
        End Sub
        Public Sub New(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte)
            _Color = Color.FromArgb(R, G, B)
        End Sub
        Public Sub New(ByVal A As Byte, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte)
            _Color = Color.FromArgb(A, R, G, B)
        End Sub
        Public Sub New(ByVal argb As Integer)
            _Color = Color.FromArgb(argb)
        End Sub
#End Region

#Region "色のARGB"
        ''' <summary>アルファ</summary>
        Public Property A() As Byte
            Get
                Return _Color.A
            End Get
            Set(ByVal value As Byte)
                _Color = Color.FromArgb(value, R, G, B)
            End Set
        End Property
        ''' <summary>赤</summary>
        Public Property R() As Byte
            Get
                Return _Color.R
            End Get
            Set(ByVal value As Byte)
                _Color = Color.FromArgb(A, value, G, B)
            End Set
        End Property
        ''' <summary>緑</summary>
        Public Property G() As Byte
            Get
                Return _Color.G
            End Get
            Set(ByVal value As Byte)
                _Color = Color.FromArgb(A, R, value, B)
            End Set
        End Property
        ''' <summary>青</summary>
        Public Property B() As Byte
            Get
                Return _Color.B
            End Get
            Set(ByVal value As Byte)
                _Color = Color.FromArgb(A, R, G, value)
            End Set
        End Property
#End Region

        ''' <summary>元になっているColor構造体に直して返します。</summary>
        Public Function ToColor() As Color
            Return _Color
        End Function


        Public Shared Widening Operator CType(ByVal color As Color) As ColorForXml
            Return color.ToForXml
        End Operator
        Public Shared Widening Operator CType(ByVal color As ColorForXml) As Color
            Return color.ToColor
        End Operator
    End Structure
End Namespace