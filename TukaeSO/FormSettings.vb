Namespace MyStyle
    ''' <summary>ウィンドウの設定の保管場所です。</summary>
    <System.Serializable()> Public Structure FormSettings

#Region "コンストラクタ"
        Public Sub New(ByVal Form As Form)
            Me.Location = Form.Location
            Me.Size = Form.Size
            Me.WindowState = Form.WindowState
        End Sub
        Public Sub New(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal WindowState As FormWindowState)
            Me.Location = New Point(X, Y)
            Me.Size = New Size(Width, Height)
            Me.WindowState = WindowState
        End Sub
        Public Sub New(ByVal Location As Point, ByVal Size As Size, ByVal WindowState As FormWindowState)
            Me.Location = Location
            Me.Size = Size
            Me.WindowState = WindowState
        End Sub
#End Region

        Private _location As Point
        ''' <summary>位置を表します。</summary>
        Public Property Location() As Point
            Get
                Return _location
            End Get
            Set(ByVal value As Point)
                _location = value
            End Set
        End Property

        Private _size As Size
        ''' <summary>大きさを表します。</summary>
        Public Property Size() As Size
            Get
                Return _size
            End Get
            Set(ByVal value As Size)
                _size = value
            End Set
        End Property

        Private _windowState As FormWindowState
        ''' <summary>状態を表します。</summary>
        Public Property WindowState() As FormWindowState
            Get
                Return _windowState
            End Get
            Set(ByVal value As FormWindowState)
                _windowState = value
            End Set
        End Property



        ''' <summary>ウィンドウにこの設定を適用します。</summary>
        ''' <param name="Form">適用先のウィンドウを指定します。</param>
        Public Sub Restore(ByVal form As Form)
            If form Is Nothing Then
                Throw New ArgumentNullException("form")
            End If
            form.Location = Me.Location
            form.Size = Me.Size
            form.WindowState = Me.WindowState
        End Sub

        ''' <summary>ファイルに保存します。</summary>
        ''' <param name="XmlFileName">保存先を指定します。</param>
        Public Sub Save(ByVal XmlFileName As FileName)
            Dim xs As New Xml.Serialization.XmlSerializer(GetType(FormSettings))
            Using sw As New IO.StreamWriter(XmlFileName.FileName)
                xs.Serialize(sw, Me)
                sw.Close()
            End Using
        End Sub

        ''' <summary>Saveメソッドを使って作成したXmlファイルからインスタンスを作成します。</summary>
        Public Shared Function FromFile(ByVal XmlFileName As FileName) As FormSettings
            If Not XmlFileName.IsExist Then
                Throw New IO.FileNotFoundException("ファイルが存在しません", XmlFileName)
            End If
            Dim xs As New Xml.Serialization.XmlSerializer(GetType(FormSettings))
            Using sr As New IO.StreamReader(XmlFileName.FileName)
                FromFile = xs.Deserialize(sr)
                sr.Close()
            End Using
        End Function
    End Structure
End Namespace