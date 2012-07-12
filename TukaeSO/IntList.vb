Option Strict On
Namespace MyStyle
    Namespace IntList
        ''' <summary>整数専用のコレクションです。</summary>
        <System.Serializable()> Public Class Int32List
            Inherits List(Of Integer)
            Public Sub New()
                MyBase.New()
            End Sub
            Public Sub New(ByVal Collection As IEnumerable(Of Integer))
                MyBase.New(Collection)
            End Sub
            ''' <summary>登録されている数の合計を取得します。</summary>
            Public ReadOnly Property Total() As Integer
                Get
                    If MyBase.Count = 0 Then
                        Return 0
                    End If
                    Dim ReTotal As Integer
                    For Each i As Integer In Me
                        ReTotal += i
                    Next
                    Return ReTotal
                End Get
            End Property
            ''' <summary>登録されている数の平均を取得します。</summary>
            ''' <returns>登録されている数の平均を返します。何もない場合はnullを返します。</returns>
            Public ReadOnly Property Average() As Decimal
                Get
                    If MyBase.Count = 0 Then
                        Return Nothing
                    End If
                    Return CDec(Total) / CDec(MyBase.Count)
                End Get
            End Property
            ''' <summary>Int64Listに変換します。</summary>
            Public Function ToInt64List() As Int64List
                Dim List64 As New Int64List
                For Each i As Integer In Me
                    List64.Add(CLng(i))
                Next
                Return List64
            End Function
            ''' <summary>DecimalListに変換します。</summary>
            Public Function ToDecimalList() As DecimalList
                Dim DList As New DecimalList
                For Each i As Integer In Me
                    DList.Add(CDec(i))
                Next
                Return DList
            End Function


            Public Shared Widening Operator CType(ByVal IntList As Int32List) As Int64List
                Return IntList.ToInt64List()
            End Operator
            Public Shared Widening Operator CType(ByVal IntList As Int32List) As DecimalList
                Return IntList.ToDecimalList
            End Operator
        End Class

        ''' <summary>長整数型の整数専用コレクションです。</summary>
        <System.Serializable()> Public Class Int64List
            Inherits List(Of Long)
            Public Sub New()
                MyBase.New()
            End Sub
            Public Sub New(ByVal Collection As IEnumerable(Of Long))
                MyBase.New(Collection)
            End Sub
            ''' <summary>登録されている数の合計を取得します。</summary>
            Public ReadOnly Property Total() As Long
                Get
                    If MyBase.Count = 0 Then
                        Return 0L
                    End If
                    Dim ReTotal As Long
                    For Each i As Long In Me
                        ReTotal += i
                    Next
                    Return ReTotal
                End Get
            End Property
            ''' <summary>登録されている数の平均を取得します。</summary>
            ''' <returns>登録されている数の平均を返します。何もない場合はnullを返します。</returns>
            Public ReadOnly Property Average() As Decimal
                Get
                    If MyBase.Count = 0 Then
                        Return Nothing
                    End If
                    Return CDec(Total) / CDec(MyBase.Count)
                End Get
            End Property
            ''' <summary>DecimalListに変換します。</summary>
            Public Function ToDecimalList() As DecimalList
                Dim DList As New DecimalList
                For Each i As Long In Me
                    DList.Add(CDec(i))
                Next
                Return DList
            End Function


            Public Shared Narrowing Operator CType(ByVal IntList As Int64List) As Int32List
                Dim List32 As New Int32List
                For Each i As Long In IntList
                    Dim AddI As Integer
                    If Integer.TryParse(CStr(i), AddI) Then
                        List32.Add(AddI)
                    Else
                        Throw New OverflowException("Integer型に入りきらない数字がありました。")
                    End If
                Next
                Return List32
            End Operator
            Public Shared Widening Operator CType(ByVal IntList As Int64List) As DecimalList
                Return IntList.ToDecimalList
            End Operator
        End Class

        ''' <summary>十進型の小数または整数専用のコレクションです。</summary>
        <System.Serializable()> Public Class DecimalList
            Inherits List(Of Decimal)
            Public Sub New()
                MyBase.New()
            End Sub
            Public Sub New(ByVal Collection As IEnumerable(Of Decimal))
                MyBase.New(Collection)
            End Sub
            ''' <summary>登録されている数の合計を取得します。</summary>
            Public ReadOnly Property Total() As Decimal
                Get
                    If MyBase.Count = 0 Then
                        Return 0D
                    End If
                    Dim ReTotal As Decimal
                    For Each i As Decimal In Me
                        ReTotal += i
                    Next
                    Return ReTotal
                End Get
            End Property
            ''' <summary>登録されている数の平均を取得します。</summary>
            ''' <returns>登録されている数の平均を返します。何もない場合はnullを返します。</returns>
            Public ReadOnly Property Average() As Decimal
                Get
                    If MyBase.Count = 0 Then
                        Return Nothing
                    End If
                    Return Total / CDec(MyBase.Count)
                End Get
            End Property


            Public Shared Narrowing Operator CType(ByVal IntList As DecimalList) As Int32List
                Dim List32 As New Int32List
                For Each i As Decimal In IntList
                    Dim AddI As Integer
                    If Integer.TryParse(Format(i, "0"), AddI) Then
                        List32.Add(AddI)
                    Else
                        Throw New OverflowException("Integer型に入りきらない数字がありました。")
                    End If
                Next
                Return List32
            End Operator
            Public Shared Narrowing Operator CType(ByVal IntList As DecimalList) As Int64List
                Dim List64 As New Int64List
                For Each i As Decimal In IntList
                    Dim AddI As Long
                    If Long.TryParse(Format(i, "0"), AddI) Then
                        List64.Add(AddI)
                    Else
                        Throw New OverflowException("Long型に入りきらない数字がありました。")
                    End If
                Next
                Return List64
            End Operator
        End Class
    End Namespace
End Namespace
