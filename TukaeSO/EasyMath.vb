Option Strict On
Imports TukaeSO.MyStyle.IntList
Imports System.Runtime.InteropServices
Imports System.Globalization

Namespace TukaeruModuleAndClass
    ''' <summary>小学校で習うような簡単な数学を集めました。</summary>
    Public Module EasyMath
        ''' <summary>最小公倍数を取得します。</summary>
        ''' <param name="Int1">１つ目の数字を指定します。</param>
        ''' <param name="Int2">２つ目の数字を指定します。</param>
        <DebuggerStepThrough()> Public Function GetLCM(ByVal Int1 As Integer, ByVal Int2 As Integer) As Integer
            Dim Int1_XList As New List(Of Integer)
            Dim Int2_XList As New List(Of Integer)
            Dim X As Integer = 0
            Do
                'Int1を確認
                For i As Integer = 1 To 10
                    Int1_XList.Add(Int1 * (X + i))
                Next
                'Int2を確認
                For i As Integer = 1 To 10
                    Int2_XList.Add(Int2 * (X + i))
                Next
                'Int1_XListとInt2_XListを比べて同じ数字がないかを確認
                Dim Index As Integer
                For Each i As Integer In Int2_XList
                    Index = Int1_XList.IndexOf(i)
                    If Not Index = -1 Then
                        Return Int1_XList(Index)
                    End If
                Next
                '10追加で再ループ
                X += 10
            Loop
        End Function
        ''' <summary>最小公倍数を取得します。</summary>
        ''' <param name="Int1">１つ目の数字を指定します。</param>
        ''' <param name="Int2">２つ目の数字を指定します。</param>
        <DebuggerStepThrough()> Public Function GetLCM(ByVal Int1 As Long, ByVal Int2 As Long) As Long
            Dim Int1_XList As New List(Of Long)
            Dim Int2_XList As New List(Of Long)
            Dim X As Long = 0
            Do
                'Int1を確認
                For i As Integer = 1 To 100
                    Int1_XList.Add(Int1 * (X + i))
                Next
                'Int2を確認
                For i As Integer = 1 To 100
                    Int2_XList.Add(Int2 * (X + i))
                Next
                'Int1_XListとInt2_XListを比べて同じ数字がないかを確認
                Dim Index As Integer
                For Each i As Long In Int2_XList
                    Index = Int1_XList.IndexOf(i)
                    If Not Index = -1 Then
                        Return Int1_XList(Index)
                    End If
                Next
                '100追加で再ループ
                X += 100
            Loop
        End Function

        ''' <summary>最大公約数を取得します。</summary>
        ''' <param name="Int1">１つ目の数字を指定します。</param>
        ''' <param name="Int2">２つ目の数字を指定します。</param>
        <DebuggerStepThrough()> Public Function GetGCD(ByVal Int1 As Integer, ByVal Int2 As Integer) As Integer
            '小さいほうを探す
            Dim Mini As Integer = Math.Min(Int1, Int2)

            Dim Int1_List As New List(Of Integer)
            Dim OKList As New List(Of Integer)
            Dim inting As Double
            'Int1を確認
            For i As Integer = 1 To Mini
                inting = Int1 / i
                If Not IsDecimal(inting) Then
                    Int1_List.Add(i)
                End If
            Next
            'Int2を確認
            For i As Integer = 1 To Mini
                inting = Int2 / i
                If Not IsDecimal(inting) AndAlso Int1_List.Contains(i) Then
                    OKList.Add(i)
                End If
            Next
            Return OKList(OKList.Count - 1)
        End Function
        ''' <summary>最大公約数を取得します。</summary>
        ''' <param name="Int1">１つ目の数字を指定します。</param>
        ''' <param name="Int2">２つ目の数字を指定します。</param>
        <DebuggerStepThrough()> Public Function GetGCD(ByVal int1 As UInteger, ByVal int2 As UInteger) As UInteger
            '小さいほうを探す
            Dim Mini As UInteger = Math.Min(int1, int2)

            Dim Int1_List As New List(Of UInteger)
            Dim OKList As New List(Of UInteger)
            Dim inting As Double
            'Int1を確認
            For i As UInteger = 1 To Mini
                inting = int1 / i
                If Not IsDecimal(inting) Then
                    Int1_List.Add(i)
                End If
            Next
            'Int2を確認
            For i As UInteger = 1 To Mini
                inting = int2 / i
                If Not IsDecimal(inting) AndAlso Int1_List.Contains(i) Then
                    OKList.Add(i)
                End If
            Next
            Return OKList(OKList.Count - 1)
        End Function

        ''' <summary>数字が小数か確認します。</summary>
        ''' <param name="Nomber">確認する数字を指定します。</param>
        ''' <returns>小数ならTrue でないならFalseを返します。</returns>
        <Extension()> Public Function IsDecimal(ByVal Nomber As Double) As Boolean
            Return Nomber.ToString.Contains(".")
        End Function
        ''' <summary>数字が小数か確認します。</summary>
        ''' <param name="Nomber">確認する数字を指定します。</param>
        ''' <returns>小数ならTrue でないならFalseを返します。</returns>
        <Extension()> Public Function IsDecimal(ByVal Nomber As Decimal) As Boolean
            Return Nomber.ToString.Contains(".")
        End Function
        ''' <summary>倍数を取得します。</summary>
        ''' <param name="Int">取得する倍数のx1となる数字を指定します。</param>
        ''' <param name="Max">最大値を指定します。戻り値にはMaxより大きい数は含まれません。</param>
        <DebuggerStepThrough()> Public Function GetMultiple(ByVal Int As Integer, ByVal Max As Integer) As Int32List
            Dim ReIntList As New Int32List
            Dim X As Integer = 0
            Do
                Dim TestInt As Integer
                For i As Integer = 1 To 10
                    TestInt = Int * (X + i)
                    If TestInt > Max Then
                        Return ReIntList
                    Else
                        ReIntList.Add(TestInt)
                    End If
                Next
                X += 10
            Loop
        End Function
        ''' <summary>倍数を取得します。</summary>
        ''' <param name="Int">取得する倍数のx1となる数字を指定します。</param>
        ''' <param name="Max">最大値を指定します。戻り値にはMaxより大きい数は含まれません。</param>
        <DebuggerStepThrough()> Public Function GetMultiple(ByVal Int As Long, ByVal Max As Long) As Int64List
            Dim ReIntList As New Int64List
            Dim X As Long = 0
            Do
                Dim TestInt As Long
                For i As Integer = 1 To 100
                    TestInt = Int * (X + i)
                    If TestInt > Max Then
                        Return ReIntList
                    Else
                        ReIntList.Add(TestInt)
                    End If
                Next
                X += 100L
            Loop
        End Function
        ''' <summary>約数を取得します。</summary>
        ''' <param name="Int">調べる数字を指定します。</param>
        <DebuggerStepThrough()> Public Function GetDiviso(ByVal Int As Integer) As Int32List
            Dim ReIntList As New Int32List
            Dim inting As Double
            For i As Integer = 1 To Int
                inting = Int / i
                If Not IsDecimal(inting) Then
                    ReIntList.Add(CInt(inting))
                End If
            Next
            Return ReIntList
        End Function
        ''' <summary>小数点以下にある数字の文字数を取得します。</summary>
        ''' <param name="dec">小数を指定します。</param>
        ''' <returns>小数点以下にある数字の文字数を返します。整数の場合0を返します。</returns>
        <Extension()> Public Function GetDecimalCount(ByVal dec As Decimal) As Integer
            Dim str = CStr(dec)
            Dim index = str.IndexOf("."c)
            If index = -1 Then
                Return 0
            End If
            str = str.Substring(index + 1)
            Return str.Length
        End Function
        ''' <summary>小数点以下にある数字の文字数を取得します。</summary>
        ''' <param name="dec">小数を指定します。</param>
        ''' <returns>小数点以下にある数字の文字数を返します。整数の場合0を返します。</returns>
        <Extension()> Public Function GetDecimalCount(ByVal dec As Double) As Integer
            Dim str = CStr(dec)
            Dim index = str.IndexOf("."c)
            If index = -1 Then
                Return 0
            End If
            str = str.Substring(index + 1)
            Return str.Length
        End Function

        ''' <summary>分数を表します。</summary>
        <Serializable(), DebuggerDisplay("{Numerator} / {Denominator}")> _
        Public NotInheritable Class Fraction
            Implements IConvertible, IFormattable, ICloneable, IComparable,  _
                    IComparable(Of Fraction), IEquatable(Of Fraction)

            ''' <summary>Fractionの初期値です。
            ''' このオブジェクトは読み取り専用です。</summary>
            <DebuggerBrowsable(DebuggerBrowsableState.Never)> Public Shared ReadOnly Empty As New Fraction(1, 1, True)
            ''' <summary>円周率です。
            ''' このオブジェクトは読み取り専用です。計算に利用してください。</summary>
            '<DebuggerBrowsable(DebuggerBrowsableState.Never)> Public Shared ReadOnly Pi As New Fraction(7, 22, True)

#Region "New"
            ''' <summary>新しい分数を作ります。</summary>
            ''' <param name="denominator">分母を指定します。</param>
            ''' <param name="numerator">分子を指定します。</param>
            Public Sub New(ByVal denominator As UInteger, ByVal numerator As UInteger)
                Me._denominator = denominator
                Me._numerator = numerator
            End Sub
            ''' <summary>新しい分数を作ります。</summary>
            ''' <param name="int">整数を指定します。</param>
            Public Sub New(ByVal int As UInteger)
                Me._denominator = 1
                Me._numerator = int
            End Sub
            ''' <summary>新しい分数を作ります。</summary>
            ''' <param name="dec">小数を指定します。</param>
            Public Sub New(ByVal dec As Double)
                Dim decCount = GetDecimalCount(dec)
                If decCount = 0 Then
                    Me._denominator = 1
                    Me._numerator = CUInt(dec)
                    Return
                End If

                Dim strDenominator = "1"
                For i As Integer = 1 To decCount
                    strDenominator &= "0"
                Next
                Dim intDenominator = CUInt(strDenominator)
                Me._denominator = intDenominator
                Me._numerator = CUInt(dec * intDenominator)
                Reduce()
            End Sub
            ''' <summary>新しい分数 1/1 を作ります。</summary>
            Public Sub New()
                Me._denominator = 1
                Me.Numerator = 1
            End Sub

            ''' <summary>新しい分数を作ります。</summary>
            ''' <param name="denominator">分母を指定します。</param>
            ''' <param name="numerator">分子を指定します。</param>
            ''' <param name="protected">プロパティを変更できるかを指定します。できるならFalse、できないならTrueを指定します。</param>
            Protected Friend Sub New(ByVal denominator As UInteger, ByVal numerator As UInteger, ByVal [protected] As Boolean)
                Me._denominator = denominator
                Me._numerator = numerator

                Me.[ReadOnly] = [protected]
            End Sub
#End Region

            ''' <summary>ProtecetObjectがTrueの時に発生する例外のメッセージです。</summary>
            <DebuggerBrowsable(DebuggerBrowsableState.Never)> Protected Const ProtectedMessage As String = _
            "このオブジェクトは読み取り専用です。使うにはCloneメソッドでコピーしてください。"

            <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private _denominator As UInteger
            ''' <summary>分母</summary>
            Public Property Denominator() As UInteger
                Get
                    Return _denominator
                End Get
                Set(ByVal value As UInteger)
                    If [ReadOnly] Then
                        Throw New ReadOnlyException(ProtectedMessage)
                    End If
                    _denominator = value
                End Set
            End Property

            <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private _numerator As UInteger
            ''' <summary>分子</summary>
            Public Property Numerator() As UInteger
                Get
                    Return _numerator
                End Get
                Set(ByVal value As UInteger)
                    If [ReadOnly] Then
                        Throw New ReadOnlyException(ProtectedMessage)
                    End If
                    _numerator = value
                End Set
            End Property

            ''' <summary>値を変更できるかを表します。</summary>
            Protected Friend ReadOnly [ReadOnly] As Boolean

            ''' <summary>現在の値を利用して新しいオブジェクトを作成します。
            ''' 読み取り専用は解除されます。</summary>
            Public Function Clone() As Object Implements System.ICloneable.Clone
                Return New Fraction(Denominator, Numerator, False)
            End Function

#Region "計算・演算子"
            ''' <summary>約分をします。</summary>
            Public Sub Reduce()
                Dim gcd = GetGCD(Me.Denominator, Me.Numerator)
                Me.Denominator \= gcd
                Me.Numerator \= gcd
            End Sub

            Public Overloads Shared Operator =(ByVal f1 As Fraction, ByVal f2 As Fraction) As Boolean
                Dim lcm = GetLCM(f1.Denominator, f2.Denominator)
                Dim f1N = f1.Numerator * (lcm / f1.Denominator)
                Dim f2N = f2.Numerator * (lcm / f2.Denominator)
                Return f1N = f2N
            End Operator

            Public Overloads Shared Operator =(ByVal f As Fraction, ByVal dec As Double) As Boolean
                Return f.ToDouble = dec
            End Operator

            Public Overloads Shared Operator =(ByVal dec As Double, ByVal f As Fraction) As Boolean
                Return dec = f.ToDouble
            End Operator

            Public Overloads Shared Operator <>(ByVal f1 As Fraction, ByVal f2 As Fraction) As Boolean
                Dim lcm = GetLCM(f1.Denominator, f2.Denominator)
                Dim f1N = f1.Numerator * (lcm / f1.Denominator)
                Dim f2N = f2.Numerator * (lcm / f2.Denominator)
                Return f1N <> f2N
            End Operator

            Public Overloads Shared Operator <>(ByVal f As Fraction, ByVal dec As Double) As Boolean
                Return f.ToDouble <> dec
            End Operator

            Public Overloads Shared Operator <>(ByVal dec As Double, ByVal f As Fraction) As Boolean
                Return dec <> f.ToDouble
            End Operator

            Public Overloads Shared Operator +(ByVal f1 As Fraction, ByVal f2 As Fraction) As Fraction
                Dim lcm = GetLCM(f1.Denominator, f2.Denominator)
                Dim f1N = CUInt(f1.Numerator * (lcm / f1.Denominator))
                Dim f2N = CUInt(f2.Numerator * (lcm / f2.Denominator))
                Return New Fraction(CUInt(lcm), f1N + f2N)
            End Operator

            Public Overloads Shared Operator +(ByVal f As Fraction, ByVal dec As Double) As Fraction
                Return f + New Fraction(dec)
            End Operator

            Public Overloads Shared Operator +(ByVal dec As Double, ByVal f As Fraction) As Fraction
                Return New Fraction(dec) + f
            End Operator

            Public Overloads Shared Operator -(ByVal f1 As Fraction, ByVal f2 As Fraction) As Fraction
                Dim lcm = GetLCM(f1.Denominator, f2.Denominator)
                Dim f1N = CUInt(f1.Numerator * (lcm / f1.Denominator))
                Dim f2N = CUInt(f2.Numerator * (lcm / f2.Denominator))
                Return New Fraction(CUInt(lcm), f1N - f2N)
            End Operator

            Public Overloads Shared Operator -(ByVal f As Fraction, ByVal dec As Double) As Fraction
                Return f - New Fraction(dec)
            End Operator

            Public Overloads Shared Operator -(ByVal dec As Double, ByVal f As Fraction) As Fraction
                Return New Fraction(dec) - f
            End Operator

            Public Overloads Shared Operator *(ByVal f1 As Fraction, ByVal f2 As Fraction) As Fraction
                Dim d = f1.Denominator * f2.Denominator
                Dim n = f1.Numerator * f2.Numerator
                Return New Fraction(d, n)
            End Operator

            Public Overloads Shared Operator *(ByVal f As Fraction, ByVal dec As Double) As Fraction
                Return f * New Fraction(dec)
            End Operator

            Public Overloads Shared Operator *(ByVal dec As Double, ByVal f As Fraction) As Fraction
                Return New Fraction(dec) * f
            End Operator

            Public Overloads Shared Operator /(ByVal dec As Double, ByVal f As Fraction) As Fraction
                Dim ReFraction As New Fraction(f.Numerator, f.Denominator)
                Return dec * ReFraction
            End Operator

            Public Overloads Shared Operator /(ByVal f As Fraction, ByVal dec As Double) As Fraction
                Dim ReFraction As New Fraction(f.Numerator, f.Denominator)
                Return ReFraction * dec
            End Operator

            Public Overloads Shared Operator /(ByVal f1 As Fraction, ByVal f2 As Fraction) As Fraction
                Dim hickF1 As New Fraction(f1.Numerator, f1.Denominator)
                Dim hickF2 As New Fraction(f2.Numerator, f2.Denominator)
                Return hickF1 * hickF2
            End Operator

            Public Overloads Shared Operator <(ByVal f1 As Fraction, ByVal f2 As Fraction) As Boolean
                Dim lcm = GetLCM(f1.Denominator, f2.Denominator)
                Dim f1N = f1.Numerator * (lcm / f1.Denominator)
                Dim f2N = f2.Numerator * (lcm / f2.Denominator)
                Return f1N < f2N
            End Operator

            Public Overloads Shared Operator <(ByVal f As Fraction, ByVal dec As Double) As Boolean
                Return f.ToDouble < dec
            End Operator

            Public Overloads Shared Operator <(ByVal dec As Double, ByVal f As Fraction) As Boolean
                Return dec < f.ToDouble
            End Operator

            Public Overloads Shared Operator >(ByVal f1 As Fraction, ByVal f2 As Fraction) As Boolean
                Dim lcm = GetLCM(f1.Denominator, f2.Denominator)
                Dim f1N = f1.Numerator * (lcm / f1.Denominator)
                Dim f2N = f2.Numerator * (lcm / f2.Denominator)
                Return f1N > f2N
            End Operator

            Public Overloads Shared Operator >(ByVal f As Fraction, ByVal dec As Double) As Boolean
                Return f.ToDouble > dec
            End Operator

            Public Overloads Shared Operator >(ByVal dec As Double, ByVal f As Fraction) As Boolean
                Return dec > f.ToDouble
            End Operator

            Public Overloads Shared Operator <=(ByVal f1 As Fraction, ByVal f2 As Fraction) As Boolean
                Dim lcm = GetLCM(f1.Denominator, f2.Denominator)
                Dim f1N = f1.Numerator * (lcm / f1.Denominator)
                Dim f2N = f2.Numerator * (lcm / f2.Denominator)
                Return f1N <= f2N
            End Operator

            Public Overloads Shared Operator <=(ByVal f As Fraction, ByVal dec As Double) As Boolean
                Return f.ToDouble <= dec
            End Operator

            Public Overloads Shared Operator <=(ByVal dec As Double, ByVal f As Fraction) As Boolean
                Return dec <= f.ToDouble
            End Operator

            Public Overloads Shared Operator >=(ByVal f1 As Fraction, ByVal f2 As Fraction) As Boolean
                Dim lcm = GetLCM(f1.Denominator, f2.Denominator)
                Dim f1N = f1.Numerator * (lcm / f1.Denominator)
                Dim f2N = f2.Numerator * (lcm / f2.Denominator)
                Return f1N >= f2N
            End Operator

            Public Overloads Shared Operator >=(ByVal f As Fraction, ByVal dec As Double) As Boolean
                Return f.ToDouble >= dec
            End Operator

            Public Overloads Shared Operator >=(ByVal dec As Double, ByVal f As Fraction) As Boolean
                Return dec >= f.ToDouble
            End Operator

            ''' <summary>整数割る整数をします。</summary>
            Public Shared Function Divide(ByVal 割られる数 As UInteger, ByVal 割る数 As UInteger) As Fraction
                Dim reFraction As New Fraction(割る数, 割られる数)
                reFraction.Reduce()
                Return reFraction
            End Function
#End Region

            ''' <summary>ToStringを書式設定しないで実行したときの書式です。</summary>
            <DebuggerBrowsable(DebuggerBrowsableState.Never)> Public Shared DefaultFormat As String = _
                                                                    "0" & vbNewLine & "-" & vbNewLine & "0"

#Region "IConvertible"
            'Privateになっているのは、通常使いません。

            Private Function GetTypeCode() As System.TypeCode Implements System.IConvertible.GetTypeCode
                Return TypeCode.Object
            End Function

            Private Function ToBoolean(ByVal provider As System.IFormatProvider) As Boolean Implements System.IConvertible.ToBoolean
                Throw New InvalidCastException("Boolean型への変換はサポートされていません。")
            End Function

            Private Function ToByte(ByVal provider As System.IFormatProvider) As Byte Implements System.IConvertible.ToByte
                Dim answer = Me.Numerator / Me.Denominator
                If IsDecimal(answer) Then
                    Throw New InvalidCastException("小数点以下があるため変換できませんでした。")
                End If
                If answer < Byte.MinValue Then
                    Throw New OverflowException("オーバーフローしました。値が小さすぎます。")
                End If
                Try
                    Return Convert.ToByte(answer, provider)
                Catch ex As OverflowException
                    Throw New OverflowException("オーバーフローしました。値が大きすぎます。", ex)
                End Try
            End Function

            Private Function ToChar(ByVal provider As System.IFormatProvider) As Char Implements System.IConvertible.ToChar
                Throw New InvalidCastException("Char型への変換はサポートされていません。")
            End Function

            Private Function ToDateTime(ByVal provider As System.IFormatProvider) As Date Implements System.IConvertible.ToDateTime
                Throw New InvalidCastException("DateTime型への変換はサポートされていません。")
            End Function

            Private Function ToDecimal(ByVal provider As System.IFormatProvider) As Decimal Implements System.IConvertible.ToDecimal
                Dim answer = Me.Numerator / Me.Denominator
                If answer < Decimal.MinValue Then
                    Throw New OverflowException("オーバーフローしました。値が小さすぎます。")
                End If
                Try
                    Return Convert.ToDecimal(answer, provider)
                Catch ex As OverflowException
                    Throw New OverflowException("オーバーフローしました。値が大きすぎます。", ex)
                End Try
            End Function

            ''' <summary>小数に変換します。</summary>
            Public Overloads Function ToDouble() As Double
                Return Me.Numerator / Me.Denominator
            End Function

            Private Overloads Function ToDouble(ByVal provider As System.IFormatProvider) As Double Implements System.IConvertible.ToDouble
                Return Me.Numerator / Me.Denominator
            End Function

            Private Function ToInt16(ByVal provider As System.IFormatProvider) As Short Implements System.IConvertible.ToInt16
                Dim answer = Me.Numerator / Me.Denominator
                If IsDecimal(answer) Then
                    Throw New InvalidCastException("小数点以下があるため変換できませんでした。")
                End If
                If answer < Short.MinValue Then
                    Throw New OverflowException("オーバーフローしました。値が小さすぎます。")
                End If
                Try
                    Return Convert.ToInt16(answer, provider)
                Catch ex As OverflowException
                    Throw New OverflowException("オーバーフローしました。値が大きすぎます。", ex)
                End Try
            End Function

            Private Overloads Function ToInt32() As Integer
                Dim answer = Me.Numerator / Me.Denominator
                If IsDecimal(answer) Then
                    Throw New InvalidCastException("小数点以下があるため変換できませんでした。")
                End If
                If answer < Integer.MinValue Then
                    Throw New OverflowException("オーバーフローしました。値が小さすぎます。")
                End If
                Try
                    Return Convert.ToInt32(answer)
                Catch ex As OverflowException
                    Throw New OverflowException("オーバーフローしました。値が大きすぎます。", ex)
                End Try
            End Function

            Private Overloads Function ToInt32(ByVal provider As System.IFormatProvider) As Integer Implements System.IConvertible.ToInt32
                Dim answer = Me.Numerator / Me.Denominator
                If IsDecimal(answer) Then
                    Throw New InvalidCastException("小数点以下があるため変換できませんでした。")
                End If
                If answer < Integer.MinValue Then
                    Throw New OverflowException("オーバーフローしました。値が小さすぎます。")
                End If
                Try
                    Return Convert.ToInt32(answer, provider)
                Catch ex As OverflowException
                    Throw New OverflowException("オーバーフローしました。値が大きすぎます。", ex)
                End Try
            End Function

            Private Function ToInt64(ByVal provider As System.IFormatProvider) As Long Implements System.IConvertible.ToInt64
                Dim answer = Me.Numerator / Me.Denominator
                If IsDecimal(answer) Then
                    Throw New InvalidCastException("小数点以下があるため変換できませんでした。")
                End If
                If answer < Long.MinValue Then
                    Throw New OverflowException("オーバーフローしました。値が小さすぎます。")
                End If
                Try
                    Return Convert.ToInt64(answer, provider)
                Catch ex As OverflowException
                    Throw New OverflowException("オーバーフローしました。値が大きすぎます。", ex)
                End Try
            End Function

            Private Function ToSByte(ByVal provider As System.IFormatProvider) As SByte Implements System.IConvertible.ToSByte
                Dim answer = Me.Numerator / Me.Denominator
                If IsDecimal(answer) Then
                    Throw New InvalidCastException("小数点以下があるため変換できませんでした。")
                End If
                If answer < SByte.MinValue Then
                    Throw New OverflowException("オーバーフローしました。値が小さすぎます。")
                End If
                Try
                    Return Convert.ToSByte(answer, provider)
                Catch ex As OverflowException
                    Throw New OverflowException("オーバーフローしました。値が大きすぎます。", ex)
                End Try
            End Function

            Private Function ToSingle(ByVal provider As System.IFormatProvider) As Single Implements System.IConvertible.ToSingle
                Dim answer = Me.Numerator / Me.Denominator
                If answer < Single.MinValue Then
                    Throw New OverflowException("オーバーフローしました。値が小さすぎます。")
                End If
                Try
                    Return Convert.ToSingle(answer, provider)
                Catch ex As OverflowException
                    Throw New OverflowException("オーバーフローしました。値が大きすぎます。", ex)
                End Try
            End Function

            ''' <summary>改行設定を指定して、文字列型に変換します。</summary>
            ''' <param name="newLineType">改行に使う文字を指定します。</param>
            Public Overloads Function ToString(ByVal newLineType As NewLineType) As String
                Dim lines(2) As String
                lines(0) = Me.Numerator.ToString
                For i As Integer = 1 To CStr(Math.Max(Denominator, Numerator)).Length
                    lines(1) &= "-"
                Next
                lines(2) = Me.Denominator.ToString

                Select Case newLineType
                    Case TukaeruModuleAndClass.NewLineType.Cr
                        Return lines(0) & vbCr & lines(1) & vbCr & lines(2)
                    Case TukaeruModuleAndClass.NewLineType.Lf
                        Return lines(0) & vbLf & lines(1) & vbLf & lines(2)
                    Case Else
                        Return lines(0) & vbCrLf & lines(1) & vbCrLf & lines(2)
                End Select
            End Function

            ''' <summary>文字列型に変換します。</summary>
            Public Overrides Function ToString() As String
                Return Me.ToString(DefaultFormat)
            End Function

            Private Overloads Function ToString(ByVal provider As System.IFormatProvider) As String Implements System.IConvertible.ToString
                If TypeOf provider Is NumberFormatInfo Then
                    Return Convert.ToString(Me.ToDouble, provider)
                Else
                    Return Me.ToString
                End If
            End Function

            ''' <summary>書式を指定して文字列型に変換します。</summary>
            ''' <param name="format">書式を指定します。1つ目の"0"は分子に、2つ目の"0"は分母に置き換えられます。
            ''' また"00"など2桁は、"分子 分母"になってしまいます。</param>
            Public Overloads Function ToString(ByVal format As String) As String
                Dim ReStr = format
                Dim ZeroCount = 0
                For i = 0 To ReStr.Length - 1
                    If ReStr(i) = "0"c Then
                        ZeroCount += 1
                        Select Case ZeroCount
                            Case 1
                                ReStr = ReStr.Substring(0, i) & Me.Numerator & ReStr.Substring(i + Me.Numerator.ToString.Length)
                            Case 2
                                ReStr = ReStr.Substring(0, i) & Me.Denominator & ReStr.Substring(i + Me.Denominator.ToString.Length)
                        End Select
                    End If
                Next
                Return ReStr
            End Function

            Private Overloads Function ToString(ByVal format As String, ByVal formatProvider As System.IFormatProvider) As String Implements System.IFormattable.ToString
                Return ToString(format)
            End Function

            Private Function ToType(ByVal conversionType As System.Type, ByVal provider As System.IFormatProvider) As Object Implements System.IConvertible.ToType
                Dim dec As IConvertible = ToDouble()
                Return dec.ToType(conversionType, provider)
            End Function

            Private Function ToUInt16(ByVal provider As System.IFormatProvider) As UShort Implements System.IConvertible.ToUInt16
                Dim answer = Me.Numerator / Me.Denominator
                If IsDecimal(answer) Then
                    Throw New InvalidCastException("小数点以下があるため変換できませんでした。")
                End If
                If answer < UShort.MinValue Then
                    Throw New OverflowException("オーバーフローしました。値が小さすぎます。")
                End If
                Try
                    Return Convert.ToUInt16(answer, provider)
                Catch ex As OverflowException
                    Throw New OverflowException("オーバーフローしました。値が大きすぎます。", ex)
                End Try
            End Function

            Private Overloads Function ToUInt32(ByVal provider As System.IFormatProvider) As UInteger Implements System.IConvertible.ToUInt32
                Dim answer = Me.Numerator / Me.Denominator
                If IsDecimal(answer) Then
                    Throw New InvalidCastException("小数点以下があるため変換できませんでした。")
                End If
                If answer < UInteger.MinValue Then
                    Throw New OverflowException("オーバーフローしました。値が小さすぎます。")
                End If
                Try
                    Return Convert.ToUInt32(answer, provider)
                Catch ex As OverflowException
                    Throw New OverflowException("オーバーフローしました。値が大きすぎます。", ex)
                End Try
            End Function

            ''' <summary>整数に変換します。</summary>
            Public Overloads Function ToUInt32() As UInteger
                Dim answer = Me.Numerator / Me.Denominator
                If IsDecimal(answer) Then
                    Throw New InvalidCastException("小数点以下があるため変換できませんでした。")
                End If
                If answer < UInteger.MinValue Then
                    Throw New OverflowException("オーバーフローしました。値が小さすぎます。")
                End If
                Try
                    Return Convert.ToUInt32(answer)
                Catch ex As OverflowException
                    Throw New OverflowException("オーバーフローしました。値が大きすぎます。", ex)
                End Try
            End Function

            Private Function ToUInt64(ByVal provider As System.IFormatProvider) As ULong Implements System.IConvertible.ToUInt64
                Dim answer = Me.Numerator / Me.Denominator
                If IsDecimal(answer) Then
                    Throw New InvalidCastException("小数点以下があるため変換できませんでした。")
                End If
                If answer < ULong.MinValue Then
                    Throw New OverflowException("オーバーフローしました。値が小さすぎます。")
                End If
                Try
                    Return Convert.ToUInt64(answer, provider)
                Catch ex As OverflowException
                    Throw New OverflowException("オーバーフローしました。値が大きすぎます。", ex)
                End Try
            End Function
#End Region

#Region "Parse"
            ''' <summary>文字列から新しい分数を作ります。</summary>
            ''' <param name="s">分数の素となる文字列を指定します。
            ''' 数字と空白、/と-が許可されています。</param>
            Public Shared Function Parse(ByVal s As String) As Fraction
                If s = vbNullString Then
                    Throw New ArgumentNullException("s")
                End If
                If IsNumeric(s) Then
                    Return New Fraction(CDbl(s))
                End If

                Dim nom1 As String = ""
                Dim nom2 As String = ""
                Using strReader As New IO.StringReader(s)
                    Dim nomCount = 0
                    Do Until strReader.Peek = -1
                        Dim nextChar = ChrW(strReader.Read)
                        Select Case nextChar
                            Case " "c, "/"c, "-"c, Chr(10), Chr(13)
                                '何もしない
                            Case "0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c
                                nomCount += 1
                                Select Case nomCount
                                    Case 1
                                        nom1 &= nextChar
                                        While IsNumeric(ChrW(strReader.Peek))
                                            nom1 &= Chr(strReader.Read)
                                        End While
                                    Case 2
                                        nom2 &= nextChar
                                        While IsNumeric(ChrW(strReader.Peek))
                                            nom2 &= Chr(strReader.Read)
                                        End While
                                    Case Else
                                        Throw New FormatException("数字が多すぎます。")
                                End Select
                            Case Else
                                Throw New FormatException("許可されていない文字が含まれています。")
                        End Select
                    Loop

                    strReader.Close()
                End Using

                Return New Fraction(CUInt(nom2), CUInt(nom1))
            End Function

            ''' <summary>文字列から新しい分数を作ります。</summary>
            ''' <returns>成功したかを返します。</returns>
            Public Shared Function TryParse(ByVal s As String, <Out()> ByRef result As Fraction) As Boolean
                Try
                    result = Parse(s)
                    Return True
                Catch ex As Exception
                    Return False
                End Try
            End Function
#End Region

#Region "CType"
            Public Overloads Shared Widening Operator CType(ByVal f As Fraction) As Double
                Return f.ToDouble
            End Operator

            Public Overloads Shared Narrowing Operator CType(ByVal f As Fraction) As Integer
                Return f.ToInt32
            End Operator

            Public Overloads Shared Narrowing Operator CType(ByVal f As Fraction) As UInteger
                Return f.ToUInt32
            End Operator

            Public Overloads Shared Narrowing Operator CType(ByVal f As Fraction) As UShort
                Return f.ToUInt16(Nothing)
            End Operator

            Public Overloads Shared Narrowing Operator CType(ByVal int As Integer) As Fraction
                If int < UInteger.MinValue Then
                    Throw New OverflowException("値が小さすぎます。")
                End If
                Return New Fraction(int)
            End Operator

            Public Overloads Shared Widening Operator CType(ByVal int As UInteger) As Fraction
                Return New Fraction(int)
            End Operator

            Public Overloads Shared Narrowing Operator CType(ByVal dec As Double) As Fraction
                Try
                    Return New Fraction(dec)
                Catch ex As OverflowException
                    Throw New OverflowException("オーバーフローしました。decが大きすぎます。", ex)
                End Try
            End Operator

            Public Overloads Shared Widening Operator CType(ByVal f As Fraction) As String
                Return f.ToString
            End Operator

            Public Overloads Shared Narrowing Operator CType(ByVal s As String) As Fraction
                Return Parse(s)
            End Operator
#End Region

#Region "CompareTo"
            ''' <summary>現在のオブジェクトを同じ型の別のオブジェクトと比較します。</summary>
            ''' <param name="obj">比較するオブジェクト。</param>
            Public Overloads Function CompareTo(ByVal obj As Object) As Integer Implements System.IComparable.CompareTo
                If Not TypeOf obj Is Fraction Then
                    Throw New ArgumentException("型が違います。", "obj")
                End If
                Return Me.CompareTo(DirectCast(obj, Fraction))
            End Function
            ''' <summary>現在のオブジェクトを同じ型の別のオブジェクトと比較します。</summary>
            ''' <param name="other">比較するオブジェクト。</param>
            Public Overloads Function CompareTo(ByVal other As Fraction) As Integer _
                            Implements System.IComparable(Of Fraction).CompareTo
                Return CInt(other.ToDouble - Me.ToDouble)
            End Function
#End Region

#Region "Equals"
            Public Overloads Function Equals(ByVal other As Fraction) As Boolean Implements System.IEquatable(Of Fraction).Equals
                Return other = Me
            End Function
            Public Overrides Function Equals(ByVal obj As Object) As Boolean
                If Not TypeOf obj Is Fraction Then
                    Throw New ArgumentException("型が違います。", "obj")
                End If

                Return DirectCast(obj, Fraction) = Me
            End Function
#End Region

        End Class
    End Module
End Namespace
