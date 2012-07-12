Namespace TukaeruModuleAndClass
    ''' <summary>INIファイルを操作します。</summary>
    Public Class IniFileSupport
        Private Declare Function WritePrivateProfileString Lib "KERNEL32.DLL" Alias "WritePrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
        Private Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
        Dim fileName As FileName
        ''' <param name="fileName">操作するファイルを指定します。</param>
        Public Sub New(ByVal fileName As FileName)
            Me.fileName = fileName
        End Sub
        ''' <summary>値を書き込みます。</summary>
        Public Sub WriteKey(ByVal section As String, ByVal keyName As String, ByVal value As String)
            Call WritePrivateProfileString(section, keyName, value, fileName)
        End Sub
        ''' <summary>値を取得します。</summary>
        Public Function ReadKey(ByVal section As String, ByVal keyName As String, ByVal [default] As String)
            Dim result = Space(255)
            Call GetPrivateProfileString(section, keyName, [default], result, Len(result), fileName)
            Return Left(result, InStr(result, Chr(0)) - 1)
        End Function
    End Class
End Namespace