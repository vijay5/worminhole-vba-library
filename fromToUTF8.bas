Private Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Public Function ToUTF8(ByVal sText As String) As String
    Dim nRet As Long, strRet As String

    strRet = String(Len(sText) * 2, vbNullChar)
    nRet = WideCharToMultiByte(65001, &H0, StrPtr(sText), Len(sText), StrPtr(strRet), Len(sText) * 2, 0&, 0&)
    
    ToUTF8 = Left(StrConv(strRet, vbUnicode), nRet)
End Function

Public Function FromUTF8(ByVal sText As String) As String
    Dim nRet As Long, strRet As String
    
    strRet = String(Len(sText), vbNullChar)
    nRet = MultiByteToWideChar(65001, &H0, sText, Len(sText), StrPtr(strRet), Len(strRet))
    
    FromUTF8 = Left(strRet, nRet)
End Function