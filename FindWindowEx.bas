Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hwndChildToLookAfter As Long, ByVal className As String, ByVal windowName As String) As Long
