''' возвращает код языка интерфейса Excel
Function excelUILanguage() As Long
    excelUILanguage = Application.LanguageSettings.LanguageID(2) ' 2 = msoLanguageIDUI
End Function
