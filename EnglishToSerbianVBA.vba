Dim WordList As Object

Sub TranslateAndAnalyze()
    Dim sourceWord As String
    sourceWord = InputBox("Enter the word to translate and analyze:", "Translate and Analyze")
    
    ' Google Translate API
    Dim googleApiKey As String
    googleApiKey = "YOUR_GOOGLE_API_KEY"
    
    Dim translateUrl As String
    translateUrl = "https://translation.googleapis.com/language/translate/v2?key=" & googleApiKey & "&q=" & sourceWord & "&source=en&target=sr"
    
    Dim xmlHttpTranslate As Object
    Set xmlHttpTranslate = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlHttpTranslate.Open "GET", translateUrl, False
    xmlHttpTranslate.send
    
    Dim translatedText As String
    translatedText = Split(Split(xmlHttpTranslate.responseText, """translatedText"":""")(1), """")(0)
    
    ' WordsAPI
    Dim wordsApiKey As String
    wordsApiKey = "YOUR_WORDSAPI_KEY"
    
    Dim wordsApiUrl As String
    wordsApiUrl = "https://wordsapiv1.p.rapidapi.com/words/" & sourceWord
    
    Dim xmlHttpWordsAPI As Object
    Set xmlHttpWordsAPI = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlHttpWordsAPI.Open "GET", wordsApiUrl, False
    xmlHttpWordsAPI.setRequestHeader "X-RapidAPI-Key", wordsApiKey
    xmlHttpWordsAPI.send
    
    Dim wordDefinition As String
    wordDefinition = Split(Split(xmlHttpWordsAPI.responseText, """definition"":""")(1), """")(0)
    
    ' Output results to Excel
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Translation and Analysis"
    
    ws.Cells(1, 1).Value = "Source Word"
    ws.Cells(1, 2).Value = "Translated Word"
    ws.Cells(1, 3).Value = "Word Definition"
    
    ws.Cells(2, 1).Value = sourceWord
    ws.Cells(2, 2).Value = translatedText
    ws.Cells(2, 3).Value = wordDefinition
    
    ' Update word list
    If WordList Is Nothing Then
        Set WordList = CreateObject("Scripting.Dictionary")
    End If
    If WordList.Exists(sourceWord) Then
        WordList(sourceWord) = WordList(sourceWord) + 1
    Else
        WordList(sourceWord) = 1
    End If
    
    ' Output results to Excel
    ' ... (ostatak koda ostaje isti kao u prethodnom odgovoru)
    
    ' Update word list worksheet
    UpdateWordListSheet
End Sub

Sub UpdateWordListSheet()
    Dim wsWordList As Worksheet
    On Error Resume Next
    Set wsWordList = ThisWorkbook.Sheets("WordList")
    On Error GoTo 0
    
    If wsWordList Is Nothing Then
        Set wsWordList = ThisWorkbook.Sheets.Add
        wsWordList.Name = "WordList"
        wsWordList.Cells(1, 1).Value = "Word"
        wsWordList.Cells(1, 2).Value = "Count"
    End If
    
    wsWordList.Cells.ClearContents
    
    wsWordList.Cells(1, 1).Value = "Word"
    wsWordList.Cells(1, 2).Value = "Count"
    
    Dim i As Long
    Dim keys As Variant
    keys = WordList.keys
    For i = 0 To WordList.Count - 1
        wsWordList.Cells(i + 2, 1).Value = keys(i)
        wsWordList.Cells(i + 2, 2).Value = WordList(keys(i))
    Next i
    
    ' Sort word list
    wsWordList.Sort.SortFields.Clear
    wsWordList.Sort.SortFields.Add Key:=Range("B2:B" & WordList.Count + 1), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    wsWordList.Sort.SetRange Range("A1:B" & WordList.Count + 1)
    wsWordList.Sort.Header = xlYes
    wsWordList.Sort.MatchCase = False
    wsWordList.Sort.Orientation = xlTopToBottom
    wsWordList.Sort.SortMethod = xlPinYin
    wsWordList.Sort.Apply
End Sub
