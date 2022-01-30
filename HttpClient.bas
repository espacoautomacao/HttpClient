Attribute VB_Name = "HttpClient"
'Essa biblioteca só funciona se o módulo JsonConverter estiver junto
'É necessario importar as bibliotecas a baixo
'Microsoft Scripting Runtime
'Microsoft WinHttp Services
'Para interagir entre os objetos use For Each item In Items caso a API retorne uma lista
'Para um retorno simples de objeto apenas use resultado("nomePropriedade")

Function GetFromAPI(url As String) As Variant
    Dim hReq As Object
    Dim json As Dictionary
    Dim i As Long
    Dim j As Long
    Dim propriedade As Variant
    Dim lista As Variant
    Dim response As String
    Dim arr As Variant
    Dim contar As Long
    Dim nivelArr As Long
    Dim itemCol As New Collection
    
    If InStr(1, url, "?", 1) <> 0 Then
        url = url & "&cb=" & Timer() * 100
    Else
        url = url & "?cb=" & Timer() * 100
    End If
	
	On Error GoTo errNull
    
    Set hReq = CreateObject("MSXML2.XMLHTTP")
        With hReq
            .Open "GET", url, False
            .Send
        End With
    
    response = "{""data"":" & hReq.ResponseText & "}"
    
    If hReq.Status <> 200 Then
        Call Err.Raise(Number:=vbObjectError + 513, Description:="Algo deu Errado, Erro: " & hReq.Status & " " & hReq.StatusText)
        Exit Function
    End If
    
    If response = "{""data"":}" Then
        itemCol.Add response
        Set GetFromAPI = itemCol
        Exit Function
    End If
    
    Set json = JsonConverter.ParseJson(response)
    
    Set GetFromAPI = json("data")
	Exit Function
    
errNull:
    Set GetFromAPI = Nothing
    Exit Function
End Function

Function PostAPI(url As String, jsonString As String) As Variant
  Dim objHTTP As New WinHttpRequest
  Dim jsonResposta As String
  Dim jsonItems As New Collection
  Dim jsonDictionary As New Scripting.Dictionary
  Dim dictionaryItems As Scripting.Dictionary
  Dim resultado As Variant
  
  objHTTP.Open "POST", url, False
  objHTTP.SetRequestHeader "Content-Type", "application/json"
  
  If jsonString = "" Then
    Exit Function
  End If
  objHTTP.SetTimeouts 100000, 100000, 100000, 100000
  objHTTP.Send jsonString
  
  If objHTTP.Status <> 200 Then
    MsgBox "Ocorreu um erro " & objHTTP.Status & " - " & objHTTP.StatusText
    Exit Function
  End If
  
  jsonResposta = objHTTP.ResponseText
  jsonResposta = "{""data"":" & jsonResposta & "}"
  
  On Error Resume Next
    Set dictionaryItems = JsonConverter.ParseJson(jsonResposta)
  On Error GoTo 0
  
  If dictionaryItems Is Nothing Or IsNull(dictionaryItems) Then
    PostAPI = Replace(Replace(Replace(jsonResposta, """data"":", ""), "{", ""), "}", "")
    Exit Function
  End If
  
  Set PostAPI = dictionaryItems("data")
  
End Function