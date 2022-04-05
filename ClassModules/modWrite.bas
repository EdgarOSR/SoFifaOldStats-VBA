Attribute VB_Name = "modForm"
'Option Explicit
'
'Private Sub ClearOldData()
'
'      ' limpa as tabelas de cada planilha do arquivo
'      shJson.Range("A1").CurrentRegion.Delete
'      shJson.Range("A1").Select
'
'End Sub
'
'Private Function ParseJson(ByVal playerInfos As Variant, ByVal playerSkills As Variant, ByVal playerTraits As Variant) As String
'
'      ' inicializa o objeto dicionario
'      Dim dict As New Dictionary
'      Dim k As Variant
'      Dim json As String, lastKey As String
'      Dim i As Integer
'
'      ' inicia criacao do json
'      Let json = "{"
'
'      For i = 1 To 3
'
'            Select Case i
'                  Case 1
'                        Set dict = playerInfos
'                  Case 2
'                        Set dict = playerSkills
'                  Case 3
'                        Set dict = playerTraits
'            End Select
'
'            If (i = 2) Then
'                  json = json & """" & "skills" & """" & ":{"
'            ElseIf (i = 3) Then
'                  json = json & """" & "player_traits" & """" & ":["
'            End If
'
'            For Each k In dict.Keys
'
'                  If (i = 3) Then
'                        json = json & """" & dict(k) & """"
'                  ElseIf (k = "positions") Then
'                        json = json & """" & "positions" & """" & ": [" & dict(k) & "]"
'                  ElseIf (IsNumeric(dict(k))) Then
'                        json = json & """" & k & """" & ":" & dict(k)
'                  Else
'                        json = json & """" & k & """" & ":" & """" & dict(k) & """"
'                  End If
'
'                  If ((k = dict.Keys(dict.Count - 1)) And (i = 2)) Then
'                        json = json & "},"
'                  Else
'                        json = json & ","
'                  End If
'            Next k
'      Next i
'
'      json = json & "]}"
'      json = Replace(json, ",]}", "]}")
'
'      ParseJson = json
'
'End Function
'
'Public Sub WriteJson(ByVal playerInfos As Variant, ByVal playerSkills As Variant, ByVal playerTraits As Variant)
'
''Selenium VBA
''https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0
'
''ChromeDriver
''https://chromedriver.storage.googleapis.com/index.html?path=96.0.4664.45/
'
'Dim ieApp As New ChromeDriver
'Dim ieField As IHTMLEmbedElement
'Dim ieButton As IHTMLButtonElement
'
'Dim json As String
'Let json = ParseJson(playerInfos, playerSkills, playerTraits)
'
'shJson.Range("A9").vNewValue = "Pesquisando ID... Aguarde a mensagem"
'
''ieApp.AddArgument ("--headless")
'
'ieApp.Get "https://www.pes6.es/stats/fifa-to-pes6"
'
'ieApp.Wait (1800)
'ieApp.FindElementById("showManualStats").Click
'ieApp.FindElementById("offlineStats").Clear
'ieApp.FindElementById("offlineStats").SendKeys (json)
'ieApp.FindElementById("convertOffline").Click
'ieApp.Wait (1200)
'
'If ieApp.FindElementById("copy").IsDisplayed Then
'      ieApp.FindElementById("copy").Click
'      MsgBox "Cole o valor no editor", vbInformation
'End If
'
'shJson.Range("A9").vNewValue = ""
'
'ieApp.Quit
'
'End Sub
'
''Public Sub WriteJson(ByVal playerInfos As Variant, ByVal playerSkills As Variant, ByVal playerTraits As Variant)
''
''      ' limpa os valores antigos
''      Call ClearOldData
''
''      ' converte os dados obtidos do site SoFifa em Json
''      shJson.Range("A1").Select
''      ActiveCell.vNewValue = ParseJson(playerInfos, playerSkills, playerTraits)
''
''      MsgBox "Finalizado", vbInformation
''
''End Sub
'
'Public Sub OpenWebChrome()
'
''Selenium VBA
''https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0
'
''ChromeDriver
''https://chromedriver.storage.googleapis.com/index.html?path=100.0.4896.60/
'
'Dim ieApp As New ChromeDriver
'Dim ieField As IHTMLEmbedElement
'Dim ieButton As IHTMLButtonElement
'
'ieApp.Get "https://www.pes6.es/stats/fifa-to-pes6"
'
'ieApp.FindElementsById("offlineStats").Values = WriteJson(
'Set ieButton = ieApp.FindElementById("convertOffline")
'
'Do While ieApp.Busy Or ieApp.readyState <> READYSTATE_COMPLETE
'      DoEvents
'Loop
'
''Application.Wait TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + vSegundos)
''
''Do While ln <= UltCel.Row
''
''    ie.document.getElementById("doc").vNewValue = W.Cells(ln, col)
''    ie.document.getElementById("consultar").Click
''
''    Application.Wait TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + vSegundos)
''
''    On Error Resume Next
''        vErro = ie.document.getElementById("mensagem").innerText
''
''    On Error GoTo 0
''
''    If vErro = "Informe um termo válido! " Then
''        ie.document.getElementById("consultar").Click
''        Application.Wait TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + vSegundos)
''    ElseIf vErro = "Informe um termo válido! " Then
''        W.Cells(ln, col + 1).vNewValue = "'" & vErro
''    ElseIf Trim(vErro) = "CPF inválido" Then
''        W.Cells(ln, col + 1).vNewValue = "'" & vErro
''    ElseIf Trim(vErro) = "#Erro: Tente novamente!" Then
''        W.Cells(ln, col + 1).vNewValue = "'" & vErro
''    Else
''        vErro = vbNullString
''    End If
''
''    Do While ie.Busy
''    Loop
''
''    If vErro = vbNullString Then
''
''        vNome = ie.document.getElementsByClassName("dados nome")(0).innerText
''        vDados = ie.document.getElementsByClassName("dados texto")(0).innerText
''        vSituacao = ie.document.getElementsByClassName("dados situacao")(0).innerText
''
''        W.Cells(ln, col + 1) = vNome
''        W.Cells(ln, col + 2) = vSituacao
''        W.Cells(ln, col + 3) = vDados
''
''        vNome = vbNullString
''        vDados = vbNullString
''        vSituacao = vbNullString
''
''        ie.document.getElementById("btnVoltar").Click
''
''    Else
''
''        ie.navigate "https://www.situacaocadastral.com.br/"
''        W.Cells(ln, col + 1) = "Dados inválidos para consulta"
''
''    End If
''
''    ln = ln + 1
''
''    Application.Wait TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + vSegundos)
''
''Loop
''
''ie.Quit
''
''W.UsedRange.EntireColumn.AutoFit
''
''Application.ScreenUpdating = True
''
''DoEvents
''MsgBox "Consulta realizada com sucesso!"
''
''Set ie = Nothing
''Set UltCel = Nothing
''Set W = Nothing
'
'End Sub
'
