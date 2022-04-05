Private Function ParseJson(ByVal playerInfos As Variant, ByVal playerSkills As Variant, ByVal playerTraits As Variant) As String

      ' inicializa o objeto dicionario
      Dim dict As New Dictionary
      Dim k As Variant
      Dim json As String, lastKey As String
      Dim i As Integer

      ' inicia criacao do json
      Let json = "{"

      For i = 1 To 3

            Select Case i
                  Case 1
                        Set dict = playerInfos
                  Case 2
                        Set dict = playerSkills
                  Case 3
                        Set dict = playerTraits
            End Select

            If (i = 2) Then
                  json = json & """" & "skills" & """" & ":{"
            ElseIf (i = 3) Then
                  json = json & """" & "player_traits" & """" & ":["
            End If

            For Each k In dict.Keys

                  If (i = 3) Then
                        json = json & """" & dict(k) & """"
                  ElseIf (k = "positions") Then
                        json = json & """" & "positions" & """" & ": [" & dict(k) & "]"
                  ElseIf (IsNumeric(dict(k))) Then
                        json = json & """" & k & """" & ":" & dict(k)
                  Else
                        json = json & """" & k & """" & ":" & """" & dict(k) & """"
                  End If

                  If ((k = dict.Keys(dict.Count - 1)) And (i = 2)) Then
                        json = json & "},"
                  Else
                        json = json & ","
                  End If
            Next k
      Next i

      json = json & "]}"
      json = Replace(json, ",]}", "]}")

      ParseJson = json

End Function