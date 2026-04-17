# Quiz-de-pitagoras
Sub CriarQuizPitagorasComSom()
    Dim pptPres As Object: Set pptPres = ActivePresentation
    Dim sli As Object
    Dim shp As Object
    Dim i As Integer, j As Integer
    Dim perguntas(1 To 10) As String
    Dim opcoes(1 To 10, 1 To 4) As String
    Dim corretas(1 To 10) As Integer

    ' --- BANCO DE DADOS ---
    perguntas(1) = "Qual é a fórmula do Teorema de Pitágoras?"
    opcoes(1, 1) = "a² + b² = c²": opcoes(1, 2) = "a + b = c": opcoes(1, 3) = "a² - b² = c²": opcoes(1, 4) = "L * L = A"
    corretas(1) = 1

    perguntas(2) = "O Teorema de Pitágoras aplica-se a qual triângulo?"
    opcoes(2, 1) = "Equilátero": opcoes(2, 2) = "Isósceles": opcoes(2, 3) = "Retângulo": opcoes(2, 4) = "Acutângulo"
    corretas(2) = 3

    perguntas(3) = "Como se chama o lado oposto ao ângulo de 90°?"
    opcoes(3, 1) = "Cateto adjacente": opcoes(3, 2) = "Hipotenusa": opcoes(3, 3) = "Cateto oposto": opcoes(3, 4) = "Base"
    corretas(3) = 2

    perguntas(4) = "Catetos 3 e 4. Quanto mede a hipotenusa?"
    opcoes(4, 1) = "5": opcoes(4, 2) = "7": opcoes(4, 3) = "25": opcoes(4, 4) = "6"
    corretas(4) = 1

    perguntas(5) = "Hipotenusa 10 e um cateto 6. Qual o outro cateto?"
    opcoes(5, 1) = "4": opcoes(5, 2) = "16": opcoes(5, 3) = "8": opcoes(5, 4) = "64"
    corretas(5) = 3

    ' (Complete as outras 5 perguntas seguindo o padrão acima para chegar a 10)

    ' --- CONSTRUÇÃO ---
    For i = 1 To 5 ' Altere para 10 se preencher todas acima
        Set sli = pptPres.Slides.Add(pptPres.Slides.Count + 1, 12)
        sli.FollowMasterBackground = False
        sli.Background.Fill.ForeColor.RGB = RGB(15, 30, 60)

        ' Pergunta
        Set shp = sli.Shapes.AddTextbox(1, 50, 40, 800, 80)
        With shp.TextFrame.TextRange
            .Text = "QUESTÃO " & i & vbCrLf & perguntas(i)
            .Font.Size = 34: .Font.Color.RGB = RGB(255, 255, 255): .Font.Bold = True
            .ParagraphFormat.Alignment = 2
        End With

        ' Botões
        For j = 1 To 4
            Set shp = sli.Shapes.AddShape(5, 150, 150 + (j * 75), 600, 60)
            With shp
                .Fill.ForeColor.RGB = RGB(70, 70, 70)
                .TextFrame.TextRange.Text = Chr(64 + j) & ") " & opcoes(i, j)
                .TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)

                ' AÇÃO DE SOM AO CLICAR
                With .ActionSettings(1) ' 1 = ppMouseClick
                    .Action = 0 ' ppActionNone (não pula slide, só toca som)
                    If j = corretas(i) Then
                        .SoundEffect.Name = "Applause" ' Som de Aplausos
                    Else
                        .SoundEffect.Name = "Bomb" ' Som de Explosão/Erro
                    End If
                End With

                ' ANIMAÇÃO DE COR
                Dim eff As Object
                Set eff = sli.TimeLine.MainSequence.AddEffect(shp, 10, , 1)
                If j = corretas(i) Then
                    eff.EffectParameters.Color2.RGB = RGB(0, 180, 0)
                Else
                    eff.EffectParameters.Color2.RGB = RGB(180, 0, 0)
                End If
            End With
        Next j
    Next i

    MsgBox "Quiz com sons e cores criado!", vbInformation
End Sub
