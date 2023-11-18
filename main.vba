Sub PreencherVivino()
    ' Inicializa o Internet Explorer para fazer scraping
    Dim ie As New InternetExplorer
    Dim doc As New HTMLDocument
    Dim celulaAtiva As Range

    ' Define a célula inicial para a busca
    Set celulaAtiva = Range(Range("K2").Value)

    ' Loop para percorrer as células enquanto não estiverem vazias
    Do Until IsEmpty(celulaAtiva)
        ie.Visible = False

        ' Constrói a URL de busca com os valores da célula e seu offset
        Dim searchUrl As String
        searchUrl = "https://www.vivino.com/search/wines?q=" & _
                     LCase(Replace(celulaAtiva.Value, " ", "+")) & "+" & _
                     celulaAtiva.Offset(0, 1).Value
        Debug.Print searchUrl
        ie.navigate searchUrl

        ' Aguarda o carregamento completo da página
        Do
            DoEvents
        Loop Until ie.readyState = READYSTATE_COMPLETE
        Application.Wait (Now + TimeValue("0:00:02"))
        Set doc = ie.document

        ' Extrai e atribui os dados desejados às células correspondentes
        AtualizarCelula doc, celulaAtiva, "wine-price-value", 4, "—", 0, True
        AtualizarCelula doc, celulaAtiva, "wine-card__name", 3, "—", "N/A", False
        AtualizarCelula doc, celulaAtiva, "wine-card__region", 5, "—", "N/A", False
        AtualizarCelula doc, celulaAtiva, "average__number", 6, "—", 0, False

        ' Move para a próxima célula
        Set celulaAtiva = celulaAtiva.Offset(1, 0)
    Loop
End Sub

' Função auxiliar para atualizar as células com os dados extraídos
Private Sub AtualizarCelula(doc As HTMLDocument, celula As Range, className As String, _
                            offset As Integer, defaultText As String, defaultValue As Variant, _
                            isNumeric As Boolean)
    Dim element As Object
    Set element = doc.getElementsByClassName(className)(0)

    ' Verifica se o elemento está presente e atualiza a célula
    If IsNull(element) Or element.innerText = defaultText Then
        celula.Offset(0, offset).Value = defaultValue
    Else
        If isNumeric Then
            celula.Offset(0, offset).Value = CDec(element.innerText)
        Else
            celula.Offset(0, offset).Value = element.innerText
        End If
    End If
End Sub
