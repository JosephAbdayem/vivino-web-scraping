Sub PreencherVivino()
    Dim ie As New InternetExplorer
    Dim doc As New HTMLDocument
    Dim celulaAtiva As Range
    Set celulaAtiva = Range(Range("K2").Value)

    Do Until IsEmpty(celulaAtiva)
        Dim ecoll   As Object
        ie.Visible = False
        Debug.Print "https://www.vivino.com/search/wines?q=" & LCase(Replace(celulaAtiva.Value, " ", "+")) & "+" & celulaAtiva.Offset(0, 1).Value
        ie.navigate "https://www.vivino.com/search/wines?q=" & LCase(Replace(celulaAtiva.Value, " ", "+")) & "+" & celulaAtiva.Offset(0, 1).Value
        Do
            DoEvents
        Loop Until ie.readyState = READYSTATE_COMPLETE
        Application.Wait (Now + TimeValue("0:00:02"))
        Set doc = ie.document

        'Verificar se o valor do preço está presente
        If doc.getElementsByClassName("wine-price-value")(0).innerHTML = "—" Then
            celulaAtiva.Offset(0, 4).Value = 0
        Else
            celulaAtiva.Offset(0, 4).Value = CDec(doc.getElementsByClassName("wine-price-value")(0).innerHTML)
        End If

        'Verificar se o nome está presente
        If doc.getElementsByClassName("wine-card__name")(0).innerText = "—" Then
            celulaAtiva.Offset(0, 3).Value = "N/A"
        Else
            celulaAtiva.Offset(0, 3).Value = doc.getElementsByClassName("wine-card__name")(0).innerText
        End If

        'Verificar se a região está presente
        If doc.getElementsByClassName("wine-card__region")(0).innerText = "—" Then
            celulaAtiva.Offset(0, 5).Value = "N/A"
        Else
            celulaAtiva.Offset(0, 5) = doc.getElementsByClassName("wine-card__region")(0).innerText
        End If

        'Verificar se a nota média está presente
        If doc.getElementsByClassName("average__number")(0).innerHTML = "—" Then
            celulaAtiva.Offset(0, 6).Value = 0
        Else
            celulaAtiva.Offset(0, 6).Value = doc.getElementsByClassName("average__number")(0).innerText
        End If

        Set celulaAtiva = celulaAtiva.Offset(1, 0)
    Loop
End Sub
