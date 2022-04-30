Sub Teste()
    Dim ie As New InternetExplorer
    Dim doc As New HTMLDocument
    Range("A2").Select
    Do Until IsEmpty(ActiveCell)
        Dim ecoll   As Object
        ie.Visible = False
        Debug.Print "https://www.vivino.com/search/wines?q=" & LCase(Replace(ActiveCell.Value, " ", "+")) & "+" & ActiveCell.Offset(0, 1).Value
        ie.navigate "https://www.vivino.com/search/wines?q=" & LCase(Replace(ActiveCell.Value, " ", "+")) & "+" & ActiveCell.Offset(0, 1).Value
        Do
            DoEvents
        Loop Until ie.readyState = READYSTATE_COMPLETE
        Application.Wait (Now + TimeValue("0:00:02"))
        Set doc = ie.document
        ActiveCell.Offset(0, 4).Value = CDec(doc.getElementsByClassName("wine-price-value")(0).innerHTML)
        ActiveCell.Offset(0, 3) = doc.getElementsByClassName("wine-card__name")(0).innerText
        ActiveCell.Offset(0, 5) = CDec(doc.getElementsByClassName("average__number")(0).innerText)
        ActiveCell.Offset(1, 0).Select
    Loop
End Sub
