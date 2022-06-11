Sub PreencherVivino()
    Dim ie As New InternetExplorer
    Dim doc As New HTMLDocument
    Range(Range("K2").Value).Select
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
        If doc.getElementsByClassName("wine-price-value")(0).innerHTML = "—" Then
            ActiveCell.Offset(0, 4).Value = 0
        Else
            ActiveCell.Offset(0, 4).Value = CDec(doc.getElementsByClassName("wine-price-value")(0).innerHTML)
        End If

        If doc.getElementsByClassName("wine-card__name")(0).innerText = "—" Then
            ActiveCell.Offset(0, 3).Value = "N/A"
        Else
            ActiveCell.Offset(0, 3).Value = doc.getElementsByClassName("wine-card__name")(0).innerText
        End If

        If doc.getElementsByClassName("wine-card__region")(0).innerText = "—" Then
            ActiveCell.Offset(0, 5).Value = "N/A"
        Else
            ActiveCell.Offset(0, 5) = doc.getElementsByClassName("wine-card__region")(0).innerText
        End If

        If doc.getElementsByClassName("average__number")(0).innerHTML = "—" Then
            ActiveCell.Offset(0, 6).Value = 0
        Else
            ActiveCell.Offset(0, 6).Value = doc.getElementsByClassName("average__number")(0).innerText
        End If
        ActiveCell.Offset(1, 0).Select
    Loop
End Sub
