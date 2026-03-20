Attribute VB_Name = "Module_Pontaj"
'=====================================================
' ERP HR Module - Automatizare Pontaj
'=====================================================

Option Explicit

' Generează tabel pontaj pentru luna curentă
' Populează automat cu toți angajații activi
Sub GenereazaPontajLunar()
    Dim wsPontaj As Worksheet
    Dim wsAngajati As Worksheet
    Dim lastRowAng As Long
    Dim lastRowPontaj As Long
    Dim luna As Integer
    Dim an As Integer
    Dim i As Long
    Dim newRow As Long

    Set wsPontaj = Sheets("Pontaj")
    Set wsAngajati = Sheets("Angajati")

    ' Solicită luna și anul
    luna = Month(Date)
    an = Year(Date)

    luna = Application.InputBox("Introduceți luna (1-12):", "Pontaj Lunar", luna, Type:=1)
    If luna < 1 Or luna > 12 Then
        MsgBox "Luna invalidă!", vbExclamation
        Exit Sub
    End If

    an = Application.InputBox("Introduceți anul:", "Pontaj Lunar", an, Type:=1)
    If an < 2020 Or an > 2030 Then
        MsgBox "An invalid!", vbExclamation
        Exit Sub
    End If

    ' Verifică dacă pontajul pentru luna respectivă există deja
    lastRowPontaj = wsPontaj.Cells(wsPontaj.Rows.Count, 1).End(xlUp).Row
    Dim exists As Boolean
    exists = False
    For i = 3 To lastRowPontaj
        If wsPontaj.Cells(i, 4).Value = luna And wsPontaj.Cells(i, 5).Value = an Then
            exists = True
            Exit For
        End If
    Next i

    If exists Then
        If MsgBox("Pontajul pentru " & luna & "/" & an & " există deja. Continuați?", _
                   vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If

    ' Adaugă rânduri pentru fiecare angajat activ
    lastRowAng = wsAngajati.Cells(wsAngajati.Rows.Count, 1).End(xlUp).Row
    newRow = lastRowPontaj + 1

    Application.ScreenUpdating = False

    For i = 3 To lastRowAng
        ' Verifică dacă angajatul este activ
        If wsAngajati.Cells(i, 15).Value = "Activ" Then
            wsPontaj.Cells(newRow, 1).Value = wsAngajati.Cells(i, 1).Value ' ID
            wsPontaj.Cells(newRow, 4).Value = luna
            wsPontaj.Cells(newRow, 5).Value = an

            ' Populează zilele cu "P" pentru zile lucrătoare, "LS" pentru weekend
            Dim d As Integer
            Dim dayDate As Date
            Dim daysInMonth As Integer
            daysInMonth = Day(DateSerial(an, luna + 1, 0))

            For d = 1 To 31
                If d <= daysInMonth Then
                    dayDate = DateSerial(an, luna, d)
                    If Weekday(dayDate, vbMonday) <= 5 Then
                        wsPontaj.Cells(newRow, 5 + d).Value = "P"
                    Else
                        wsPontaj.Cells(newRow, 5 + d).Value = "LS"
                    End If
                Else
                    wsPontaj.Cells(newRow, 5 + d).Value = ""
                End If
            Next d

            newRow = newRow + 1
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Pontaj generat pentru " & luna & "/" & an & "!" & vbNewLine & _
           "Angajați adăugați: " & (newRow - lastRowPontaj - 1), vbInformation
End Sub

' Calculează totalurile pentru pontajul selectat
Sub CalculeazaTotaluriPontaj()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim d As Integer
    Dim dayRange As String

    Set ws = Sheets("Pontaj")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Application.ScreenUpdating = False

    For i = 3 To lastRow
        If ws.Cells(i, 1).Value <> "" Then
            ' Totaluri folosind COUNTIF
            ' Coloana 37 = Zile Lucrate (P + TP)
            ws.Cells(i, 37).FormulaR1C1 = "=COUNTIF(RC[-31]:RC[-1],""P"")+COUNTIF(RC[-31]:RC[-1],""TP"")"
            ' Coloana 38 = Total CO
            ws.Cells(i, 38).FormulaR1C1 = "=COUNTIF(RC[-32]:RC[-2],""CO"")"
            ' Coloana 39 = Total CM
            ws.Cells(i, 39).FormulaR1C1 = "=COUNTIF(RC[-33]:RC[-3],""CM"")"
            ' Coloana 40 = Total Absențe
            ws.Cells(i, 40).FormulaR1C1 = "=COUNTIF(RC[-34]:RC[-4],""A"")+COUNTIF(RC[-34]:RC[-4],""AM"")"
            ' Coloana 41 = Total OS
            ws.Cells(i, 41).FormulaR1C1 = "=COUNTIF(RC[-35]:RC[-5],""OS"")"
        End If
    Next i

    Application.ScreenUpdating = True
    MsgBox "Totaluri pontaj recalculate!", vbInformation
End Sub
