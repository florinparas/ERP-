Attribute VB_Name = "Module_Salarizare"
'=====================================================
' ERP HR Module - Calcul Salarizare & Fluturași
'=====================================================

Option Explicit

' Constante fiscale (pot fi citite și din foaia Configurare)
Private Const CAS_RATE As Double = 0.25
Private Const CASS_RATE As Double = 0.1
Private Const TAX_RATE As Double = 0.1
Private Const CAM_RATE As Double = 0.0225

' Generează stat de plată pentru o lună
Sub GenereazaStatPlata()
    Dim wsSal As Worksheet
    Dim wsAng As Worksheet
    Dim wsPontaj As Worksheet
    Dim wsConfig As Worksheet
    Dim lastRowAng As Long
    Dim lastRowSal As Long
    Dim luna As Integer
    Dim an As Integer
    Dim i As Long
    Dim newRow As Long
    Dim salariuBrut As Double
    Dim zileLucrate As Long
    Dim zileTotale As Long
    Dim idAngajat As String
    Dim salID As Long

    Set wsSal = Sheets("Salarizare")
    Set wsAng = Sheets("Angajati")
    Set wsPontaj = Sheets("Pontaj")

    ' Solicită luna și anul
    luna = Month(Date)
    an = Year(Date)

    luna = Application.InputBox("Luna pentru stat de plată (1-12):", "Salarizare", luna, Type:=1)
    an = Application.InputBox("Anul:", "Salarizare", an, Type:=1)

    lastRowAng = wsAng.Cells(wsAng.Rows.Count, 1).End(xlUp).Row
    lastRowSal = wsSal.Cells(wsSal.Rows.Count, 1).End(xlUp).Row
    newRow = lastRowSal + 1
    salID = lastRowSal - 1  ' Aproximativ

    Application.ScreenUpdating = False

    For i = 3 To lastRowAng
        If wsAng.Cells(i, 15).Value = "Activ" Then
            idAngajat = wsAng.Cells(i, 1).Value
            salID = salID + 1

            ' ID
            wsSal.Cells(newRow, 1).Value = "S" & Format(salID, "000")

            ' ID Angajat
            wsSal.Cells(newRow, 2).Value = idAngajat

            ' Luna, An
            wsSal.Cells(newRow, 4).Value = luna
            wsSal.Cells(newRow, 5).Value = an

            ' Salariu Brut din contract
            salariuBrut = GetSalariuBrut(idAngajat)
            wsSal.Cells(newRow, 6).Value = salariuBrut

            ' Zile lucrate din pontaj
            zileLucrate = GetZileLucratePontaj(idAngajat, luna, an)
            wsSal.Cells(newRow, 7).Value = zileLucrate

            ' Zile totale lucrătoare în lună
            zileTotale = GetZileLucratoareLuna(luna, an)
            wsSal.Cells(newRow, 8).Value = zileTotale

            ' Restul se calculează cu formule (deja setate în template)
            ' Dar le calculăm și direct pentru siguranță
            Dim salProp As Double
            Dim cas As Double
            Dim cass As Double
            Dim bazaImpoz As Double
            Dim impozit As Double
            Dim salNet As Double
            Dim cam As Double

            If zileTotale > 0 Then
                salProp = Round(salariuBrut * zileLucrate / zileTotale, 2)
            Else
                salProp = 0
            End If

            cas = Round(salProp * CAS_RATE, 2)
            cass = Round(salProp * CASS_RATE, 2)
            bazaImpoz = Application.WorksheetFunction.Max(salProp - cas - cass, 0)
            impozit = Round(bazaImpoz * TAX_RATE, 2)
            salNet = salProp - cas - cass - impozit
            cam = Round(salProp * CAM_RATE, 2)

            wsSal.Cells(newRow, 9).Value = salProp
            wsSal.Cells(newRow, 10).Value = cas
            wsSal.Cells(newRow, 11).Value = cass
            wsSal.Cells(newRow, 12).Value = bazaImpoz
            wsSal.Cells(newRow, 13).Value = 0  ' Deducere personală
            wsSal.Cells(newRow, 14).Value = impozit
            wsSal.Cells(newRow, 15).Value = 0  ' Alte deduceri
            wsSal.Cells(newRow, 16).Value = 0  ' Tichete masă
            wsSal.Cells(newRow, 17).Value = salNet
            wsSal.Cells(newRow, 18).Value = cam
            wsSal.Cells(newRow, 19).Value = salProp + cam

            ' Formatare RON
            Dim c As Integer
            For c = 6 To 19
                If c <> 7 And c <> 8 Then
                    wsSal.Cells(newRow, c).NumberFormat = "#,##0.00 ""RON"""
                End If
            Next c

            newRow = newRow + 1
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Stat de plată generat pentru " & luna & "/" & an & "!" & vbNewLine & _
           "Angajați procesați: " & (newRow - lastRowSal - 1), vbInformation
End Sub

' Obține salariul brut din foaia Contracte
Private Function GetSalariuBrut(idAngajat As String) As Double
    Dim wsContracte As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set wsContracte = Sheets("Contracte")
    lastRow = wsContracte.Cells(wsContracte.Rows.Count, 1).End(xlUp).Row

    GetSalariuBrut = 0
    For i = 3 To lastRow
        If wsContracte.Cells(i, 2).Value = idAngajat And _
           wsContracte.Cells(i, 11).Value = "Activ" Then
            GetSalariuBrut = wsContracte.Cells(i, 9).Value
            Exit For
        End If
    Next i
End Function

' Obține zilele lucrate din pontaj
Private Function GetZileLucratePontaj(idAngajat As String, luna As Integer, an As Integer) As Long
    Dim wsPontaj As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set wsPontaj = Sheets("Pontaj")
    lastRow = wsPontaj.Cells(wsPontaj.Rows.Count, 1).End(xlUp).Row

    GetZileLucratePontaj = 0
    For i = 3 To lastRow
        If wsPontaj.Cells(i, 1).Value = idAngajat And _
           wsPontaj.Cells(i, 4).Value = luna And _
           wsPontaj.Cells(i, 5).Value = an Then
            GetZileLucratePontaj = wsPontaj.Cells(i, 37).Value  ' Coloana Total Zile Lucrate
            Exit For
        End If
    Next i
End Function

' Calculează zilele lucrătoare într-o lună
Private Function GetZileLucratoareLuna(luna As Integer, an As Integer) As Long
    Dim d As Date
    Dim count As Long
    Dim lastDay As Integer

    lastDay = Day(DateSerial(an, luna + 1, 0))
    count = 0

    Dim i As Integer
    For i = 1 To lastDay
        d = DateSerial(an, luna, i)
        If Weekday(d, vbMonday) <= 5 Then
            count = count + 1
        End If
    Next i

    GetZileLucratoareLuna = count
End Function

' Generează fluturaș de salariu pentru un angajat
Sub GenereazaFluturas()
    Dim wsSal As Worksheet
    Dim rowSel As Long

    Set wsSal = Sheets("Salarizare")

    If ActiveSheet.Name <> "Salarizare" Then
        MsgBox "Navigați la foaia Salarizare și selectați un rând!", vbExclamation
        Exit Sub
    End If

    rowSel = ActiveCell.Row
    If rowSel < 3 Then
        MsgBox "Selectați un rând cu date de salarizare!", vbExclamation
        Exit Sub
    End If

    Dim msg As String
    msg = "═══════════════════════════════════════" & vbNewLine
    msg = msg & "        FLUTURAȘ DE SALARIU" & vbNewLine
    msg = msg & "═══════════════════════════════════════" & vbNewLine & vbNewLine
    msg = msg & "Angajat: " & wsSal.Cells(rowSel, 3).Value & vbNewLine
    msg = msg & "Luna/An: " & wsSal.Cells(rowSel, 4).Value & "/" & wsSal.Cells(rowSel, 5).Value & vbNewLine
    msg = msg & "───────────────────────────────────────" & vbNewLine
    msg = msg & "Salariu Brut:        " & Format(wsSal.Cells(rowSel, 6).Value, "#,##0.00") & " RON" & vbNewLine
    msg = msg & "Zile Lucrate:        " & wsSal.Cells(rowSel, 7).Value & " / " & wsSal.Cells(rowSel, 8).Value & vbNewLine
    msg = msg & "Salariu Proporțional: " & Format(wsSal.Cells(rowSel, 9).Value, "#,##0.00") & " RON" & vbNewLine
    msg = msg & "───────────────────────────────────────" & vbNewLine
    msg = msg & "REȚINERI:" & vbNewLine
    msg = msg & "  CAS (25%):         -" & Format(wsSal.Cells(rowSel, 10).Value, "#,##0.00") & " RON" & vbNewLine
    msg = msg & "  CASS (10%):        -" & Format(wsSal.Cells(rowSel, 11).Value, "#,##0.00") & " RON" & vbNewLine
    msg = msg & "  Impozit (10%):     -" & Format(wsSal.Cells(rowSel, 14).Value, "#,##0.00") & " RON" & vbNewLine
    msg = msg & "  Alte Deduceri:     -" & Format(wsSal.Cells(rowSel, 15).Value, "#,##0.00") & " RON" & vbNewLine
    msg = msg & "───────────────────────────────────────" & vbNewLine
    msg = msg & "Tichete Masă:        +" & Format(wsSal.Cells(rowSel, 16).Value, "#,##0.00") & " RON" & vbNewLine
    msg = msg & "═══════════════════════════════════════" & vbNewLine
    msg = msg & "SALARIU NET:         " & Format(wsSal.Cells(rowSel, 17).Value, "#,##0.00") & " RON" & vbNewLine
    msg = msg & "═══════════════════════════════════════"

    MsgBox msg, vbInformation, "Fluturaș de Salariu"
End Sub
