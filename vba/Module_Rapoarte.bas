Attribute VB_Name = "Module_Rapoarte"
'=====================================================
' ERP HR Module - Generare Rapoarte
'=====================================================

Option Explicit

' Raport angajați pe departament
Sub RaportAngajatiDepartament()
    Dim wsAng As Worksheet
    Dim wsDept As Worksheet
    Dim lastRowAng As Long
    Dim lastRowDept As Long
    Dim i As Long, j As Long
    Dim dept As String
    Dim count As Long
    Dim msg As String

    Set wsAng = Sheets("Angajati")
    Set wsDept = Sheets("Departamente")

    lastRowAng = wsAng.Cells(wsAng.Rows.Count, 1).End(xlUp).Row
    lastRowDept = wsDept.Cells(wsDept.Rows.Count, 1).End(xlUp).Row

    msg = "═══════════════════════════════════════" & vbNewLine
    msg = msg & "  RAPORT ANGAJAȚI PE DEPARTAMENT" & vbNewLine
    msg = msg & "  Data: " & Format(Date, "DD.MM.YYYY") & vbNewLine
    msg = msg & "═══════════════════════════════════════" & vbNewLine & vbNewLine

    Dim totalAngajati As Long
    totalAngajati = 0

    For i = 3 To lastRowDept
        dept = wsDept.Cells(i, 2).Value
        If dept <> "" Then
            count = 0
            For j = 3 To lastRowAng
                If wsAng.Cells(j, 11).Value = dept And wsAng.Cells(j, 15).Value = "Activ" Then
                    count = count + 1
                End If
            Next j
            msg = msg & dept & ": " & count & " angajați" & vbNewLine
            totalAngajati = totalAngajati + count
        End If
    Next i

    msg = msg & vbNewLine & "───────────────────────────────────────" & vbNewLine
    msg = msg & "TOTAL ANGAJAȚI ACTIVI: " & totalAngajati & vbNewLine

    MsgBox msg, vbInformation, "Raport Departamente"
End Sub

' Raport concedii luna curentă
Sub RaportConcediiLuna()
    Dim wsConcedii As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim msg As String
    Dim countAprobate As Long
    Dim countAsteptare As Long
    Dim countRespinse As Long

    Set wsConcedii = Sheets("Concedii")
    lastRow = wsConcedii.Cells(wsConcedii.Rows.Count, 1).End(xlUp).Row

    countAprobate = 0
    countAsteptare = 0
    countRespinse = 0

    msg = "═══════════════════════════════════════" & vbNewLine
    msg = msg & "  RAPORT CONCEDII" & vbNewLine
    msg = msg & "  Data: " & Format(Date, "DD.MM.YYYY") & vbNewLine
    msg = msg & "═══════════════════════════════════════" & vbNewLine & vbNewLine

    msg = msg & "CONCEDII ACTIVE/VIITOARE:" & vbNewLine
    msg = msg & "───────────────────────────────────────" & vbNewLine

    For i = 3 To lastRow
        If wsConcedii.Cells(i, 1).Value <> "" Then
            Select Case wsConcedii.Cells(i, 8).Value
                Case "Aprobat": countAprobate = countAprobate + 1
                Case "În Așteptare": countAsteptare = countAsteptare + 1
                Case "Respins": countRespinse = countRespinse + 1
            End Select

            ' Afișează doar concediile active sau viitoare
            If wsConcedii.Cells(i, 6).Value >= Date And wsConcedii.Cells(i, 8).Value <> "Respins" Then
                msg = msg & "  " & wsConcedii.Cells(i, 3).Value & " | "
                msg = msg & wsConcedii.Cells(i, 4).Value & " | "
                msg = msg & Format(wsConcedii.Cells(i, 5).Value, "DD.MM.YYYY") & " - "
                msg = msg & Format(wsConcedii.Cells(i, 6).Value, "DD.MM.YYYY") & " | "
                msg = msg & wsConcedii.Cells(i, 8).Value & vbNewLine
            End If
        End If
    Next i

    msg = msg & vbNewLine & "───────────────────────────────────────" & vbNewLine
    msg = msg & "Aprobate: " & countAprobate & vbNewLine
    msg = msg & "În Așteptare: " & countAsteptare & vbNewLine
    msg = msg & "Respinse: " & countRespinse & vbNewLine

    MsgBox msg, vbInformation, "Raport Concedii"
End Sub

' Raport costuri salariale pe departament
Sub RaportCosturiSalariale()
    Dim wsSal As Worksheet
    Dim wsAng As Worksheet
    Dim wsDept As Worksheet
    Dim lastRowSal As Long
    Dim lastRowDept As Long
    Dim i As Long, j As Long
    Dim msg As String
    Dim dept As String
    Dim totalDept As Double
    Dim totalGeneral As Double

    Set wsSal = Sheets("Salarizare")
    Set wsAng = Sheets("Angajati")
    Set wsDept = Sheets("Departamente")

    lastRowSal = wsSal.Cells(wsSal.Rows.Count, 1).End(xlUp).Row
    Dim lastRowDeptVal As Long
    lastRowDeptVal = wsDept.Cells(wsDept.Rows.Count, 1).End(xlUp).Row

    msg = "═══════════════════════════════════════" & vbNewLine
    msg = msg & "  RAPORT COSTURI SALARIALE" & vbNewLine
    msg = msg & "  Data: " & Format(Date, "DD.MM.YYYY") & vbNewLine
    msg = msg & "═══════════════════════════════════════" & vbNewLine & vbNewLine

    totalGeneral = 0

    For i = 3 To lastRowDeptVal
        dept = wsDept.Cells(i, 2).Value
        If dept <> "" Then
            totalDept = 0
            For j = 3 To lastRowSal
                If wsSal.Cells(j, 2).Value <> "" Then
                    ' Verifică departamentul angajatului
                    Dim idAng As String
                    idAng = wsSal.Cells(j, 2).Value
                    Dim deptAng As String
                    deptAng = Application.WorksheetFunction.VLookup(idAng, _
                        wsAng.Range("A:K"), 11, False)
                    If deptAng = dept Then
                        totalDept = totalDept + wsSal.Cells(j, 19).Value  ' Cost Total Angajator
                    End If
                End If
            Next j
            msg = msg & dept & ": " & Format(totalDept, "#,##0.00") & " RON" & vbNewLine
            totalGeneral = totalGeneral + totalDept
        End If
    Next i

    msg = msg & vbNewLine & "───────────────────────────────────────" & vbNewLine
    msg = msg & "COST TOTAL: " & Format(totalGeneral, "#,##0.00") & " RON" & vbNewLine

    MsgBox msg, vbInformation, "Raport Costuri Salariale"
End Sub
