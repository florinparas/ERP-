Attribute VB_Name = "Module_Navigation"
'=====================================================
' ERP HR Module - Navigare între foi
' Importați acest modul în Excel VBA Editor
'=====================================================

Option Explicit

' Navigare către foaia specificată
Sub NavigateTo(sheetName As String)
    On Error Resume Next
    Sheets(sheetName).Activate
    Range("A1").Select
    On Error GoTo 0
End Sub

' Navigare rapidă - Dashboard
Sub GoToDashboard()
    NavigateTo "Dashboard"
End Sub

Sub GoToAngajati()
    NavigateTo "Angajati"
End Sub

Sub GoToContracte()
    NavigateTo "Contracte"
End Sub

Sub GoToDepartamente()
    NavigateTo "Departamente"
End Sub

Sub GoToPontaj()
    NavigateTo "Pontaj"
End Sub

Sub GoToConcedii()
    NavigateTo "Concedii"
End Sub

Sub GoToSalarizare()
    NavigateTo "Salarizare"
End Sub

Sub GoToEvaluari()
    NavigateTo "Evaluari"
End Sub

Sub GoToTraining()
    NavigateTo "Training"
End Sub

Sub GoToRecrutare()
    NavigateTo "Recrutare"
End Sub

Sub GoToConfigurare()
    NavigateTo "Configurare"
End Sub

' Meniu principal de navigare
Sub ShowNavigationMenu()
    Dim choice As Integer
    Dim msg As String

    msg = "NAVIGARE ERP - MODUL HR" & vbNewLine & vbNewLine
    msg = msg & "1 - Dashboard" & vbNewLine
    msg = msg & "2 - Angajati" & vbNewLine
    msg = msg & "3 - Contracte" & vbNewLine
    msg = msg & "4 - Departamente" & vbNewLine
    msg = msg & "5 - Pontaj" & vbNewLine
    msg = msg & "6 - Concedii" & vbNewLine
    msg = msg & "7 - Salarizare" & vbNewLine
    msg = msg & "8 - Evaluari" & vbNewLine
    msg = msg & "9 - Training" & vbNewLine
    msg = msg & "10 - Recrutare" & vbNewLine
    msg = msg & "11 - Configurare" & vbNewLine

    choice = Application.InputBox(msg, "Navigare", Type:=1)

    Select Case choice
        Case 1: GoToDashboard
        Case 2: GoToAngajati
        Case 3: GoToContracte
        Case 4: GoToDepartamente
        Case 5: GoToPontaj
        Case 6: GoToConcedii
        Case 7: GoToSalarizare
        Case 8: GoToEvaluari
        Case 9: GoToTraining
        Case 10: GoToRecrutare
        Case 11: GoToConfigurare
    End Select
End Sub
