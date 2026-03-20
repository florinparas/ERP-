Attribute VB_Name = "Module_Utils"
'=====================================================
' ERP HR Module - Funcții Utilitare
'=====================================================

Option Explicit

' Validare CNP românesc (verificare de bază)
Public Function ValidareCNP(cnp As String) As Boolean
    ' CNP are exact 13 cifre
    If Len(cnp) <> 13 Then
        ValidareCNP = False
        Exit Function
    End If

    ' Toate caracterele trebuie să fie cifre
    Dim i As Integer
    For i = 1 To 13
        If Not IsNumeric(Mid(cnp, i, 1)) Then
            ValidareCNP = False
            Exit Function
        End If
    Next i

    ' Prima cifră: 1-8
    Dim s As Integer
    s = CInt(Mid(cnp, 1, 1))
    If s < 1 Or s > 8 Then
        ValidareCNP = False
        Exit Function
    End If

    ' Verificare cifră control
    Dim constanta As String
    constanta = "279146358279"
    Dim suma As Long
    suma = 0
    For i = 1 To 12
        suma = suma + CInt(Mid(cnp, i, 1)) * CInt(Mid(constanta, i, 1))
    Next i

    Dim rest As Integer
    rest = suma Mod 11
    If rest = 10 Then rest = 1

    If rest = CInt(Mid(cnp, 13, 1)) Then
        ValidareCNP = True
    Else
        ValidareCNP = False
    End If
End Function

' Extrage data nașterii din CNP
Public Function DataNastereDinCNP(cnp As String) As Date
    If Len(cnp) <> 13 Then
        DataNastereDinCNP = 0
        Exit Function
    End If

    Dim s As Integer
    Dim an As Integer
    Dim luna As Integer
    Dim zi As Integer

    s = CInt(Mid(cnp, 1, 1))
    an = CInt(Mid(cnp, 2, 2))
    luna = CInt(Mid(cnp, 4, 2))
    zi = CInt(Mid(cnp, 6, 2))

    Select Case s
        Case 1, 2: an = 1900 + an
        Case 3, 4: an = 1800 + an
        Case 5, 6: an = 2000 + an
        Case 7, 8: an = 2000 + an  ' Rezidenți
    End Select

    On Error Resume Next
    DataNastereDinCNP = DateSerial(an, luna, zi)
    On Error GoTo 0
End Function

' Extrage sexul din CNP
Public Function SexDinCNP(cnp As String) As String
    If Len(cnp) <> 13 Then
        SexDinCNP = ""
        Exit Function
    End If

    Dim s As Integer
    s = CInt(Mid(cnp, 1, 1))

    Select Case s
        Case 1, 3, 5, 7: SexDinCNP = "M"
        Case 2, 4, 6, 8: SexDinCNP = "F"
        Case Else: SexDinCNP = ""
    End Select
End Function

' Calculează vârsta
Public Function CalculeazaVarsta(dataNasterii As Date) As Integer
    CalculeazaVarsta = DateDiff("yyyy", dataNasterii, Date)
    If Date < DateSerial(Year(Date), Month(dataNasterii), Day(dataNasterii)) Then
        CalculeazaVarsta = CalculeazaVarsta - 1
    End If
End Function

' Calculează vechimea în muncă (ani și luni)
Public Function CalculeazaVechime(dataAngajarii As Date) As String
    Dim ani As Integer
    Dim luni As Integer

    ani = DateDiff("yyyy", dataAngajarii, Date)
    If Date < DateSerial(Year(Date), Month(dataAngajarii), Day(dataAngajarii)) Then
        ani = ani - 1
    End If

    luni = DateDiff("m", dataAngajarii, Date) - (ani * 12)

    CalculeazaVechime = ani & " ani, " & luni & " luni"
End Function

' Backup fișier curent
Sub BackupFisier()
    Dim backupPath As String
    Dim originalName As String

    originalName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    backupPath = ThisWorkbook.Path & "\" & originalName & _
                 "_backup_" & Format(Now, "YYYYMMDD_HHMMSS") & ".xlsm"

    ThisWorkbook.SaveCopyAs backupPath
    MsgBox "Backup creat: " & vbNewLine & backupPath, vbInformation, "Backup"
End Sub

' Protejare foi (cu excepția celulelor editabile)
Sub ProtejeazaFoi()
    Dim ws As Worksheet
    Dim parola As String

    parola = InputBox("Introduceți parola de protecție:", "Protecție Foi")
    If parola = "" Then Exit Sub

    For Each ws In ThisWorkbook.Worksheets
        ws.Protect Password:=parola, _
            UserInterfaceOnly:=True, _
            AllowFiltering:=True, _
            AllowSorting:=True
    Next ws

    MsgBox "Toate foile au fost protejate!", vbInformation
End Sub

' Deprotejare foi
Sub DeprotejeazaFoi()
    Dim ws As Worksheet
    Dim parola As String

    parola = InputBox("Introduceți parola:", "Deprotecție Foi")
    If parola = "" Then Exit Sub

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:=parola
    Next ws
    On Error GoTo 0

    MsgBox "Toate foile au fost deprotejate!", vbInformation
End Sub
