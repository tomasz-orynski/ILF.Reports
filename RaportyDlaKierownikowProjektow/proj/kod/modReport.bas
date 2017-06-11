Attribute VB_Name = "modReport"
Option Base 1
Option Explicit

Public Sub Generate()
On Error GoTo errHandler
Dim data As clsReportsData
Dim dtStart As Date
Dim dtEnd As Date
Dim DictProjects As Scripting.Dictionary
Dim DictProjectTeams As Scripting.Dictionary

    Set data = New clsReportsData
    data.dtStart = CDate(sheetReport.Cells(3, 3))
    data.dtEnd = CDate(sheetReport.Cells(4, 3))
    Set data.DictProjects = ReadProjects
    Set data.DictProjectTeams = ReadProjectTeams
    WriteReports data
    Exit Sub
    
errHandler:
    MsgBox Err.Description
End Sub

Private Sub WriteReports(data As clsReportsData)
Dim v As Variant

    For Each v In data.DictProjectTeams.Keys
        WriteReport v, data
    Next
End Sub

Private Sub WriteReport(ByVal teamKey As String, data As clsReportsData)
Dim v As Variant
Dim s As String
Dim path As String
Dim title As String
Dim info As String
Dim eWorkbook As Excel.Workbook
Dim eWorkSheet As Excel.Worksheet
Dim team As clsTeam
Dim cnt As Long
Dim row As Long
Dim row2 As Long
Dim projName As Variant
Dim emplName As Variant
Dim dictConsumed As Scripting.Dictionary
Dim dictTotalBudget As Scripting.Dictionary
Dim dblTotalMH As Double
Dim dictPlannedActComp As Scripting.Dictionary
Dim dictPlannedActComp2 As Scripting.Dictionary
Dim dictPlannedActCompE As Scripting.Dictionary
Dim delRow As Boolean

    Set team = data.DictProjectTeams(teamKey)
    title = "Raport dla kierownika pionu " & team.DivisionLeader & " za okres " & Format$(data.dtStart, "YYYYMMDD") & "-" & Format$(data.dtEnd, "YYYYMMDD")
    info = "W za³¹czeniu raport za okres od " & Format$(data.dtStart, "YYYY-MM-DD") & " do " & Format$(data.dtEnd, "YYYY-MM-DD") & " dla (" & team.DivisionLeader & ")." & vbCrLf _
        & "Proszê o akceptacjê zestawienia przy u¿yciu przycisku confirmed w za³¹czonym pliku oraz wpisywanie swoich uwag."
    
    path = Workbook.path & "\szablony\template.xltm"
    Set eWorkbook = Application.Workbooks.Add(path)
    eWorkbook.Activate
    Set eWorkSheet = eWorkbook.Sheets.Item(1)
    eWorkSheet.Activate
    eWorkSheet.Range("H2").Value = data.dtStart
    eWorkSheet.Range("J2").Value = data.dtEnd
    eWorkSheet.Range("C4").Value = team.AreaName
    eWorkSheet.Range("C5").Value = team.DivisionNameShort
    eWorkSheet.Range("C6").Value = team.TeamName
    eWorkSheet.Range("C7").Value = team.TeamLeader
    eWorkSheet.Range("C8").Value = team.DivisionLeader
    
    cnt = team.DictEmployees.Count * data.DictProjects.Count - 1
    row = 17
    While cnt > 0
        eWorkSheet.Rows(row + 1).Select
        Selection.EntireRow.Insert
        eWorkSheet.Rows(row).Select
        Selection.Copy
        eWorkSheet.Rows(row + 1).Select
        eWorkSheet.Paste
        cnt = cnt - 1
    Wend
    
    For Each projName In data.DictProjects.Keys
        Call ReadProjectTeamEmployees_TotalBudget(teamKey, projName, data.dtStart, dictTotalBudget)
        Call ReadProjectTeamEmployees_ConsumedAndTotalMH(teamKey, projName, data.dtStart, data.dtEnd, dictConsumed, dblTotalMH)
        Call ReadProjectTeamEmployees_PlannedActComp(teamKey, projName, data.dtStart, data.dtEnd, dictPlannedActComp2, dictPlannedActComp)
        Call ReadProjectTeamEmployees_PlannedActCompE(teamKey, projName, data.dtStart, data.dtEnd, dictPlannedActCompE)
        For Each emplName In team.DictEmployees.Keys
            delRow = True
            eWorkSheet.Range("B" & row).Value = projName
            eWorkSheet.Range("C" & row).Value = emplName
            
            If dictPlannedActComp2.Exists(emplName) Then
                eWorkSheet.Range("D" & row).Value = dictPlannedActComp2(emplName)
                delRow = delRow And dictPlannedActComp2(emplName) = 0
            Else
                eWorkSheet.Range("D" & row).Value = 0
            End If
            If dictConsumed.Exists(emplName) Then
                eWorkSheet.Range("E" & row).Value = dictConsumed(emplName)
                delRow = delRow And dictConsumed(emplName) = 0
            Else
                eWorkSheet.Range("E" & row).Value = 0
            End If
            If dictTotalBudget.Exists(emplName) Then
                eWorkSheet.Range("H" & row).Value = dictTotalBudget(emplName)
            Else
                eWorkSheet.Range("H" & row).Value = 0
            End If
            eWorkSheet.Range("I" & row).Value = dblTotalMH
            If dictPlannedActCompE.Exists(emplName) Then
                eWorkSheet.Range("L" & row).Value = dictPlannedActCompE(emplName)
                delRow = delRow And dictPlannedActCompE(emplName) = 0
            Else
                eWorkSheet.Range("L" & row).Value = 0
            End If
            
            If delRow Then
                eWorkSheet.Rows(row).Select
                Selection.Delete
            Else
                row = row + 1
            End If
        Next
    Next
    
    row2 = MakeTable(eWorkSheet, 17, row) + 7
    row = row2
    
    cnt = team.DictEmployees.Count * data.DictProjects.Count - 1
    While cnt > 0
        eWorkSheet.Rows(row + 1).Select
        Selection.EntireRow.Insert
        eWorkSheet.Rows(row).Select
        Selection.Copy
        eWorkSheet.Rows(row + 1).Select
        eWorkSheet.Paste
        cnt = cnt - 1
    Wend
    
    For Each projName In data.DictProjects.Keys
        Call ReadProjectTeamEmployees_TotalBudget2(teamKey, projName, data.dtStart, dictTotalBudget)
        Call ReadProjectTeamEmployees_ConsumedAndTotalMH2(teamKey, projName, data.dtStart, data.dtEnd, dictConsumed, dblTotalMH)
        Call ReadProjectTeamEmployees_PlannedActComp2(teamKey, projName, data.dtStart, data.dtEnd, dictPlannedActComp2, dictPlannedActComp)
        Call ReadProjectTeamEmployees_PlannedActCompE2(teamKey, projName, data.dtStart, data.dtEnd, dictPlannedActCompE)
        For Each emplName In team.DictEmployees.Keys
            delRow = True
            eWorkSheet.Range("B" & row).Value = projName
            eWorkSheet.Range("C" & row).Value = emplName
            
            If dictPlannedActComp2.Exists(emplName) Then
                eWorkSheet.Range("D" & row).Value = dictPlannedActComp2(emplName)
                delRow = delRow And dictPlannedActComp2(emplName) = 0
            Else
                eWorkSheet.Range("D" & row).Value = 0
            End If
            If dictConsumed.Exists(emplName) Then
                eWorkSheet.Range("E" & row).Value = dictConsumed(emplName)
                delRow = delRow And dictConsumed(emplName) = 0
            Else
                eWorkSheet.Range("E" & row).Value = 0
            End If
            If dictTotalBudget.Exists(emplName) Then
                eWorkSheet.Range("H" & row).Value = dictTotalBudget(emplName)
            Else
                eWorkSheet.Range("H" & row).Value = 0
            End If
            eWorkSheet.Range("I" & row).Value = dblTotalMH
            If dictPlannedActCompE.Exists(emplName) Then
                eWorkSheet.Range("L" & row).Value = dictPlannedActCompE(emplName)
                delRow = delRow And dictPlannedActCompE(emplName) = 0
            Else
                eWorkSheet.Range("L" & row).Value = 0
            End If
            
            If delRow Then
                eWorkSheet.Rows(row).Select
                Selection.Delete
            Else
                row = row + 1
            End If
        Next
    Next
    
    MakeTable eWorkSheet, row2, row
    
    eWorkSheet.Range("A1").Select
    
    eWorkSheet.Outline.ShowLevels 2
    eWorkSheet.Calculate
    eWorkSheet.Cells.Locked = True
    eWorkSheet.Range("F5:O7").Locked = False
    'eWorkSheet.Protect "haslo"
    eWorkbook.CustomDocumentProperties.Item("_SAVE_PATH_").Value = team.SaveEmailPath
    'eWorkbook.Protect "haslo"
    path = Workbook.path & "\dokumenty\Raport." & Format(Now, "YYYYMMDDhhmmss") & ".xlsm"
    eWorkbook.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    eWorkbook.Close False
    Set eWorkbook = Nothing
    SendReport path, title, info, team.DivisionLeaderEmail
End Sub

Private Function MakeTable(ByVal eWorkSheet As Excel.Worksheet, ByVal rowBeg As Long, ByVal rowEnd As Long) As Long
Dim row As Long
Dim s As String

    If rowBeg = rowEnd Then
        eWorkSheet.Rows(rowBeg).Select
    Else
        eWorkSheet.Range("B" & rowBeg - 1 & ":O" & rowEnd - 1).Select
        Application.CutCopyMode = False
        Selection.Subtotal GroupBy:=1, Function:=xlSum, _
            TotalList:=Array(3, 4, 5, 7, 8, 10, 11, 12, 13), _
            Replace:=True, PageBreaks:=False, SummaryBelowData:=True
        
        Do
            s = eWorkSheet.Cells(rowEnd, 2).Value
            If Len(s) = 0 Then Exit Do
            rowEnd = rowEnd + 1
        Loop
        rowEnd = rowEnd - 1
        eWorkSheet.Rows(rowEnd).Select
        Selection.Delete
        eWorkSheet.Rows(rowEnd).Select
    End If
    Selection.Delete
    rowEnd = rowEnd - 1

    row = rowEnd
    While row >= rowBeg
        s = eWorkSheet.Cells(row, 2)
        If InStr(1, s, "Sum") > 0 Then
            eWorkSheet.Range("F" & row - 1 & ":G" & row - 1).Copy
            eWorkSheet.Range("F" & row & ":G" & row).PasteSpecial xlPasteAll
            eWorkSheet.Range("I" & row - 1 & ":K" & row - 1).Copy
            eWorkSheet.Range("I" & row & ":K" & row).PasteSpecial xlPasteAll
            eWorkSheet.Range("M" & row - 1 & ":O" & row - 1).Copy
            eWorkSheet.Range("M" & row & ":O" & row).PasteSpecial xlPasteAll
        
            eWorkSheet.Range("H" & row).Copy
            eWorkSheet.Range("H" & row).PasteSpecial xlPasteValues
            eWorkSheet.Range("L" & row).Copy
            eWorkSheet.Range("L" & row).PasteSpecial xlPasteValues
            
            eWorkSheet.Range("B" & row & ":O" & row).Font.Bold = True
            With eWorkSheet.Range("B" & row & ":E" & row)
                .Interior.Color = &HF0F0F0
            End With
            With eWorkSheet.Range("G" & row & ":J" & row)
                .Interior.Color = &HF0F0F0
            End With
            With eWorkSheet.Range("L" & row & ":M" & row)
                .Interior.Color = &HF0F0F0
            End With
            With eWorkSheet.Range("O" & row & ":M" & row)
                .Interior.Color = &HF0F0F0
            End With
        Else
            With eWorkSheet.Range("H" & row & ":K" & row)
                .ClearContents
                .ClearFormats
                .Interior.ColorIndex = 0
            End With
            With eWorkSheet.Range("M" & row & ":O" & row)
                .ClearContents
                .ClearFormats
                .Interior.ColorIndex = 0
            End With
        End If
        row = row - 1
    Wend
    eWorkSheet.Range("B" & rowBeg - 3 & ":O" & rowEnd).Select
    MakeBorders
    MakeTable = rowEnd
End Function


Private Sub MakeBorders()
Dim borderWeight As XlBorderWeight

    borderWeight = XlBorderWeight.xlMedium
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = borderWeight
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = borderWeight
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = borderWeight
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = borderWeight
    End With
End Sub

Private Sub SendReport(path As String, title As String, info As String, addr As String)
On Error GoTo errHandler
Dim oApp As Outlook.Application
Dim oMail As Outlook.MailItem

    Set oApp = CreateObject("Outlook.Application")
    Set oMail = oApp.CreateItem(olMailItem)
    oMail.To = addr
    oMail.Subject = title
    oMail.Attachments.Add path, , , title
    oMail.Body = info
    oMail.Save
    Exit Sub

errHandler:
    MsgBox Err.Description
End Sub


Private Function ReadProjectTeams() As Scripting.Dictionary
Dim sheet As Worksheet
Dim row As Long
Dim s As String
Dim dict As Scripting.Dictionary
Dim team As clsTeam

    Set sheet = sheetProjTeams
    Set dict = New Scripting.Dictionary
    Set ReadProjectTeams = dict
    row = 2
    Do
        s = sheet.Cells(row, 1)
        If Len(s) = 0 Then Exit Function
        
        Set team = New clsTeam
        team.DivisionName = s
        team.AreaName = sheet.Cells(row, 2)
        team.TeamName = sheet.Cells(row, 3)
        team.TeamLeader = sheet.Cells(row, 4)
        team.DivisionLeader = sheet.Cells(row, 5)
        team.DivisionLeaderEmail = sheet.Cells(row, 6)
        team.DivisionNameShort = sheet.Cells(row, 7)
        team.SaveEmailPath = sheet.Cells(row, 8)
        Set team.DictEmployees = ReadProjectTeamEmployees(s)
        dict.Add s, team
    
        row = row + 1
    Loop
End Function

Private Function ReadProjectTeamEmployees(ByVal DivisionName As String) As Scripting.Dictionary
Dim dict As Scripting.Dictionary
Dim sheet As Worksheet
Dim row As Long
Dim s As String
Dim r As Excel.Range

    Set dict = New Scripting.Dictionary
    Set ReadProjectTeamEmployees = dict
    Set sheet = sheetProjTeamMembers
    
    Set r = sheet.Range("$A$1:$F$585")
    r.AutoFilter
    r.AutoFilter Field:=3, Criteria1:=DivisionName
    row = 2
    Do
        If Not sheet.Rows(row).Hidden Then
            s = sheet.Cells(row, 1)
            If Len(s) = 0 Then Exit Function
            dict.Add s, Null
        End If
        row = row + 1
    Loop
End Function

Private Sub ReadProjectTeamEmployees_TotalBudget(ByVal DivisionName As String, ByVal projName As String, dt As Date, _
    ByRef dictTotalBudget As Scripting.Dictionary)
Dim sheet As Worksheet
Dim row As Long
Dim s As String
Dim r As Excel.Range
Dim v As Variant

    Set dictTotalBudget = New Scripting.Dictionary
    Set sheet = sheetMH
    
    Set r = sheet.Range("$A$1:$FS$9999")
    r.AutoFilter
    r.AutoFilter Field:=2, Criteria1:=DivisionName
    r.AutoFilter Field:=6, Criteria1:=projName
    row = 2
    Do
        If Not sheet.Rows(row).Hidden Then
            s = sheet.Cells(row, 1)
            If Len(s) = 0 Then Exit Sub
            v = sheet.Cells(row, 7)
            If IsNumeric(v) Then
                v = CDbl(v)
            Else
                v = 0
            End If
            If dictTotalBudget.Exists(s) Then
                dictTotalBudget(s) = v + dictTotalBudget(s)
            Else
                dictTotalBudget.Add s, v
            End If
        End If
        row = row + 1
    Loop
End Sub

Private Sub ReadProjectTeamEmployees_ConsumedAndTotalMH(ByVal DivisionName As String, ByVal projName As String, dtStart As Date, dtEnd As Date, _
    ByRef dictConsumed As Scripting.Dictionary, _
    ByRef dblTotalMH As Double)
Dim sheet As Worksheet
Dim row As Long
Dim s As String
Dim r As Excel.Range
Dim v As Variant
Dim dt As Date

    dblTotalMH = 0
    Set dictConsumed = New Scripting.Dictionary
    Set sheet = sheetUMH
    
    Set r = sheet.Range("$A$5:$L$999999")
    r.AutoFilter
    r.AutoFilter Field:=12, Criteria1:=DivisionName
    r.AutoFilter Field:=8, Criteria1:=projName
    row = 6
    Do
        If Not sheet.Rows(row).Hidden Then
            s = sheet.Cells(row, 5)
            If Len(s) = 0 Then Exit Sub
            dt = CDate(sheet.Cells(row, 1))
            If dt >= dtStart And dt <= dtEnd Then
                v = sheet.Cells(row, 7)
                If IsNumeric(v) Then
                    v = CDbl(v)
                Else
                    v = 0
                End If
                If dictConsumed.Exists(s) Then
                    dictConsumed(s) = v + dictConsumed(s)
                Else
                    dictConsumed.Add s, v
                End If
            End If
            If dt <= dtEnd Then
                v = sheet.Cells(row, 7)
                If IsNumeric(v) Then
                    dblTotalMH = dblTotalMH + CDbl(v)
                End If
            End If
        End If
        row = row + 1
    Loop
End Sub

Private Sub ReadProjectTeamEmployees_PlannedActComp(ByVal DivisionName As String, ByVal projName As String, dtStart As Date, dtEnd As Date, _
    ByRef dictPlannedActComp2 As Scripting.Dictionary, _
    ByRef dictPlannedActComp As Scripting.Dictionary)
Dim sheet As Worksheet
Dim row As Long
Dim s As String
Dim r As Excel.Range
Dim v As Variant
Dim dts As String
Dim dti As Long
Dim dtiStart As Long
Dim dtiEnd As Long

    dtiStart = Year(dtStart) * 100 + Month(dtStart)
    dtiEnd = Year(dtEnd) * 100 + Month(dtEnd)
    Set dictPlannedActComp2 = New Scripting.Dictionary
    Set dictPlannedActComp = New Scripting.Dictionary
    Set sheet = sheetPAC
    
    Set r = sheet.Range("$A$2:$O$999999")
    r.AutoFilter
    r.AutoFilter Field:=12, Criteria1:=DivisionName
    r.AutoFilter Field:=1, Criteria1:=projName
    row = 6
    Do
        If Not sheet.Rows(row).Hidden Then
            s = sheet.Cells(row, 4)
            If Len(s) = 0 Then Exit Sub
            dts = sheet.Cells(row, 8)
            dti = CLng(dts)
            If dti >= dtiStart Then
                v = sheet.Cells(row, 9)
                If IsNumeric(v) Then
                    v = CDbl(v)
                Else
                    v = 0
                End If
                If dictPlannedActComp.Exists(s) Then
                    dictPlannedActComp(s) = v + dictPlannedActComp(s)
                Else
                    dictPlannedActComp.Add s, v
                End If
                If dti <= dtiEnd Then
                    If dictPlannedActComp2.Exists(s) Then
                        dictPlannedActComp2(s) = v + dictPlannedActComp2(s)
                    Else
                        dictPlannedActComp2.Add s, v
                    End If
                End If
            End If
        End If
        row = row + 1
    Loop
End Sub

Private Sub ReadProjectTeamEmployees_PlannedActCompE(ByVal DivisionName As String, ByVal projName As String, dtStart As Date, dtEnd As Date, _
    ByRef dictPlannedActCompE As Scripting.Dictionary)
Dim sheet As Worksheet
Dim row As Long
Dim s As String
Dim r As Excel.Range
Dim v As Variant
Dim dts As String
Dim dti As Long
Dim dtiStart As Long
Dim dtiEnd As Long

    dtiStart = Year(dtStart) * 100 + Month(dtStart)
    dtiEnd = Year(dtEnd) * 100 + Month(dtEnd)
    Set dictPlannedActCompE = New Scripting.Dictionary
    Set sheet = sheetPACE
    
    Set r = sheet.Range("$A$2:$O$999999")
    r.AutoFilter
    r.AutoFilter Field:=12, Criteria1:=DivisionName
    r.AutoFilter Field:=1, Criteria1:=projName
    row = 6
    Do
        If Not sheet.Rows(row).Hidden Then
            s = sheet.Cells(row, 4)
            If Len(s) = 0 Then Exit Sub
            dts = sheet.Cells(row, 8)
            dti = CLng(dts) + 1
            If dti >= dtiStart Then
                v = sheet.Cells(row, 9)
                If IsNumeric(v) Then
                    v = CDbl(v)
                Else
                    v = 0
                End If
                If dictPlannedActCompE.Exists(s) Then
                    dictPlannedActCompE(s) = v + dictPlannedActCompE(s)
                Else
                    dictPlannedActCompE.Add s, v
                End If
            End If
        End If
        row = row + 1
    Loop
End Sub






Private Sub ReadProjectTeamEmployees_TotalBudget2(ByVal DivisionName As String, ByVal projName As String, dt As Date, _
    ByRef dictTotalBudget As Scripting.Dictionary)
Dim sheet As Worksheet
Dim row As Long
Dim s As String
Dim r As Excel.Range
Dim v As Variant

    Set dictTotalBudget = New Scripting.Dictionary
    Set sheet = sheetCost
    
    Set r = sheet.Range("$A$1:$FS$9999")
    r.AutoFilter
    r.AutoFilter Field:=2, Criteria1:=DivisionName
    r.AutoFilter Field:=6, Criteria1:=projName
    row = 2
    Do
        If Not sheet.Rows(row).Hidden Then
            s = sheet.Cells(row, 1)
            If Len(s) = 0 Then Exit Sub
            v = sheet.Cells(row, 7)
            If IsNumeric(v) Then
                v = CDbl(v)
            Else
                v = 0
            End If
            If dictTotalBudget.Exists(s) Then
                dictTotalBudget(s) = v + dictTotalBudget(s)
            Else
                dictTotalBudget.Add s, v
            End If
        End If
        row = row + 1
    Loop
End Sub

Private Sub ReadProjectTeamEmployees_ConsumedAndTotalMH2(ByVal DivisionName As String, ByVal projName As String, dtStart As Date, dtEnd As Date, _
    ByRef dictConsumed As Scripting.Dictionary, _
    ByRef dblTotalMH As Double)
Dim sheet As Worksheet
Dim row As Long
Dim s As String
Dim r As Excel.Range
Dim v As Variant
Dim dt As Date

    dblTotalMH = 0
    Set dictConsumed = New Scripting.Dictionary
    Set sheet = sheetUMH
    
    Set r = sheet.Range("$A$5:$L$999999")
    r.AutoFilter
    r.AutoFilter Field:=12, Criteria1:=DivisionName
    r.AutoFilter Field:=8, Criteria1:=projName
    row = 6
    Do
        If Not sheet.Rows(row).Hidden Then
            s = sheet.Cells(row, 5)
            If Len(s) = 0 Then Exit Sub
            dt = CDate(sheet.Cells(row, 1))
            If dt >= dtStart And dt <= dtEnd Then
                v = sheet.Cells(row, 10)
                If IsNumeric(v) Then
                    v = CDbl(v)
                Else
                    v = 0
                End If
                If dictConsumed.Exists(s) Then
                    dictConsumed(s) = v + dictConsumed(s)
                Else
                    dictConsumed.Add s, v
                End If
            End If
            If dt <= dtEnd Then
                v = sheet.Cells(row, 10)
                If IsNumeric(v) Then
                    dblTotalMH = dblTotalMH + CDbl(v)
                End If
            End If
        End If
        row = row + 1
    Loop
End Sub

Private Sub ReadProjectTeamEmployees_PlannedActComp2(ByVal DivisionName As String, ByVal projName As String, dtStart As Date, dtEnd As Date, _
    ByRef dictPlannedActComp2 As Scripting.Dictionary, _
    ByRef dictPlannedActComp As Scripting.Dictionary)
Dim sheet As Worksheet
Dim row As Long
Dim s As String
Dim r As Excel.Range
Dim v As Variant
Dim dts As String
Dim dti As Long
Dim dtiStart As Long
Dim dtiEnd As Long

    dtiStart = Year(dtStart) * 100 + Month(dtStart)
    dtiEnd = Year(dtEnd) * 100 + Month(dtEnd)
    Set dictPlannedActComp2 = New Scripting.Dictionary
    Set dictPlannedActComp = New Scripting.Dictionary
    Set sheet = sheetPAC
    
    Set r = sheet.Range("$A$2:$O$999999")
    r.AutoFilter
    r.AutoFilter Field:=12, Criteria1:=DivisionName
    r.AutoFilter Field:=1, Criteria1:=projName
    row = 6
    Do
        If Not sheet.Rows(row).Hidden Then
            s = sheet.Cells(row, 4)
            If Len(s) = 0 Then Exit Sub
            dts = sheet.Cells(row, 8)
            dti = CLng(dts)
            If dti >= dtiStart Then
                v = sheet.Cells(row, 14)
                If IsNumeric(v) Then
                    v = CDbl(v)
                Else
                    v = 0
                End If
                If dictPlannedActComp.Exists(s) Then
                    dictPlannedActComp(s) = v + dictPlannedActComp(s)
                Else
                    dictPlannedActComp.Add s, v
                End If
                If dti <= dtiEnd Then
                    If dictPlannedActComp2.Exists(s) Then
                        dictPlannedActComp2(s) = v + dictPlannedActComp2(s)
                    Else
                        dictPlannedActComp2.Add s, v
                    End If
                End If
            End If
        End If
        row = row + 1
    Loop
End Sub

Private Sub ReadProjectTeamEmployees_PlannedActCompE2(ByVal DivisionName As String, ByVal projName As String, dtStart As Date, dtEnd As Date, _
    ByRef dictPlannedActCompE As Scripting.Dictionary)
Dim sheet As Worksheet
Dim row As Long
Dim s As String
Dim r As Excel.Range
Dim v As Variant
Dim dts As String
Dim dti As Long
Dim dtiStart As Long
Dim dtiEnd As Long

    dtiStart = Year(dtStart) * 100 + Month(dtStart)
    dtiEnd = Year(dtEnd) * 100 + Month(dtEnd)
    Set dictPlannedActCompE = New Scripting.Dictionary
    Set sheet = sheetPACE
    
    Set r = sheet.Range("$A$2:$O$999999")
    r.AutoFilter
    r.AutoFilter Field:=12, Criteria1:=DivisionName
    r.AutoFilter Field:=1, Criteria1:=projName
    row = 6
    Do
        If Not sheet.Rows(row).Hidden Then
            s = sheet.Cells(row, 4)
            If Len(s) = 0 Then Exit Sub
            dts = sheet.Cells(row, 8)
            dti = CLng(dts) + 1
            If dti >= dtiStart Then
                v = sheet.Cells(row, 14)
                If IsNumeric(v) Then
                    v = CDbl(v)
                Else
                    v = 0
                End If
                If dictPlannedActCompE.Exists(s) Then
                    dictPlannedActCompE(s) = v + dictPlannedActCompE(s)
                Else
                    dictPlannedActCompE.Add s, v
                End If
            End If
        End If
        row = row + 1
    Loop
End Sub




Private Function ReadProjects() As Scripting.Dictionary
Dim sheet As Worksheet
Dim row As Long
Dim s As String
Dim dict As Scripting.Dictionary

    Set sheet = sheetProjects
    Set dict = New Scripting.Dictionary
    Set ReadProjects = dict
    row = 2
    Do
        s = sheet.Cells(row, 1)
        If Len(s) = 0 Then Exit Function
        dict.Add s, Null
        row = row + 1
    Loop
End Function

