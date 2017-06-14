Attribute VB_Name = "mod_WorkCard4EmpCons"
Option Explicit
Option Base 1

Public Sub GenerateWorkCards4EmpConst()
On Error GoTo errHandler
Dim dict As Scripting.Dictionary
Dim i As Long
Dim row As Long
Dim fullName As String
Dim position As String
Dim dtYear As Long
Dim dtMonth As Long

    Set dict = ReadData(dtYear, dtMonth)
    
    row = 2
    Do
        fullName = Trim$(sheet_EmpCons.Cells(row, 2))
        If Len(fullName) = 0 Then Exit Do
        position = Trim$(sheet_EmpCons.Cells(row, 1))
        GenerateWorkCard fullName, position, dtYear, dtMonth, mod_Commons.GetItem(dict, fullName)
        row = row + 1
    Loop
    Exit Sub
    
errHandler:
    MsgBox Err.Description
End Sub

Private Sub GenerateWorkCard(ByVal fullName As String, ByVal position As String, ByVal dtYear As Long, ByVal dtMonth As Long, ByVal dictPerson As Scripting.Dictionary)
Dim path As String
Dim eWorkbook As Excel.Workbook
Dim eWorkSheet As Excel.Worksheet
Dim wProps As Object
Dim dt As Date
Dim i As Long
Dim j As Long
Dim keys
Dim row As Long
Dim rowOff As Long
Dim s As String
Dim cnt As Long
Dim cntRow As Long

    path = Workbook.path & "\szablony\wykaz_czynnosci.xltx"
    Set eWorkbook = Application.Workbooks.Open(path)
    eWorkbook.Activate
    Set eWorkSheet = eWorkbook.Sheets.Item(1)
    eWorkSheet.Activate
    eWorkSheet.Range("C8").Value = fullName
    eWorkSheet.Range("G8").Value = position
    eWorkSheet.Range("C9").Value = dtMonth
    eWorkSheet.Range("G9").Value2 = dtYear
    dt = DateSerial(dtYear, dtMonth, 1)
    j = CInt(Format(dt, "ww", vbMonday, vbFirstFourDays))
    eWorkSheet.Range("A11").Value2 = j
    eWorkSheet.Range("A18").Value2 = j + 1
    eWorkSheet.Range("A25").Value2 = j + 2
    eWorkSheet.Range("A32").Value2 = j + 3
    eWorkSheet.Range("A39").Value2 = j + 4
    i = Weekday(dt, vbMonday) - 1
    rowOff = 10 + i
    dt = DateAdd("d", -i, dt)
    eWorkSheet.Range("A11").Value2 = j
    eWorkSheet.Range("B11").Value2 = dt
    eWorkSheet.Calculate
    
    cntRow = 46
    row = 11
    While row < cntRow
        dt = CDate(eWorkSheet.Cells(row, 2))
        If month(dt) <> dtMonth Then
            With eWorkSheet.Rows(row)
                '.RowHeight = 0
                .EntireRow.Hidden = True
            End With
        End If
        row = row + 1
    Wend
    
    keys = dictPerson.keys
    cnt = 0
    For i = LBound(keys) To UBound(keys)
        dt = CDate(keys(i))
        s = dictPerson.Item(keys(i))
        row = Day(dt) + rowOff
        eWorkSheet.Range("F" & row).Value2 = s
        'eWorkSheet.Range("F" & row & ":G" & row).AutoFit
        eWorkSheet.Range("E" & row).Value2 = GetWorkPlace(s)
        eWorkSheet.Range("D" & row).Value2 = 1
        cnt = cnt + 1
    Next
    eWorkSheet.Range("C" & cntRow).Value2 = cnt
    
    path = Workbook.path & "\dokumenty\[" & fullName & "].xlsx"
    eWorkbook.SaveCopyAs Filename:=path
    eWorkbook.Close False
End Sub

Private Function GetWorkPlace(s As String) As String
Const p1 = "Poznañ"
Const p2 = "Warszawa"
Const p3 = "Bonikowo"
Dim r As String

    If InStr(1, s, p1) Then r = r & " " & p1
    If InStr(1, s, p2) Then r = r & " " & p2
    If InStr(1, s, p3) Then r = r & " " & p3
    GetWorkPlace = Trim$(r)
End Function


Private Function ReadData(ByRef dtYear As Long, ByRef dtMonth As Long) As Scripting.Dictionary
Dim dict As Scripting.Dictionary
Dim dictPerson As Scripting.Dictionary
Dim dictPersonDay As Scripting.Dictionary
Dim row As Long
Dim fullName As String
Dim dt As Date
Dim descr As String

    dt = CDate(sheet_TS.Cells(2, 2))
    dtYear = year(dt)
    dtMonth = month(dt)
    Set dict = New Scripting.Dictionary
    row = 7
    Do
        fullName = sheet_TS.Cells(row, 3)
        If Len(fullName) = 0 Then Exit Do
        If dict.Exists(fullName) Then
            Set dictPerson = dict.Item(fullName)
        Else
            Set dictPerson = New Scripting.Dictionary
            dict.Add fullName, dictPerson
        End If
        dt = CDate(sheet_TS.Cells(row, 1))
        descr = CStr(sheet_TS.Cells(row, 6))
        If dictPerson.Exists(dt) Then
            descr = dictPerson.Item(dt) & vbCrLf & descr
            dictPerson.Item(dt) = descr
        Else
            dictPerson.Add dt, descr
        End If
        row = row + 1
    Loop

    Set ReadData = dict
End Function

