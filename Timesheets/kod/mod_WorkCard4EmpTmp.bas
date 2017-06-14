Attribute VB_Name = "mod_WorkCard4EmpTmp"
Option Explicit
Option Base 1

Public Sub GenerateWorkCards4EmpTmp()
On Error GoTo errHandler
Dim dict As Scripting.Dictionary
Dim i As Long
Dim wApp As Word.Application
Dim row As Long
Dim fullName As String
Dim cardNumber As String
Dim position As String
Dim dt As String


    Set wApp = CreateObject("Word.Application")
    wApp.Visible = True
    Set dict = ReadData(dt)
    
    row = 2
    Do
        fullName = Trim$(sheet_EmpTmp.Cells(row, 2))
        If Len(fullName) = 0 Then Exit Do
        position = Trim$(sheet_EmpTmp.Cells(row, 1))
        cardNumber = Trim$(sheet_EmpTmp.Cells(row, 3)) & dt
        GenerateWorkCard wApp, fullName, position, cardNumber, mod_Commons.GetItem(dict, fullName)
        row = row + 1
    Loop
   
    wApp.Quit
    Exit Sub
    
errHandler:
    MsgBox Err.Description
End Sub

Private Sub GenerateWorkCard(ByVal wApp As Word.Application, ByVal fullName As String, ByVal position As String, ByVal cardNumber As String, ByVal dictPerson As Scripting.Dictionary)
Dim path As String
Dim wDoc As Word.Document
Dim wProps As Object
Dim daysCount As Long

    path = Workbook.path & "\szablony\karta_pracy.dotx"
    Set wDoc = wApp.Documents.Open(path)
    Set wProps = wDoc.CustomDocumentProperties
    daysCount = GenerateWorkCardTable(wDoc, fullName, dictPerson)
    wProps.Item("_fullName_").Value = fullName
    wProps.Item("_position_").Value = position
    wProps.Item("_cardNumber_").Value = cardNumber
    wProps.Item("_daysCount_").Value = CStr(daysCount)
    mod_Commons.UpdateFields wDoc
    path = Workbook.path & "\dokumenty\[" & fullName & "].docx"
    wDoc.SaveAs2 Filename:=path, FileFormat:=wdFormatXMLDocument
    wDoc.Close
End Sub

Private Function GenerateWorkCardTable(wDoc As Word.Document, ByVal fullName As String, ByVal dictPerson As Scripting.Dictionary) As Long
Dim wTbl As Word.Table
Dim wTblRowSrc As Word.row
Dim wTblRowDst As Word.row
Dim keys
Dim i As Long

    keys = dictPerson.keys
    mod_Commons.QuickSort keys
    Set wTbl = wDoc.Tables.Item(1)
    Set wTblRowSrc = wTbl.Rows(2)
    For i = LBound(keys) To UBound(keys)
        wTblRowSrc.Select
        wDoc.Application.Selection.InsertRowsAbove
        Set wTblRowDst = wTbl.Rows(wTbl.Rows.Count - 2)
        wTblRowDst.Cells(1).Range.Text = (i + 1) & "."
        wTblRowDst.Cells(2).Range.Text = keys(i)
        wTblRowDst.Cells(3).Range.Text = dictPerson.Item(keys(i))
    Next
    wTblRowSrc.Delete
    GenerateWorkCardTable = i
End Function

Private Function ReadData(ByRef dateOut As String) As Scripting.Dictionary
Dim dict As Scripting.Dictionary
Dim dictPerson As Scripting.Dictionary
Dim dictPersonDay As Scripting.Dictionary
Dim row As Long
Dim fullName As String
Dim dt As Date
Dim descr As String

    dt = CDate(sheet_TS.Cells(2, 2))
    dateOut = Format$(dt, "MM") & "/" & Format$(dt, "yyyy")
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
