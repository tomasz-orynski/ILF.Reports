Attribute VB_Name = "modCommons"
Option Explicit
Option Base 1

Public Function GetItem(ByVal dict As Scripting.Dictionary, ByVal key As String) As Scripting.Dictionary
    If dict.Exists(key) Then
        Set GetItem = dict.Item(key)
    Else
        Set GetItem = New Scripting.Dictionary
    End If
End Function

Public Sub QuickSort(arr)
    QuickSortAsc arr, LBound(arr), UBound(arr)
End Sub

Private Sub QuickSortAsc(arr, Lo As Long, Hi As Long)
  Dim varPivot As Variant
  Dim varTmp As Variant
  Dim tmpLow As Long
  Dim tmpHi As Long
  tmpLow = Lo
  tmpHi = Hi
  varPivot = arr((Lo + Hi) \ 2)
  Do While tmpLow <= tmpHi
    Do While arr(tmpLow) < varPivot And tmpLow < Hi
      tmpLow = tmpLow + 1
    Loop
    Do While varPivot < arr(tmpHi) And tmpHi > Lo
      tmpHi = tmpHi - 1
    Loop
    If tmpLow <= tmpHi Then
      varTmp = arr(tmpLow)
      arr(tmpLow) = arr(tmpHi)
      arr(tmpHi) = varTmp
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
    End If
  Loop
  If Lo < tmpHi Then QuickSortAsc arr, Lo, tmpHi
  If tmpLow < Hi Then QuickSortAsc arr, tmpLow, Hi
End Sub

