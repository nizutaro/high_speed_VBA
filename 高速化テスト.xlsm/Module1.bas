Attribute VB_Name = "Module1"
Option Explicit

Sub for計測()

Dim i As Long
Dim j As Long

Application.ScreenUpdating = False

Debug.Print Time & "-計測開始"
    For j = 2 To Worksheets(esh.抽出).Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To Worksheets(esh.マスタ).Cells(Rows.Count, 1).End(xlUp).Row
            If Worksheets(esh.マスタ).Cells(i, 1) = Worksheets(esh.抽出).Cells(j, 1) Then
                Worksheets(esh.抽出).Cells(j, 2) = Worksheets(esh.マスタ).Cells(i, 2)
                Exit For
            End If
        Next
    Next
Application.ScreenUpdating = True

Debug.Print Time & "-計測終了"
End Sub

Sub 配列_vlookup()
Dim i As Long, j As Long
Dim B As Variant
Application.ScreenUpdating = False
Debug.Print Time & "-計測開始"
ReDim B(9999, 0)

For i = 2 To Worksheets(esh.抽出).Cells(Rows.Count, 1).End(xlUp).Row
    On Error GoTo Errhandl
      B(i - 2, 0) = WorksheetFunction.VLookup(Worksheets(esh.抽出).Cells(i, 1), Worksheets(esh.マスタ).Range(Worksheets(esh.マスタ).Cells(2, 1), Worksheets(esh.マスタ).Cells(Rows.Count, 2).End(xlUp)), 2, False)
      
Next
Worksheets(esh.抽出).Range("b2:b10001") = B
Application.ScreenUpdating = True

Errhandl:
B(i - 2, 0) = "該当なし"
Err.Clear
Resume Next
 
Debug.Print Time & "-計測終了"
End Sub



Sub 配列_vlookup_改良()

'配列プログラム
    Dim i As Long, j As Long
    Dim n As Long
    Dim B As Variant
    Application.ScreenUpdating = False
    n = Worksheets(esh.抽出).Cells(Rows.Count, 1).End(xlUp).Row
    Debug.Print Time & "-計測開始"
    ReDim B(n - 2, 0)
    
    For i = 2 To n
        On Error GoTo Errhandl
          B(i - 2, 0) = WorksheetFunction.VLookup(Worksheets(esh.抽出).Cells(i, 1), _
          Worksheets(esh.マスタ).Range(Worksheets(esh.マスタ).Cells(2, 1), Worksheets(esh.マスタ).Cells(Rows.Count, 2).End(xlUp)), 2, False)
          
    Next
    Worksheets("抽出").Range(Worksheets("抽出").Range("A2"), _
    Worksheets("抽出").Cells(Rows.Count, 1).End(xlUp)).Offset(0, 1) = B
    
    Application.ScreenUpdating = True
    Debug.Print Time & "-計測終了"
    Exit Sub
    
Errhandl:
    B(i - 2, 0) = "該当なし"
    Err.Clear
    Resume Next
     
    
End Sub

