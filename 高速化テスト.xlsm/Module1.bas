Attribute VB_Name = "Module1"
Option Explicit

Sub for�v��()

Dim i As Long
Dim j As Long

Application.ScreenUpdating = False

Debug.Print Time & "-�v���J�n"
    For j = 2 To Worksheets(esh.���o).Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To Worksheets(esh.�}�X�^).Cells(Rows.Count, 1).End(xlUp).Row
            If Worksheets(esh.�}�X�^).Cells(i, 1) = Worksheets(esh.���o).Cells(j, 1) Then
                Worksheets(esh.���o).Cells(j, 2) = Worksheets(esh.�}�X�^).Cells(i, 2)
                Exit For
            End If
        Next
    Next
Application.ScreenUpdating = True

Debug.Print Time & "-�v���I��"
End Sub

Sub �z��_vlookup()
Dim i As Long, j As Long
Dim B As Variant
Application.ScreenUpdating = False
Debug.Print Time & "-�v���J�n"
ReDim B(9999, 0)

For i = 2 To Worksheets(esh.���o).Cells(Rows.Count, 1).End(xlUp).Row
    On Error GoTo Errhandl
      B(i - 2, 0) = WorksheetFunction.VLookup(Worksheets(esh.���o).Cells(i, 1), Worksheets(esh.�}�X�^).Range(Worksheets(esh.�}�X�^).Cells(2, 1), Worksheets(esh.�}�X�^).Cells(Rows.Count, 2).End(xlUp)), 2, False)
      
Next
Worksheets(esh.���o).Range("b2:b10001") = B
Application.ScreenUpdating = True

Errhandl:
B(i - 2, 0) = "�Y���Ȃ�"
Err.Clear
Resume Next
 
Debug.Print Time & "-�v���I��"
End Sub



Sub �z��_vlookup_����()

'�z��v���O����
    Dim i As Long, j As Long
    Dim n As Long
    Dim B As Variant
    Application.ScreenUpdating = False
    n = Worksheets(esh.���o).Cells(Rows.Count, 1).End(xlUp).Row
    Debug.Print Time & "-�v���J�n"
    ReDim B(n - 2, 0)
    
    For i = 2 To n
        On Error GoTo Errhandl
          B(i - 2, 0) = WorksheetFunction.VLookup(Worksheets(esh.���o).Cells(i, 1), _
          Worksheets(esh.�}�X�^).Range(Worksheets(esh.�}�X�^).Cells(2, 1), Worksheets(esh.�}�X�^).Cells(Rows.Count, 2).End(xlUp)), 2, False)
          
    Next
    Worksheets("���o").Range(Worksheets("���o").Range("A2"), _
    Worksheets("���o").Cells(Rows.Count, 1).End(xlUp)).Offset(0, 1) = B
    
    Application.ScreenUpdating = True
    Debug.Print Time & "-�v���I��"
    Exit Sub
    
Errhandl:
    B(i - 2, 0) = "�Y���Ȃ�"
    Err.Clear
    Resume Next
     
    
End Sub

