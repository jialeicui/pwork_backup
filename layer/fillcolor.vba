Option Explicit
Dim colortable(5) As Integer
Sub init_color_table()
    colortable(0) = 4
    colortable(1) = 6
    colortable(2) = 45
    colortable(3) = 3
    colortable(4) = 9
End Sub

Sub color()
    Dim row As Integer
    Dim col As Integer
    Dim i As Integer
    Dim c As Range
    
    Call init_color_table
    
    For row = 1 To 100
        For col = 1 To 50
            Set c = Cells(row, col)
            If c <> "" And c >= 0 And c <= 4 Then
                c.Interior.ColorIndex = colortable(c.Value)
            End If
        Next
    Next
End Sub

Sub rollback()
    Range("A1", "ZZ100").Interior.ColorIndex = 0
End Sub
