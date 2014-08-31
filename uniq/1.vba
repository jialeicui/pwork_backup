Option Explicit
Dim g_flag As String
Sub fix_center()
    Dim row As Long
    Dim col As Integer
    Dim center As String
    Dim pos As String
    For row = 2 To 96782
        pos = "EE2"
        center = Range("EE" & row).Value
        If Len(center) > 4 Then
            Range("EE" & row) = Left(center, 4)
        End If
    Next
End Sub
'判断一个范围是否有选择
Function range_ok(r As Range, Optional full_wrong As Boolean = False) As Boolean
    Dim c As Range
    '如果不是全选算错
    If Not full_wrong Then
        For Each c In r
            If c.Value <> "" Then
                range_ok = True
                Exit Function
            End If
        Next
    Else
        Dim emp As Boolean
        Dim have As Boolean
        emp = False
        have = False
        For Each c In r
            If c.Value <> "" Then
                have = True
            Else
                emp = True
            End If
        Next
        '如果既有空的也有填的,说明ok
        If emp And have Then
            range_ok = True
            Exit Function
        End If
    End If
    '如果有错误,全都替换为特定标志
    For Each c In r
        c.Value = g_flag
    Next
    r.Interior.ColorIndex = 42
    range_ok = False
End Function

Function is_ok(row As Long) As Boolean
    Dim rs As Variant
    Dim r As Variant
    Dim res_col As Integer
    res_col = Range("EO1").Column
    rs = Array("P", "Q", "R,W", "X,AL", "AM,AU", "AV,BO", "BP,CX", "CY,DH")
    
    is_ok = True
    '身份证
    If Len(Range("J" & row)) < 15 Then
        is_ok = False
        Range("J" & row) = g_flag
    End If
    res_col = res_col + 1
    For Each r In rs
        Dim p As Integer
        Dim ss As String
        Dim testr As Range
        ss = CStr(r)
        p = InStr(ss, ",")
        If p = 0 Then
            Set testr = Range(ss & row)
        Else
            Set testr = Range(Left(ss, p - 1) & row, Right(ss, Len(ss) - p) & row)
        End If
        
        If Not range_ok(testr) Then
            is_ok = False
        End If
    Next
    '血压
    Dim sbp As Long
    Dim dbp As Long
    sbp = Range("DI" & row).Value
    dbp = Range("DJ" & row).Value
    If sbp Mod 10 = 0 And dbp Mod 10 = 0 Then
        Range("DI" & row, "DJ" & row).Interior.ColorIndex = 42
        is_ok = False
        Range("DI" & row).Value = g_flag
        Range("DJ" & row).Value = g_flag
    End If
    
End Function

Sub proc()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    g_flag = "missing"
    Dim row As Long
    For row = 2 To 100000
        If Range("A" & row) = "" Then
            Exit Sub
        End If
        If is_ok(row) Then
        End If
    Next
End Sub
