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
Function range_ok(r As Range, Optional judge As String = "opt", Optional full_wrong As Boolean = False) As Boolean
    Dim c As Range
    If judge = "opt" Or judge = "1" Then
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
        range_ok = False
    Else
        If judge = "2" Then
            range_ok = True
        Else
            range_ok = False
        End If
    End If
    If Not range_ok Then
        '如果有错误,全都替换为特定标志
        For Each c In r
            c.Value = g_flag
        Next
        r.Interior.ColorIndex = 42
    End If
End Function

Function is_ok(row As Long) As Boolean
    Dim rs As Variant
    Dim r As Variant

    rs = Array("R", "S", "AO,AW", "AX:AY,BM", "BS:BN,BR", "DM:BT,DB", "EB:DC,DL")
    
    is_ok = True
    '身份证
    If Len(Range("M" & row)) < 15 Then
        is_ok = False
        Range("M" & row) = g_flag
    End If
    For Each r In rs
        Dim p As Integer
        Dim ss As String
        Dim opt As String
        Dim testr As Range
        Dim is_opt As Boolean
        is_opt = False
        ss = CStr(r)
        p = InStr(ss, ":")
        If p <> 0 Then
            is_opt = True
            opt = Left(ss, p - 1)
            ss = Right(ss, Len(ss) - p)
        End If
        p = InStr(ss, ",")
        If p = 0 Then
            Set testr = Range(ss & row)
        Else
            Set testr = Range(Left(ss, p - 1) & row, Right(ss, Len(ss) - p) & row)
        End If
        
        Dim ok As Boolean
        If is_opt Then
            ok = range_ok(testr, Range(opt & row))
        Else
            ok = range_ok(testr)
        End If
        If Not ok Then
            is_ok = False
            '有前置判断
            If is_opt Then
                Range(opt & row).Value = g_flag
            End If
        End If

    Next
    
    If Not range_ok(Range("T" & row, "Y" & row), "opt", True) Then
        is_ok = False
    End If
    If Not range_ok(Range("Z" & row, "AN" & row), "opt", True) Then
        is_ok = False
    End If
    '血压
    Dim sbp As Long
    Dim dbp As Long
    sbp = Range("DN" & row).Value
    dbp = Range("DO" & row).Value
    If sbp Mod 10 = 0 And dbp Mod 10 = 0 Then
        Range("DN" & row, "DO" & row).Interior.ColorIndex = 42
        is_ok = False
        Range("DN" & row).Value = g_flag
        Range("DO" & row).Value = g_flag
    End If
    
End Function

Sub proc()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    g_flag = "missing"
    Dim row As Long
    For row = 2 To 100000
        If Range("A" & row).Value = "" Then
            Exit Sub
        End If
        If is_ok(row) Then
            Range("FT" & row) = "合格"
        End If
    Next
End Sub

