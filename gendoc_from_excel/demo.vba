Option Explicit

Sub CreateWord()
    Dim mypath, Newname, i, XB, wApp, valMap, row, dSheet
    mypath = ThisWorkbook.Path & "\"
    valMap = Split("B,B,G,M,G,H,K,L,Q,R,U,V,W,X,AA,AB,AO,AP,AS,AT,AU,AV,AY,AZ,BB,BJ,BM,BN,BP,BR,BT,BU,BV,BW", ",")
    dSheet = "关键指标汇总"
    
    For row = 4 To 36
        Dim pro As String
        pro = ThisWorkbook.Worksheets(dSheet).Cells(row, 2)
        Newname = "PRE" & pro & "EXT.doc"
        FileCopy mypath & "模板.doc", mypath & Newname 
        Set wApp = CreateObject("word.application")
        With wApp
            .Visible = True
            Dim curDoc
            curDoc = .Documents.Open(mypath & Newname)

            
            For i = 0 To UBound(valMap)
                Do While .Selection.Find.Execute("JJ" & i & "J")
                    .Selection.Text = Trim(ThisWorkbook.Worksheets(dSheet).Range(valMap(i) & row).Text())
                    .Selection.HomeKey Unit:=6
                Loop
            Next
            .Documents.Save
            .Quit
    
        End With
    
        Set wApp = Nothing
    Next
    
End Sub