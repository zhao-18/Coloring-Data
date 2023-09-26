Attribute VB_Name = "Module1"
Sub 指定データの上下5Pと10P網掛け()
Dim strUserName As String


strUserName = Application.UserName
WorkbookName = ActiveWorkbook.Name

'ColorIndexP10 = Cells(3, 5)   '★10％以上色指定箇所
Set col10 = Application.InputBox("10%以上の色があるセルを選択してください。", "色選択 ", Type:=8)
    col10.Select
    ColorIndexP10 = Selection.Interior.Color
    
'ColorIndexP5 = Cells(4, 5)    '★5％以上色指定箇所
Set col5 = Application.InputBox("5%以上の色があるセルを選択してください。", "色選択 ", Type:=8)
    col5.Select
    ColorIndexP5 = Selection.Interior.Color
    
'ColorIndexM10 = Cells(5, 5)   '★-10％以上色指定箇所
Set col10_2 = Application.InputBox("10%以下の色があるセルを選択してください。", "色選択 ", Type:=8)
    col10_2.Select
    ColorIndexM10 = Selection.Interior.Color
    
'ColorIndexM5 = Cells(6, 5)    '★-5％以上色指定箇所
Set col5_2 = Application.InputBox("5%以下の色があるセルを選択してください。", "色選択 ", Type:=8)
    col5_2.Select
    ColorIndexM5 = Selection.Interior.Color
    
Windows("★[列用]指定データと比較し上下5%、10%のデータを網掛けする.xlsm").Activate

'■■ポイント差の設定----------
'ここの数字を変更することで、ポイント差の設定ができます。

'point1 = Cells(3, 7)          '■5は5ポイント差
point1 = 5

'point2 = Cells(4, 7)          '■10は10ポイント差
point2 = 10

'--ポイント差の設定----------

With Cells(3, 2).Interior
    .Color = ColorIndexP10
    .Pattern = xlSolid
End With

With Cells(4, 2).Interior
    .Color = ColorIndexP5
    .Pattern = xlSolid
End With

With Cells(5, 2).Interior
    .Color = ColorIndexM10
    .Pattern = xlSolid
End With

With Cells(6, 2).Interior
    .Color = ColorIndexM5
    .Pattern = xlSolid
End With

Windows(WorkbookName).Activate

Set b = Application.InputBox("比較の基準となるデータを選択してください。", "列に網掛け ", Type:=8)
    b.Select
    基準列 = Selection.Column
    
Set a = Application.InputBox("網掛けしたい部分のデータを選択してください。", "列に網掛け ", Type:=8)
    a.Select
    上 = Selection.Row
    下 = 上 + Selection.Rows.Count - 1
    左 = Selection.Column
    右 = 左 + Selection.Columns.Count - 1
    Selection.Interior.ColorIndex = xlNone
    Selection.Font.ColorIndex = 0
    Selection.FormatConditions.Delete
    
For m = 左 To 右
    For n = 上 To 下
        基準値 = Cells(n, 基準列)
        比較値 = Cells(n, m)
        'サンプル数 = Cells(n, 左 - 1)
        
        'If サンプル数 < 30 Then GoTo 100
        If Not IsNumeric(比較値) Or Not IsNumeric(基準値) Then GoTo 100
        
        If 比較値 >= 基準値 + point2 Then
            Cells(n, m).Select
                With Selection.Interior
                    .Color = ColorIndexP10
                    .Pattern = xlSolid
                End With
            GoTo 100
        End If
        
        If 比較値 >= 基準値 + point1 Then
            Cells(n, m).Select
            With Selection.Interior
                .Color = ColorIndexP5
                .Pattern = xlSolid
            End With
            GoTo 100
        End If
        
        If 基準値 - point2 >= 比較値 Then
            Cells(n, m).Select
            With Selection.Interior
                .Color = ColorIndexM10
                .Pattern = xlSolid
            End With
            GoTo 100
        End If
        
        If 基準値 - point1 >= 比較値 Then
            Cells(n, m).Select
            With Selection.Interior
                .Color = ColorIndexM5
                .Pattern = xlSolid
            End With
        End If
        
100

    Next
Next

MsgBox strUserName & "さん、ざっとでいいから、データをチェックしてね、ざっと！", , "ご挨拶"
End Sub
