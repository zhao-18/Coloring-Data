Attribute VB_Name = "Module1"
Sub �w��f�[�^�̏㉺5P��10P�Ԋ|��()
Dim strUserName As String


strUserName = Application.UserName
WorkbookName = ActiveWorkbook.Name

'ColorIndexP10 = Cells(3, 5)   '��10���ȏ�F�w��ӏ�
Set col10 = Application.InputBox("10%�ȏ�̐F������Z����I�����Ă��������B", "�F�I�� ", Type:=8)
    col10.Select
    ColorIndexP10 = Selection.Interior.Color
    
'ColorIndexP5 = Cells(4, 5)    '��5���ȏ�F�w��ӏ�
Set col5 = Application.InputBox("5%�ȏ�̐F������Z����I�����Ă��������B", "�F�I�� ", Type:=8)
    col5.Select
    ColorIndexP5 = Selection.Interior.Color
    
'ColorIndexM10 = Cells(5, 5)   '��-10���ȏ�F�w��ӏ�
Set col10_2 = Application.InputBox("10%�ȉ��̐F������Z����I�����Ă��������B", "�F�I�� ", Type:=8)
    col10_2.Select
    ColorIndexM10 = Selection.Interior.Color
    
'ColorIndexM5 = Cells(6, 5)    '��-5���ȏ�F�w��ӏ�
Set col5_2 = Application.InputBox("5%�ȉ��̐F������Z����I�����Ă��������B", "�F�I�� ", Type:=8)
    col5_2.Select
    ColorIndexM5 = Selection.Interior.Color
    
Windows("��[��p]�w��f�[�^�Ɣ�r���㉺5%�A10%�̃f�[�^��Ԋ|������.xlsm").Activate

'�����|�C���g���̐ݒ�----------
'�����̐�����ύX���邱�ƂŁA�|�C���g���̐ݒ肪�ł��܂��B

'point1 = Cells(3, 7)          '��5��5�|�C���g��
point1 = 5

'point2 = Cells(4, 7)          '��10��10�|�C���g��
point2 = 10

'--�|�C���g���̐ݒ�----------

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

Set b = Application.InputBox("��r�̊�ƂȂ�f�[�^��I�����Ă��������B", "��ɖԊ|�� ", Type:=8)
    b.Select
    ��� = Selection.Column
    
Set a = Application.InputBox("�Ԋ|�������������̃f�[�^��I�����Ă��������B", "��ɖԊ|�� ", Type:=8)
    a.Select
    �� = Selection.Row
    �� = �� + Selection.Rows.Count - 1
    �� = Selection.Column
    �E = �� + Selection.Columns.Count - 1
    Selection.Interior.ColorIndex = xlNone
    Selection.Font.ColorIndex = 0
    Selection.FormatConditions.Delete
    
For m = �� To �E
    For n = �� To ��
        ��l = Cells(n, ���)
        ��r�l = Cells(n, m)
        '�T���v���� = Cells(n, �� - 1)
        
        'If �T���v���� < 30 Then GoTo 100
        If Not IsNumeric(��r�l) Or Not IsNumeric(��l) Then GoTo 100
        
        If ��r�l >= ��l + point2 Then
            Cells(n, m).Select
                With Selection.Interior
                    .Color = ColorIndexP10
                    .Pattern = xlSolid
                End With
            GoTo 100
        End If
        
        If ��r�l >= ��l + point1 Then
            Cells(n, m).Select
            With Selection.Interior
                .Color = ColorIndexP5
                .Pattern = xlSolid
            End With
            GoTo 100
        End If
        
        If ��l - point2 >= ��r�l Then
            Cells(n, m).Select
            With Selection.Interior
                .Color = ColorIndexM10
                .Pattern = xlSolid
            End With
            GoTo 100
        End If
        
        If ��l - point1 >= ��r�l Then
            Cells(n, m).Select
            With Selection.Interior
                .Color = ColorIndexM5
                .Pattern = xlSolid
            End With
        End If
        
100

    Next
Next

MsgBox strUserName & "����A�����Ƃł�������A�f�[�^���`�F�b�N���ĂˁA�����ƁI", , "�����A"
End Sub
