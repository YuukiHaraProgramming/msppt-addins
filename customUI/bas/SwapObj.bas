Attribute VB_Name = "SwapObj"

Sub SwapObj()
    Dim selectedShapes As ShapeRange
    Dim positions() As Variant
    Dim i As Integer
    
    ' �I�������I�u�W�F�N�g���擾����
    Set selectedShapes = ActiveWindow.Selection.ShapeRange
    
    ' �I���I�u�W�F�N�g��2�ȏ�Ȃ��ꍇ�͏I��
    If selectedShapes.Count < 2 Then
        MsgBox "���Ȃ��Ƃ�2�̃I�u�W�F�N�g��I�����Ă��������B"
        Exit Sub
    End If
    
    ' �I�u�W�F�N�g�̈ʒu��ۑ�����z���������
    ReDim positions(1 To selectedShapes.Count, 1 To 2)
    
    ' �e�I�u�W�F�N�g�̈ʒu��ۑ�
    For i = 1 To selectedShapes.Count
        positions(i, 1) = selectedShapes(i).Left
        positions(i, 2) = selectedShapes(i).Top
    Next i
    
    ' �e�I�u�W�F�N�g��������̃I�u�W�F�N�g�̈ʒu�Ɉړ�
    For i = 1 To selectedShapes.Count - 1
        selectedShapes(i).Left = positions(i + 1, 1)
        selectedShapes(i).Top = positions(i + 1, 2)
    Next i
    
    ' �Ō�̃I�u�W�F�N�g���ŏ��̃I�u�W�F�N�g�̈ʒu�Ɉړ�
    selectedShapes(selectedShapes.Count).Left = positions(1, 1)
    selectedShapes(selectedShapes.Count).Top = positions(1, 2)
    
    ' �N���[���A�b�v
    Set selectedShapes = Nothing
End Sub

