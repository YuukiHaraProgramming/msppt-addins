Attribute VB_Name = "ReplaceObj"
Sub ReplaceObj()
    Dim selectedShapes As ShapeRange
    Dim referenceShape As shape
    Dim positions() As Variant
    Dim i As Integer
    
    ' �I�������I�u�W�F�N�g���擾����
    Set selectedShapes = ActiveWindow.Selection.ShapeRange
    
    ' �I���I�u�W�F�N�g��2�ȏ�Ȃ��ꍇ�͏I��
    If selectedShapes.Count < 2 Then
        MsgBox "���Ȃ��Ƃ�2�̃I�u�W�F�N�g��I�����Ă��������B"
        Exit Sub
    End If
    
    ' �Ō�ɑI�������I�u�W�F�N�g���Q�ƃI�u�W�F�N�g�Ƃ��Đݒ肷��
    Set referenceShape = selectedShapes(selectedShapes.Count)
    
    ' �I�u�W�F�N�g�̈ʒu��ۑ�����z���������
    ReDim positions(1 To selectedShapes.Count - 1, 1 To 2)
    
    ' �Ō�ɑI�����ꂽ�I�u�W�F�N�g���������e�I�u�W�F�N�g�̈ʒu��ۑ�
    For i = 1 To selectedShapes.Count - 1
        positions(i, 1) = selectedShapes(i).Left
        positions(i, 2) = selectedShapes(i).Top
    Next i
    
    ' �Q�ƃI�u�W�F�N�g��N-1�R�s�[����
    Dim copies() As shape
    ReDim copies(1 To selectedShapes.Count - 1)
    For i = 1 To selectedShapes.Count - 1
        Set copies(i) = referenceShape.Duplicate.Item(1)
    Next i
    
    ' ���̃I�u�W�F�N�g���폜����
    For i = 1 To selectedShapes.Count - 1
        selectedShapes(i).Delete
    Next i
    
    ' �ۑ������ʒu�ɃR�s�[�����I�u�W�F�N�g��z�u����
    For i = 1 To UBound(copies)
        copies(i).Left = positions(i, 1)
        copies(i).Top = positions(i, 2)
    Next i
    
    ' �N���[���A�b�v
    Set selectedShapes = Nothing
    Set referenceShape = Nothing
    For i = 1 To UBound(copies)
        Set copies(i) = Nothing
    Next i
End Sub

