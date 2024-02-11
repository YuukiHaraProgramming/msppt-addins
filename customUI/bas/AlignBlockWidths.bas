Attribute VB_Name = "AlignBlockWidths"
Sub AlignBlockWidths()
    Dim selectedShapes As ShapeRange
    Dim referenceShape As shape
    Dim shape As shape
    
    ' �I�������u���b�N���擾����
    Set selectedShapes = ActiveWindow.Selection.ShapeRange
    
    ' �Ō�ɑI�������u���b�N���Q�ƃu���b�N�Ƃ��Đݒ肷��
    Set referenceShape = selectedShapes(selectedShapes.Count)
    
    ' �Q�ƃu���b�N�̕����擾����
    Dim referenceWidth As Single
    referenceWidth = referenceShape.Width
    
    ' �u���b�N�̕����Q�ƃu���b�N�̕��ɑ�����
    For Each shape In selectedShapes
        shape.Width = referenceWidth
    Next shape
    
    ' �N���[���A�b�v
    Set selectedShapes = Nothing
    Set referenceShape = Nothing
    Set shape = Nothing
End Sub

