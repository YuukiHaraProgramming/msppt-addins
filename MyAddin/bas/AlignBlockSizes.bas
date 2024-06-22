Attribute VB_Name = "AlignBlockSizes"
Sub AlignBlockSizes()
    Dim selectedShapes As ShapeRange
    Dim referenceShape As shape
    Dim shape As shape
    
    ' �I�������u���b�N���擾����
    Set selectedShapes = ActiveWindow.Selection.ShapeRange
    
    ' �Ō�ɑI�������u���b�N���Q�ƃu���b�N�Ƃ��Đݒ肷��
    Set referenceShape = selectedShapes(selectedShapes.Count)
    
    ' �u���b�N�̕����Q�ƃu���b�N�̕��ɑ�����
    Dim referenceWidth As Single
    referenceWidth = referenceShape.Width
    For Each shape In selectedShapes
        shape.Width = referenceWidth
    Next shape
    
    ' �u���b�N�̏c�����Q�ƃu���b�N�̏c���ɑ�����
    Dim referenceHeight As Single
    referenceHeight = referenceShape.Height
    For Each shape In selectedShapes
        shape.Height = referenceHeight
    Next shape
    
    ' �N���[���A�b�v
    Set selectedShapes = Nothing
    Set referenceShape = Nothing
    Set shape = Nothing
End Sub

