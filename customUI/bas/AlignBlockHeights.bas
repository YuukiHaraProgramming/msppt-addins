Attribute VB_Name = "AlignBlockHeights"
Sub AlignBlockHeights()
    Dim selectedShapes As ShapeRange
    Dim referenceShape As shape
    Dim shape As shape
    
    ' �I�������u���b�N���擾����
    Set selectedShapes = ActiveWindow.Selection.ShapeRange
    
    ' �Ō�ɑI�������u���b�N���Q�ƃu���b�N�Ƃ��Đݒ肷��
    Set referenceShape = selectedShapes(selectedShapes.Count)
    
    ' �Q�ƃu���b�N�̏c�����擾����
    Dim referenceHeight As Single
    referenceHeight = referenceShape.Height
    
    ' �u���b�N�̏c�����Q�ƃu���b�N�̏c���ɑ�����
    For Each shape In selectedShapes
        shape.Height = referenceHeight
    Next shape
End Sub

