Attribute VB_Name = "ToggleTextWrap"
Sub ToggleTextWrap()
    Dim selectedShapes As ShapeRange
    Dim shape As shape
    
    ' �I�������I�u�W�F�N�g���擾����
    Set selectedShapes = ActiveWindow.Selection.ShapeRange
    
    ' �e�I�u�W�F�N�g�́u�e�L�X�g���}�`���Ő܂�Ԃ��v�̐ݒ���g�O������
    For Each shape In selectedShapes
        With shape.textFrame
            If .WordWrap = msoTrue Then
                .WordWrap = msoFalse
            Else
                .WordWrap = msoTrue
            End If
        End With
    Next shape
    
    ' �N���[���A�b�v
    Set selectedShapes = Nothing
    Set shape = Nothing
End Sub

