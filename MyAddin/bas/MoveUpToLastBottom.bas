Attribute VB_Name = "MoveUpToLastBottom"
' Move up to align with the bottom edge of the last selected object

Sub MoveUpToLastBottom()
    Dim selectedItem As Shape
    Dim referenceItem As Shape
    Dim slide As slide
    Dim refBottomPosition As Single
    Dim i As Integer
    
    ' ���݂̃A�N�e�B�u�X���C�h���擾
    Set slide = ActiveWindow.View.slide
    
    ' �I���I�u�W�F�N�g���Ȃ��ꍇ�͏I��
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "�I�u�W�F�N�g���I������Ă��܂���B"
        Exit Sub
    End If
    
    ' �Ō�ɑI�����ꂽ�I�u�W�F�N�g����I�u�W�F�N�g�Ƃ��Đݒ�
    Set referenceItem = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count)
    
    ' ��I�u�W�F�N�g�̉��[�̈ʒu���擾
    refBottomPosition = referenceItem.Top + referenceItem.Height
    
    ' ���̂��ׂẴI�u�W�F�N�g���ړ�
    For i = 1 To ActiveWindow.Selection.ShapeRange.Count - 1
        Set selectedItem = ActiveWindow.Selection.ShapeRange(i)
        
        ' ��I�u�W�F�N�g�ȊO����Ɉړ�
        selectedItem.Top = refBottomPosition
    Next i
    
    ' �N���[���A�b�v
    Set selectedItem = Nothing
    Set referenceItem = Nothing
    Set slide = Nothing
End Sub

