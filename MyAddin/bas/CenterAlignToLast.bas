Attribute VB_Name = "CenterAlignToLast"

Sub CenterAlignToLast()
    Dim selectedShapes As ShapeRange
    Dim lastShape As shape
    Dim slide As slide
    Dim referenceLeft As Single
    Dim referenceWidth As Single
    Dim i As Integer
    
    ' ���݂̃A�N�e�B�u�X���C�h���擾
    Set slide = ActiveWindow.View.slide
    
    ' �I���I�u�W�F�N�g���Ȃ��ꍇ�͏I��
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "�I�u�W�F�N�g���I������Ă��܂���B"
        Exit Sub
    End If
    
    ' �I�������u���b�N���擾����
    Set selectedShapes = ActiveWindow.Selection.ShapeRange
    
    ' �Ō�ɑI�������u���b�N���擾
    Set lastShape = selectedShapes(selectedShapes.Count)
    
    ' �Q�ƃu���b�N�̕��ƍ��[���擾
    referenceWidth = lastShape.Width
    referenceLeft = lastShape.Left
    
    ' ���̃I�u�W�F�N�g�𒆉�����
    For i = 1 To selectedShapes.Count - 1
        If selectedShapes(i).Type <> msoPlaceholder Then
            ' �Ō�̃I�u�W�F�N�g�ȊO�����E��������
            selectedShapes(i).Left = referenceLeft + (referenceWidth - selectedShapes(i).Width) / 2
        End If
    Next i
    
    ' �N���[���A�b�v
    Set selectedShapes = Nothing
    Set lastShape = Nothing
    Set slide = Nothing
End Sub

