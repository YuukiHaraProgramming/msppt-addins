Attribute VB_Name = "InsertSource"

Sub InsertSource()
    ' �萔
    Const cmToPoints As Double = 28.3465
    
    ' �ϐ��錾
    Dim slide As slide
    Dim rect As shape
    Dim leftPos As Double
    Dim topPos As Double
    Dim shapeWidth As Double
    Dim shapeHeight As Double
    Dim fontColor As Long
    
    ' ���݂̃A�N�e�B�u�X���C�h���擾
    Set slide = ActiveWindow.View.slide
    
    ' �����`�̈ʒu�ƃT�C�Y��ݒ� (cm���|�C���g�ɕϊ�)
    leftPos = 0.56 * cmToPoints
    topPos = 17.94 * cmToPoints
    shapeWidth = 17.4 * cmToPoints
    shapeHeight = 1 * cmToPoints
    fontColor = RGB(0, 0, 0)
    
    ' �����`�̃I�u�W�F�N�g���쐬
    Set rect = slide.Shapes.AddShape(Type:=msoShapeRectangle, Left:=leftPos, Top:=topPos, Width:=shapeWidth, Height:=shapeHeight)
    
    ' �����`�̃v���p�e�B��ݒ�
    With rect
        .Fill.Transparency = 1 ' �h��Ԃ��̓����x��ݒ�i�h��Ԃ��Ȃ��j
        .line.Transparency = 1 ' �A�E�g���C���̓����x��ݒ�i�A�E�g���C���Ȃ��j
        .textFrame.TextRange.Text = "�o���F" ' �e�L�X�g��ݒ�
        With .textFrame.TextRange.Font
            .Size = 11
            .Color.RGB = fontColor
        End With
        .textFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft ' �e�L�X�g���������ɐݒ�
        .TextFrame2.VerticalAnchor = msoAnchorTop ' �e�L�X�g���㑵���ɐݒ�
    End With
    
    ' �N���[���A�b�v
    Set rect = Nothing
    Set slide = Nothing
End Sub

