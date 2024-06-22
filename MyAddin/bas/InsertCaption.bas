Attribute VB_Name = "InsertCaption"
Sub InsertCaption()
    Dim slide As slide
    Dim line As shape
    Dim rect As shape
    Dim groupItems As ShapeRange
    Dim groupShape As shape
    Dim captionText As String
    Dim fontColor As Long
    Dim lineColor As Long
    Dim cmToPoints As Single
    Dim leftPosition As Single
    Dim topPosition As Single
    
    ' cm���|�C���g�ɕϊ�����萔
    cmToPoints = 28.3465
    
    ' ���݂̃A�N�e�B�u�X���C�h���擾
    Set slide = ActiveWindow.View.slide
    
    ' �e�L�X�g�ƐF��ݒ�
    captionText = "caption"
    fontColor = RGB(89, 89, 89)
    lineColor = RGB(89, 89, 89)
    
    ' �����쐬
    Set line = slide.Shapes.AddLine(BeginX:=0, BeginY:=0, EndX:=26.4 * cmToPoints, EndY:=0)
    With line.line
        .Weight = 0.75
        .ForeColor.RGB = lineColor
    End With
    
    ' �����`���쐬
    Set rect = slide.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=24.25 * cmToPoints, Height:=1 * cmToPoints)
    With rect
        .Fill.Transparency = 1
        .line.Transparency = 1
        .textFrame.TextRange.Text = captionText
        With .textFrame.TextRange.Font
            .Name = "Yu Gothic UI"
            .Size = 14
            .Bold = msoTrue
            .Color.RGB = fontColor
        End With
        .textFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
    End With
    
    ' ���ƒ����`�����������������ō��킹��
    line.Top = rect.Top + rect.Height - line.Height
    line.Left = rect.Left
    
    ' ���ƒ����`���O���[�v������
    Set groupItems = slide.Shapes.Range(Array(line.Name, rect.Name))
    Set groupShape = groupItems.Group
    
    ' �O���[�v�������I�u�W�F�N�g���w��̈ʒu�ɔz�u����
    leftPosition = 0.56 * cmToPoints
    topPosition = 4.29 * cmToPoints
    groupShape.Left = leftPosition
    groupShape.Top = topPosition
    
    ' �N���[���A�b�v
    Set line = Nothing
    Set rect = Nothing
    Set groupItems = Nothing
    Set groupShape = Nothing
    Set slide = Nothing
End Sub

