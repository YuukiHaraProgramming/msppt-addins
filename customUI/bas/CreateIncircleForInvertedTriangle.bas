Attribute VB_Name = "CreateIncircleForInvTriangle"
' Create an incircle for the selected isosceles inverted triangle

Sub CreateIncircleForInvertedTriangle()
    Dim selectedShape As Shape
    Dim vertices(1 To 3, 1 To 2) As Single
    Dim sideAB, sideBC, sideCA As Single
    Dim semiPerimeter, triangleArea, inradius As Single
    Dim centerX, centerY As Single
    Dim incircle As Shape
    
    ' �O�p�`���I������Ă��邩�m�F
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "�O�p�`���I������Ă��܂���B"
        Exit Sub
    End If
    
    ' �I�������O�p�`���擾
    Set selectedShape = ActiveWindow.Selection.ShapeRange(1)
    
    ' �O�p�`�̒��_���W���擾
    vertices(1, 1) = selectedShape.Left
    vertices(1, 2) = selectedShape.Top
    vertices(2, 1) = selectedShape.Left + selectedShape.Width
    vertices(2, 2) = vertices(1, 2)
    vertices(3, 1) = (selectedShape.Left + vertices(2, 1)) / 2
    vertices(3, 2) = selectedShape.Top + selectedShape.Height
    
    ' �O�p�`�̕ӂ̒������v�Z
    sideAB = Distance(vertices(1, 1), vertices(1, 2), vertices(2, 1), vertices(2, 2))
    sideBC = Distance(vertices(2, 1), vertices(2, 2), vertices(3, 1), vertices(3, 2))
    sideCA = Distance(vertices(3, 1), vertices(3, 2), vertices(1, 1), vertices(1, 2))
    
    ' �w�����̌�����p���ĎO�p�`�̖ʐς��v�Z
    semiPerimeter = (sideAB + sideBC + sideCA) / 2
    triangleArea = Sqr(semiPerimeter * (semiPerimeter - sideAB) * (semiPerimeter - sideBC) * (semiPerimeter - sideCA))
    
    ' ���ډ~�̔��a���v�Z
    inradius = triangleArea / semiPerimeter
    
    ' �O�p�`�̓��S�̍��W���v�Z
    centerX = (sideBC * vertices(1, 1) + sideCA * vertices(2, 1) + sideAB * vertices(3, 1)) / (sideAB + sideBC + sideCA)
    centerY = (sideBC * vertices(1, 2) + sideCA * vertices(2, 2) + sideAB * vertices(3, 2)) / (sideAB + sideBC + sideCA)
    
    ' �~���쐬
    Set incircle = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeOval, centerX - inradius, centerY - inradius, 2 * inradius, 2 * inradius)
    '�}�`�̐��Ȃ�
    incircle.Line.Visible = msoFalse
    
    ' �N���[���A�b�v
    Set selectedShape = Nothing
    Set incircle = Nothing
End Sub

' 2�_�Ԃ̋������v�Z����֐�
Function Distance(x1 As Single, y1 As Single, x2 As Single, y2 As Single) As Single
    Distance = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
End Function


