Attribute VB_Name = "CreateIncircleForTriangle"
' Create an incircle for the selected isosceles triangle

Sub CreateIncircleForTriangle()
    Dim selectedShape As Shape
    Dim vertices(1 To 3, 1 To 2) As Single
    Dim sideAB, sideBC, sideCA As Single
    Dim semiPerimeter, triangleArea, inradius As Single
    Dim centerX, centerY As Single
    Dim incircle As Shape
    
    ' 三角形が選択されているか確認
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "三角形が選択されていません。"
        Exit Sub
    End If
    
    ' 選択した三角形を取得
    Set selectedShape = ActiveWindow.Selection.ShapeRange(1)
    
    ' 三角形の頂点座標を取得
    vertices(1, 1) = selectedShape.Left
    vertices(1, 2) = selectedShape.Top + selectedShape.Height
    vertices(2, 1) = selectedShape.Left + selectedShape.Width
    vertices(2, 2) = vertices(1, 2)
    vertices(3, 1) = (selectedShape.Left + vertices(2, 1)) / 2
    vertices(3, 2) = selectedShape.Top
    
    ' 三角形の辺の長さを計算
    sideAB = Distance(vertices(1, 1), vertices(1, 2), vertices(2, 1), vertices(2, 2))
    sideBC = Distance(vertices(2, 1), vertices(2, 2), vertices(3, 1), vertices(3, 2))
    sideCA = Distance(vertices(3, 1), vertices(3, 2), vertices(1, 1), vertices(1, 2))
    
    ' ヘロンの公式を用いて三角形の面積を計算
    semiPerimeter = (sideAB + sideBC + sideCA) / 2
    triangleArea = Sqr(semiPerimeter * (semiPerimeter - sideAB) * (semiPerimeter - sideBC) * (semiPerimeter - sideCA))
    
    ' 内接円の半径を計算
    inradius = triangleArea / semiPerimeter
    
    ' 三角形の内心の座標を計算
    centerX = (sideBC * vertices(1, 1) + sideCA * vertices(2, 1) + sideAB * vertices(3, 1)) / (sideAB + sideBC + sideCA)
    centerY = (sideBC * vertices(1, 2) + sideCA * vertices(2, 2) + sideAB * vertices(3, 2)) / (sideAB + sideBC + sideCA)
    
    ' 円を作成
    Set incircle = ActiveWindow.View.Slide.Shapes.AddShape(msoShapeOval, centerX - inradius, centerY - inradius, 2 * inradius, 2 * inradius)
    ' 図形の線なし
    incircle.Line.Visible = msoFalse
    
    ' クリーンアップ
    Set selectedShape = Nothing
    Set incircle = Nothing
End Sub

' 2点間の距離を計算する関数
Function Distance(x1 As Single, y1 As Single, x2 As Single, y2 As Single) As Single
    Distance = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
End Function

