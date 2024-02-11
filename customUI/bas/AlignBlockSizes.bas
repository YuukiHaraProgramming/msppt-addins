Attribute VB_Name = "AlignBlockSizes"
Sub AlignBlockSizes()
    Dim selectedShapes As ShapeRange
    Dim referenceShape As shape
    Dim shape As shape
    
    ' 選択したブロックを取得する
    Set selectedShapes = ActiveWindow.Selection.ShapeRange
    
    ' 最後に選択したブロックを参照ブロックとして設定する
    Set referenceShape = selectedShapes(selectedShapes.Count)
    
    ' ブロックの幅を参照ブロックの幅に揃える
    Dim referenceWidth As Single
    referenceWidth = referenceShape.Width
    For Each shape In selectedShapes
        shape.Width = referenceWidth
    Next shape
    
    ' ブロックの縦幅を参照ブロックの縦幅に揃える
    Dim referenceHeight As Single
    referenceHeight = referenceShape.Height
    For Each shape In selectedShapes
        shape.Height = referenceHeight
    Next shape
    
    ' クリーンアップ
    Set selectedShapes = Nothing
    Set referenceShape = Nothing
    Set shape = Nothing
End Sub

