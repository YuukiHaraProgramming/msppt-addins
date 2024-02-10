Attribute VB_Name = "AlignBlockHeights"
Sub AlignBlockHeights()
    Dim selectedShapes As ShapeRange
    Dim referenceShape As shape
    Dim shape As shape
    
    ' 選択したブロックを取得する
    Set selectedShapes = ActiveWindow.Selection.ShapeRange
    
    ' 最後に選択したブロックを参照ブロックとして設定する
    Set referenceShape = selectedShapes(selectedShapes.Count)
    
    ' 参照ブロックの縦幅を取得する
    Dim referenceHeight As Single
    referenceHeight = referenceShape.Height
    
    ' ブロックの縦幅を参照ブロックの縦幅に揃える
    For Each shape In selectedShapes
        shape.Height = referenceHeight
    Next shape
End Sub

