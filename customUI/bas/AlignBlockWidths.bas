Attribute VB_Name = "AlignBlockWidths"
Sub AlignBlockWidths()
    Dim selectedShapes As ShapeRange
    Dim referenceShape As shape
    Dim shape As shape
    
    ' 選択したブロックを取得する
    Set selectedShapes = ActiveWindow.Selection.ShapeRange
    
    ' 最後に選択したブロックを参照ブロックとして設定する
    Set referenceShape = selectedShapes(selectedShapes.Count)
    
    ' 参照ブロックの幅を取得する
    Dim referenceWidth As Single
    referenceWidth = referenceShape.Width
    
    ' ブロックの幅を参照ブロックの幅に揃える
    For Each shape In selectedShapes
        shape.Width = referenceWidth
    Next shape
    
    ' クリーンアップ
    Set selectedShapes = Nothing
    Set referenceShape = Nothing
    Set shape = Nothing
End Sub

