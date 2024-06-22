Attribute VB_Name = "MiddleAlignToLast"
Sub MiddleAlignToLast()
    Dim selectedShapes As ShapeRange
    Dim lastShape As shape
    Dim slide As slide
    Dim referenceTop As Single
    Dim referenceHeight As Single
    Dim i As Integer
    
    ' 現在のアクティブスライドを取得
    Set slide = ActiveWindow.View.slide
    
    ' 選択オブジェクトがない場合は終了
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "オブジェクトが選択されていません。"
        Exit Sub
    End If
    
    ' 選択したブロックを取得する
    Set selectedShapes = ActiveWindow.Selection.ShapeRange
    
    ' 最後に選択したブロックを取得
    Set lastShape = selectedShapes(selectedShapes.Count)
    
    ' 参照ブロックの高さと上端を取得
    referenceHeight = lastShape.Height
    referenceTop = lastShape.Top
    
    ' 他のオブジェクトを上下中央揃え
    For i = 1 To selectedShapes.Count - 1
        If selectedShapes(i).Type <> msoPlaceholder Then
            ' 最後のオブジェクト以外を上下中央揃え
            selectedShapes(i).Top = referenceTop + (referenceHeight - selectedShapes(i).Height) / 2
        End If
    Next i
    
    ' クリーンアップ
    Set selectedShapes = Nothing
    Set lastShape = Nothing
    Set slide = Nothing
End Sub

