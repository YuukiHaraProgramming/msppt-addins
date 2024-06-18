Attribute VB_Name = "SwapObj"

Sub SwapObj()
    Dim selectedShapes As ShapeRange
    Dim positions() As Variant
    Dim i As Integer
    
    ' 選択したオブジェクトを取得する
    Set selectedShapes = ActiveWindow.Selection.ShapeRange
    
    ' 選択オブジェクトが2つ以上ない場合は終了
    If selectedShapes.Count < 2 Then
        MsgBox "少なくとも2つのオブジェクトを選択してください。"
        Exit Sub
    End If
    
    ' オブジェクトの位置を保存する配列を初期化
    ReDim positions(1 To selectedShapes.Count, 1 To 2)
    
    ' 各オブジェクトの位置を保存
    For i = 1 To selectedShapes.Count
        positions(i, 1) = selectedShapes(i).Left
        positions(i, 2) = selectedShapes(i).Top
    Next i
    
    ' 各オブジェクトを一つずつ次のオブジェクトの位置に移動
    For i = 1 To selectedShapes.Count - 1
        selectedShapes(i).Left = positions(i + 1, 1)
        selectedShapes(i).Top = positions(i + 1, 2)
    Next i
    
    ' 最後のオブジェクトを最初のオブジェクトの位置に移動
    selectedShapes(selectedShapes.Count).Left = positions(1, 1)
    selectedShapes(selectedShapes.Count).Top = positions(1, 2)
    
    ' クリーンアップ
    Set selectedShapes = Nothing
End Sub

