Attribute VB_Name = "MoveRightToLastLeft"
' Move right to align with the left edge of the last selected object


Sub MoveRightToLastLeft()
    Dim selectedItem As Shape
    Dim referenceItem As Shape
    Dim slide As slide
    Dim refLeftPosition As Single
    Dim i As Integer
    
    ' 現在のアクティブスライドを取得
    Set slide = ActiveWindow.View.slide
    
    ' 選択オブジェクトがない場合は終了
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "オブジェクトが選択されていません。"
        Exit Sub
    End If
    
    ' 最後に選択されたオブジェクトを基準オブジェクトとして設定
    Set referenceItem = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count)
    
    ' 基準オブジェクトの左端の位置を取得
    refLeftPosition = referenceItem.Left
    
    ' 他のすべてのオブジェクトを移動
    For i = 1 To ActiveWindow.Selection.ShapeRange.Count - 1
        Set selectedItem = ActiveWindow.Selection.ShapeRange(i)
        
        ' 基準オブジェクト以外を右に移動
        selectedItem.Left = refLeftPosition - selectedItem.Width
    Next i
    
    ' クリーンアップ
    Set selectedItem = Nothing
    Set referenceItem = Nothing
    Set slide = Nothing
End Sub

