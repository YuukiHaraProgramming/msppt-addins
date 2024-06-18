Attribute VB_Name = "ReplaceObj"
Sub ReplaceObj()
    Dim selectedShapes As ShapeRange
    Dim referenceShape As shape
    Dim positions() As Variant
    Dim i As Integer
    
    ' 選択したオブジェクトを取得する
    Set selectedShapes = ActiveWindow.Selection.ShapeRange
    
    ' 選択オブジェクトが2つ以上ない場合は終了
    If selectedShapes.Count < 2 Then
        MsgBox "少なくとも2つのオブジェクトを選択してください。"
        Exit Sub
    End If
    
    ' 最後に選択したオブジェクトを参照オブジェクトとして設定する
    Set referenceShape = selectedShapes(selectedShapes.Count)
    
    ' オブジェクトの位置を保存する配列を初期化
    ReDim positions(1 To selectedShapes.Count - 1, 1 To 2)
    
    ' 最後に選択されたオブジェクトを除いた各オブジェクトの位置を保存
    For i = 1 To selectedShapes.Count - 1
        positions(i, 1) = selectedShapes(i).Left
        positions(i, 2) = selectedShapes(i).Top
    Next i
    
    ' 参照オブジェクトをN-1個コピーする
    Dim copies() As shape
    ReDim copies(1 To selectedShapes.Count - 1)
    For i = 1 To selectedShapes.Count - 1
        Set copies(i) = referenceShape.Duplicate.Item(1)
    Next i
    
    ' 元のオブジェクトを削除する
    For i = 1 To selectedShapes.Count - 1
        selectedShapes(i).Delete
    Next i
    
    ' 保存した位置にコピーしたオブジェクトを配置する
    For i = 1 To UBound(copies)
        copies(i).Left = positions(i, 1)
        copies(i).Top = positions(i, 2)
    Next i
    
    ' クリーンアップ
    Set selectedShapes = Nothing
    Set referenceShape = Nothing
    For i = 1 To UBound(copies)
        Set copies(i) = Nothing
    Next i
End Sub

