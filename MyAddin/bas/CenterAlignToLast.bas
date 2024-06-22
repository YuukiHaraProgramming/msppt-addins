Attribute VB_Name = "CenterAlignToLast"

Sub CenterAlignToLast()
    Dim selectedShapes As ShapeRange
    Dim lastShape As shape
    Dim slide As slide
    Dim referenceLeft As Single
    Dim referenceWidth As Single
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
    
    ' 参照ブロックの幅と左端を取得
    referenceWidth = lastShape.Width
    referenceLeft = lastShape.Left
    
    ' 他のオブジェクトを中央揃え
    For i = 1 To selectedShapes.Count - 1
        If selectedShapes(i).Type <> msoPlaceholder Then
            ' 最後のオブジェクト以外を左右中央揃え
            selectedShapes(i).Left = referenceLeft + (referenceWidth - selectedShapes(i).Width) / 2
        End If
    Next i
    
    ' クリーンアップ
    Set selectedShapes = Nothing
    Set lastShape = Nothing
    Set slide = Nothing
End Sub

