Attribute VB_Name = "ToggleTextWrap"
Sub ToggleTextWrap()
    Dim selectedShapes As ShapeRange
    Dim shape As shape
    
    ' 選択したオブジェクトを取得する
    Set selectedShapes = ActiveWindow.Selection.ShapeRange
    
    ' 各オブジェクトの「テキストが図形内で折り返す」の設定をトグルする
    For Each shape In selectedShapes
        With shape.textFrame
            If .WordWrap = msoTrue Then
                .WordWrap = msoFalse
            Else
                .WordWrap = msoTrue
            End If
        End With
    Next shape
    
    ' クリーンアップ
    Set selectedShapes = Nothing
    Set shape = Nothing
End Sub

