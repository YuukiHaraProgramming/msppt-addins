Attribute VB_Name = "InsertSource"

Sub InsertSource()
    ' 定数
    Const cmToPoints As Double = 28.3465
    
    ' 変数宣言
    Dim slide As slide
    Dim rect As shape
    Dim leftPos As Double
    Dim topPos As Double
    Dim shapeWidth As Double
    Dim shapeHeight As Double
    Dim fontColor As Long
    
    ' 現在のアクティブスライドを取得
    Set slide = ActiveWindow.View.slide
    
    ' 長方形の位置とサイズを設定 (cmをポイントに変換)
    leftPos = 0.56 * cmToPoints
    topPos = 17.94 * cmToPoints
    shapeWidth = 17.4 * cmToPoints
    shapeHeight = 1 * cmToPoints
    fontColor = RGB(0, 0, 0)
    
    ' 長方形のオブジェクトを作成
    Set rect = slide.Shapes.AddShape(Type:=msoShapeRectangle, Left:=leftPos, Top:=topPos, Width:=shapeWidth, Height:=shapeHeight)
    
    ' 長方形のプロパティを設定
    With rect
        .Fill.Transparency = 1 ' 塗りつぶしの透明度を設定（塗りつぶしなし）
        .line.Transparency = 1 ' アウトラインの透明度を設定（アウトラインなし）
        .textFrame.TextRange.Text = "出所：" ' テキストを設定
        With .textFrame.TextRange.Font
            .Size = 11
            .Color.RGB = fontColor
        End With
        .textFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft ' テキストを左揃えに設定
        .TextFrame2.VerticalAnchor = msoAnchorTop ' テキストを上揃えに設定
    End With
    
    ' クリーンアップ
    Set rect = Nothing
    Set slide = Nothing
End Sub

