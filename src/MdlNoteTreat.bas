Attribute VB_Name = "MdlNoteTreat"
Option Explicit

' シェイプの情報を格納する構造体
Private Type ShapeInfo
    text As String
    Top As Single
    Left As Single
End Type

Sub CopyTextFullpath()
    Dim filePath As String
    CopyToClipboard GetTextFullpath
End Sub

Sub ExportNoteToText()
    ' 各スライドのノートをテキストファイルに抽出するサブルーチン (UTF-8対応版)
    Dim sld As Slide
    Dim notesText As String
    Dim textFileName As String
    Dim stream As Object ' ADODB.Stream
    Dim notesShape As Shape

    If ActivePresentation.Path = "" Then
        MsgBox "プレゼンテーションが保存されていません。一度ファイルを保存してから実行してください。", vbExclamation, "保存確認"
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    textFileName = GetTextFullpath
    
    ' ADODB.Streamオブジェクトを作成してUTF-8で書き込み
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open

    ' プレゼンテーションの各スライドをループする
    For Each sld In ActivePresentation.Slides
        ' ノートのテキストを抽出する (エラー回避付き)
        notesText = ""
        On Error Resume Next
        Set notesShape = GetNotesBodyShape(sld)
        If Not notesShape Is Nothing Then
            notesText = notesShape.TextFrame.TextRange.text
        End If
        On Error GoTo ErrorHandler ' エラーハンドラを戻す

        Dim slideNumber As Long
        slideNumber = sld.slideNumber

        ' スライド番号とノートのテキストをストリームに書き込む
        stream.WriteText "<<< Slide " & slideNumber & vbCrLf
        stream.WriteText TrimWhitespace(notesText) & vbCrLf
        stream.WriteText vbCrLf ' 読みやすさのために空行を追加
    Next sld

    ' ファイルに保存 (上書きモード: 2 = adSaveCreateOverWrite)
    stream.SaveToFile textFileName, 2
    stream.Close
    Set stream = Nothing

    Exit Sub ' 正常終了

ErrorHandler:
    If Not stream Is Nothing Then
        If stream.State = 1 Then stream.Close ' adStateOpen = 1
    End If
    MsgBox "エラーが発生しました。エラー番号: " & Err.Number & vbCrLf & "内容: " & Err.Description & vbCrLf & "ファイル名: " & textFileName, vbCritical
End Sub

Sub ImportNoteFromText()
    ' テキストファイルからノートを読み込み、各スライドに挿入するサブルーチン (改行修正版)
    Dim sld As Slide
    Dim textFileName As String
    Dim lineText As String
    Dim slideNumber As Long
    Dim stream As Object
    Dim targetSlide As Slide
    Dim allText As String
    Dim lines() As String
    Dim i As Long
    Dim notesShape As Shape
    
    On Error GoTo ErrorHandler
    textFileName = GetTextFullpath
    
    If Dir(textFileName) = "" Then
        MsgBox "インポートするテキストファイルが見つかりません。" & vbCrLf & "先に「ノート → テキストファイル」を実行してください。", vbExclamation, "ファイル不在"
        Exit Sub
    End If
    
    ' すべての既存のスライドのノートを削除する
    For Each sld In ActivePresentation.Slides
        On Error Resume Next
        Set notesShape = GetNotesBodyShape(sld)
        If Not notesShape Is Nothing Then
            notesShape.TextFrame.TextRange.text = ""
        End If
        On Error GoTo ErrorHandler
    Next sld

    ' ADODB.Streamオブジェクトを作成してUTF-8で読み込み
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile textFileName
    
    ' 全テキストを一括で読み込む
    allText = stream.ReadText
    stream.Close
    Set stream = Nothing
    
    ' ★修正ポイント：PowerPoint内部の改行(vbCr)と、ファイルの改行(vbCrLf)を
    ' 全て一律で vbLf に統一してから分割することで、スライド内の改行を保護する。
    allText = Replace(allText, vbCrLf, vbLf)
    allText = Replace(allText, vbCr, vbLf)
    
    ' 統一した vbLf で行ごとに分割
    lines = Split(allText, vbLf)
    
    slideNumber = 0
    Set targetSlide = Nothing

    ' 各行を処理
    For i = 0 To UBound(lines)
        lineText = lines(i) ' ※ここで Replace(..., vbCr, "") をしない！
        
        If InStr(lineText, "<<< Slide ") = 1 Then
            slideNumber = CLng(Replace(lineText, "<<< Slide ", ""))
            Set targetSlide = GetSlideByNumber(slideNumber)
        ElseIf InStr(lineText, "# Slide ") = 1 Then
            slideNumber = CLng(Mid(lineText, 9))
            Set targetSlide = GetSlideByNumber(slideNumber)
        Else
            ' スライドのノートにテキストを追加する (保持したスライドに対してのみ処理)
            If Not targetSlide Is Nothing Then
                On Error Resume Next
                Set notesShape = GetNotesBodyShape(targetSlide)
                If Not notesShape Is Nothing Then
                    notesShape.TextFrame.TextRange.text = _
                        notesShape.TextFrame.TextRange.text & lineText & vbCrLf
                End If
                On Error GoTo ErrorHandler
            End If
        End If
    Next i

    ' ★追加改善：インポート後に、各スライドのノート末尾に溜まった余分な改行を綺麗に削除する
    For Each sld In ActivePresentation.Slides
        On Error Resume Next
        Set notesShape = GetNotesBodyShape(sld)
        If Not notesShape Is Nothing Then
            notesShape.TextFrame.TextRange.text = TrimWhitespace(notesShape.TextFrame.TextRange.text)
        End If
        On Error GoTo ErrorHandler
    Next sld

    Exit Sub

ErrorHandler:
    If Not stream Is Nothing Then
        If stream.State = 1 Then stream.Close
    End If
    MsgBox "エラーが発生しました。エラー番号: " & Err.Number & vbCrLf & "内容: " & Err.Description & vbCrLf & "ファイル名: " & textFileName, vbCritical
End Sub

' スライド番号から安全にスライドオブジェクトを取得する補助関数
Private Function GetSlideByNumber(ByVal sNum As Long) As Slide
    Dim s As Slide
    For Each s In ActivePresentation.Slides
        If s.slideNumber = sNum Then
            Set GetSlideByNumber = s
            Exit Function
        End If
    Next s
    Set GetSlideByNumber = Nothing
End Function

' -------------------------------------------------------------------------
' ★新規追加：ノートの本文プレースホルダーを安全に取得する関数
' -------------------------------------------------------------------------
Private Function GetNotesBodyShape(sld As Slide) As Shape
    Dim shp As Shape
    On Error Resume Next
    ' まずは本文プレースホルダー(ppPlaceholderBody)を探す
    For Each shp In sld.NotesPage.Shapes
        If shp.Type = msoPlaceholder Then
            If shp.PlaceholderFormat.Type = ppPlaceholderBody Then
                Set GetNotesBodyShape = shp
                Exit Function
            End If
        End If
    Next shp
    ' 見つからなかった場合のフォールバック(従来のインデックス指定)
    Set GetNotesBodyShape = sld.NotesPage.Shapes.Placeholders(2)
    On Error GoTo 0
End Function

Sub AddNoteInSlides()
    Dim sld As Slide
    Dim slds As SlideRange
    
    If doAllSlides Then
        Set slds = ActivePresentation.Slides.Range
    Else
        On Error Resume Next
        Set slds = ActiveWindow.Selection.SlideRange
        
        If slds Is Nothing Then
            Set slds = ActivePresentation.Slides.Range(ActiveWindow.View.Slide.SlideIndex)
        End If
        On Error GoTo 0
        
        If slds Is Nothing Then
            MsgBox "スライドを特定できませんでした。スライド画面を一度クリックしてから再実行してください。", vbExclamation, "スライド特定エラー"
            Exit Sub
        End If
    End If
    
    For Each sld In slds
        Dim slideText As String
        Dim targetNotesShape As Shape
        
        ' スライドテキストを取得（グループ化対応版）
        slideText = GetSlideText(sld)
        Set targetNotesShape = GetNotesBodyShape(sld)
        
        If Not targetNotesShape Is Nothing Then
            If doOverride Then
                targetNotesShape.TextFrame.TextRange.text = slideText
            Else
                If targetNotesShape.TextFrame.TextRange.text = "" Then
                    targetNotesShape.TextFrame.TextRange.text = slideText
                End If
            End If
        End If
    Next sld
End Sub

Sub RemoveNoteinSlides()
    Dim sld As Slide
    Dim slds As SlideRange
    Dim targetNotesShape As Shape
    
    If doAllSlides Then
        Set slds = ActivePresentation.Slides.Range
    Else
        On Error Resume Next
        Set slds = ActiveWindow.Selection.SlideRange
        
        If slds Is Nothing Then
            Set slds = ActivePresentation.Slides.Range(ActiveWindow.View.Slide.SlideIndex)
        End If
        On Error GoTo 0
        
        If slds Is Nothing Then
            MsgBox "スライドを特定できませんでした。スライド画面を一度クリックしてから再実行してください。", vbExclamation, "スライド特定エラー"
            Exit Sub
        End If
    End If
    
    For Each sld In slds
        On Error Resume Next
        Set targetNotesShape = GetNotesBodyShape(sld)
        If Not targetNotesShape Is Nothing Then
            targetNotesShape.TextFrame.TextRange.text = ""
        End If
        On Error GoTo 0
    Next sld
End Sub

' -------------------------------------------------------------------------
' ★新規追加：グループ化された図形の中身も再帰的に探索してテキストを収集する関数
' -------------------------------------------------------------------------
Private Sub CollectTextShapes(ByVal shps As Object, ByRef shapeInfos() As ShapeInfo, ByRef shapeCount As Integer)
    Dim shp As Shape
    Dim i As Integer
    
    For i = 1 To shps.Count
        Set shp = shps(i)
        
        If shp.Type = msoGroup Then
            ' グループ化されている場合は再帰呼び出しで中身を探索
            CollectTextShapes shp.GroupItems, shapeInfos, shapeCount
        Else
            ' テキストフレームを持っているか確認
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    shapeCount = shapeCount + 1
                    ReDim Preserve shapeInfos(1 To shapeCount)
                    shapeInfos(shapeCount).text = shp.TextFrame.TextRange.text
                    shapeInfos(shapeCount).Top = shp.Top
                    shapeInfos(shapeCount).Left = shp.Left
                End If
            End If
        End If
    Next i
End Sub

' スライドからテキストを取得する関数 (グループ化対応＆クリーンアップ版)
Function GetSlideText(ByVal sld As Slide) As String
    Dim shapeInfos() As ShapeInfo
    Dim shapeCount As Integer
    Dim i As Integer, j As Integer
    Dim temp As ShapeInfo
    Dim resultText As String
    
    shapeCount = 0
    
    ' スライド内の全シェイプからテキストを収集（再帰処理でグループ化対応）
    CollectTextShapes sld.Shapes, shapeInfos, shapeCount
    
    ' テキストが見つからなかった場合
    If shapeCount = 0 Then
        GetSlideText = ""
        Exit Function
    End If
    
    ' バブルソートで並べ替え (上から下、左から右へ)
    For i = 1 To shapeCount - 1
        For j = 1 To shapeCount - i
            If (shapeInfos(j).Top - shapeInfos(j + 1).Top > 5) Or _
               (Abs(shapeInfos(j).Top - shapeInfos(j + 1).Top) <= 5 And shapeInfos(j).Left > shapeInfos(j + 1).Left) Then
                temp = shapeInfos(j)
                shapeInfos(j) = shapeInfos(j + 1)
                shapeInfos(j + 1) = temp
            End If
        Next j
    Next i
    
    ' 並べ替えたテキストを結合
    For i = 1 To shapeCount
        resultText = resultText & shapeInfos(i).text & vbNewLine
    Next i
    
    GetSlideText = Trim(resultText)
End Function