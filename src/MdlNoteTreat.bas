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
        notesText = sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text
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
    ' テキストファイルからノートを読み込み、各スライドに挿入するサブルーチン (高速化＆UTF-8対応版)
    Dim sld As Slide
    Dim textFileName As String
    Dim lineText As String
    Dim slideNumber As Long
    Dim stream As Object
    Dim targetSlide As Slide
    Dim allText As String
    Dim lines() As String
    Dim i As Long
    
    On Error GoTo ErrorHandler
    textFileName = GetTextFullpath
    
    If Dir(textFileName) = "" Then
        MsgBox "インポートするテキストファイルが見つかりません。" & vbCrLf & "先に「ノート → テキストファイル」を実行してください。", vbExclamation, "ファイル不在"
        Exit Sub
    End If
    
    ' すべての既存のスライドのノートを削除する
    For Each sld In ActivePresentation.Slides
        On Error Resume Next
        sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text = ""
        On Error GoTo ErrorHandler
    Next sld

    ' ADODB.Streamオブジェクトを作成してUTF-8で読み込み
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile textFileName
    
    ' 全テキストを一括で読み込む（高速化）
    allText = stream.ReadText
    stream.Close
    Set stream = Nothing
    
    ' 改行コード(LF)で分割
    lines = Split(allText, vbLf)
    
    slideNumber = 0
    Set targetSlide = Nothing

    ' 各行を処理 (テキストの行数分だけループ)
    For i = 0 To UBound(lines)
        ' CR(キャリッジリターン)が残っている場合は除去
        lineText = Replace(lines(i), vbCr, "")
        
        If InStr(lineText, "<<< Slide ") = 1 Then
            slideNumber = CLng(Replace(lineText, "<<< Slide ", ""))
            ' ターゲットとなるスライドを取得し、変数に保持する
            Set targetSlide = GetSlideByNumber(slideNumber)
        ElseIf InStr(lineText, "# Slide ") = 1 Then
            slideNumber = CLng(Mid(lineText, 9))
            Set targetSlide = GetSlideByNumber(slideNumber)
        Else
            ' スライドのノートにテキストを追加する (保持したスライドに対してのみ処理)
            If Not targetSlide Is Nothing Then
                On Error Resume Next
                targetSlide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text = _
                    targetSlide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text & lineText & vbCrLf
                On Error GoTo ErrorHandler
            End If
        End If
    Next i

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

Sub AddNoteInSlides()
    Dim sld As Slide
    Dim slds As SlideRange
    
    If doAllSlides Then
        Set slds = ActivePresentation.Slides.Range
    Else
        On Error Resume Next ' ▼ エラーが発生してもプログラムを止めずに次へ進む
        
        ' 1. まずは現在の選択状態（サムネイル、図形、テキストなど）からスライドの取得を試みる
        Set slds = ActiveWindow.Selection.SlideRange
        
        ' 2. リボン操作中などで上記が失敗(Nothing)した場合、現在画面に表示されているスライドを取得
        If slds Is Nothing Then
            Set slds = ActivePresentation.Slides.Range(ActiveWindow.View.Slide.SlideIndex)
        End If
        
        On Error GoTo 0 ' ▲ エラー無視の設定を解除（以降の予期せぬエラーは通常通り表示する）
        
        ' 3. 万が一、何らかの理由でどうしても取得できなかった場合の安全装置
        If slds Is Nothing Then
            MsgBox "スライドを特定できませんでした。スライド画面を一度クリックしてから再実行してください。", vbExclamation, "スライド特定エラー"
            Exit Sub
        End If
    End If
    
    For Each sld In slds
        Dim slideText As String
        Dim notesShape As Shape
        Dim notesText As String
        
        ' スライドテキストを取得
        slideText = GetSlideText(sld)
        
        If doOverride Then
            sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text = slideText
        Else
            If sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text = "" Then
                sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text = slideText
            End If
        End If
    Next sld
End Sub

Sub RemoveNoteinSlides()
    Dim sld As Slide
    Dim slds As SlideRange
    
    If doAllSlides Then
        Set slds = ActivePresentation.Slides.Range
    Else
        On Error Resume Next ' ▼ エラーが発生してもプログラムを止めずに次へ進む
        
        ' 1. まずは現在の選択状態（サムネイル、図形、テキストなど）からスライドの取得を試みる
        Set slds = ActiveWindow.Selection.SlideRange
        
        ' 2. リボン操作中などで上記が失敗(Nothing)した場合、現在画面に表示されているスライドを取得
        If slds Is Nothing Then
            Set slds = ActivePresentation.Slides.Range(ActiveWindow.View.Slide.SlideIndex)
        End If
        
        On Error GoTo 0 ' ▲ エラー無視の設定を解除（以降の予期せぬエラーは通常通り表示する）
        
        ' 3. 万が一、何らかの理由でどうしても取得できなかった場合の安全装置
        If slds Is Nothing Then
            MsgBox "スライドを特定できませんでした。スライド画面を一度クリックしてから再実行してください。", vbExclamation, "スライド特定エラー"
            Exit Sub
        End If
    End If
    
    For Each sld In slds
        sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text = ""
    Next sld
End Sub

' スライドからテキストを取得する関数
Function GetSlideText(ByVal sld As Slide) As String
    Dim shp As Shape
    Dim shapeInfos() As ShapeInfo
    Dim shapeCount As Integer
    Dim i As Integer, j As Integer
    Dim temp As ShapeInfo
    Dim resultText As String
    
    ' シェイプの数を数える
    shapeCount = 0
    For Each shp In sld.Shapes
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                shapeCount = shapeCount + 1
            End If
        End If
    Next shp
    
    ' 配列を初期化
    If shapeCount > 0 Then
        ReDim shapeInfos(1 To shapeCount)
    Else
        GetSlideText = ""
        Exit Function
    End If
    
    ' シェイプの情報を配列に格納
    i = 1
    For Each shp In sld.Shapes
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                shapeInfos(i).text = shp.TextFrame.TextRange.text
                shapeInfos(i).Top = shp.Top
                shapeInfos(i).Left = shp.Left
                i = i + 1
            End If
        End If
    Next shp
    
    ' バブルソートで並べ替え
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