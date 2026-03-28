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
    ' 各スライドのノートをテキストファイルに抽出するサブルーチン

    Dim sld As Slide
    Dim notesText As String
    Dim textFileName As String
    Dim textFile As Integer
    Dim filePath As String
    Dim appath As String

    If ActivePresentation.Path = "" Then
        MsgBox "プレゼンテーションが保存されていません。一度ファイルを保存してから実行してください。", vbExclamation, "保存確認"
        Exit Sub
    End If

    On Error GoTo ErrorHandler ' エラーハンドラの定義
    textFileName = GetTextFullpath
    
    ' テキストファイルを書き込みモードでオープンする
    textFile = FreeFile
    Open textFileName For Output As #textFile

    ' プレゼンテーションの各スライドをループする
    For Each sld In ActivePresentation.Slides
        ' ノートのテキストを抽出する
        notesText = sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text

        ' スライド番号を取得する
        Dim slideNumber As Long
        slideNumber = sld.slideNumber

        ' スライド番号とノートのテキストをテキストファイルに書き込む
        Print #textFile, "<<< Slide " & slideNumber
        Print #textFile, TrimWhitespace(notesText)
        Print #textFile, "" ' 読みやすさのために空行を追加する
    Next sld

    ' テキストファイルをクローズする
    Close textFile

    Exit Sub ' 正常終了

ErrorHandler:
    ' テキストファイルをクローズする
    If textFile <> 0 Then Close textFile
    MsgBox "エラーが発生しました。エラー番号: " & Err.Number & " エラーの内容: " & Err.Description & " ファイル名: " & textFileName, vbCritical
End Sub

Sub ImportNoteFromText()
    ' テキストファイルからノートを読み込み、各スライドに挿入するサブルーチン

    Dim sld As Slide
    Dim notesText As String
    Dim textFileName As String
    Dim textFile As Integer
    Dim lineText As String
    Dim slideNumber As Long
    Dim filePath As String
    
    On Error GoTo ErrorHandler ' エラーハンドラの定義
    textFileName = GetTextFullpath
    
    If Dir(textFileName) = "" Then
        MsgBox "インポートするテキストファイルが見つかりません。" & vbCrLf & "先に「ノート → テキストファイル」を実行してください。", vbExclamation, "ファイル不在"
        Exit Sub
    End If
    
    ' ファイルを開く
    textFile = FreeFile
    Open textFileName For Input As #textFile

    ' すべての既存のスライドのノートを削除する
    For Each sld In ActivePresentation.Slides
        sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text = ""
    Next sld

    ' テキストファイルの各行を読み込む
    slideNumber = 0 ' スライド番号を初期化
    While Not EOF(textFile)
        Line Input #textFile, lineText ' 行を読み込む
        If InStr(lineText, "<<< Slide ") = 1 Then
            ' スライド番号を抽出する
            slideNumber = CLng(Replace(lineText, "<<< Slide ", ""))
        ElseIf InStr(lineText, "# Slide ") = 1 Then
            ' スライド番号を抽出する
            slideNumber = CLng(Mid(lineText, 9))
        Else
            ' スライドのノートにテキストを追加する
            For Each sld In ActivePresentation.Slides
                If sld.slideNumber = slideNumber Then
                    sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text = sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text & lineText & vbCrLf
                    Exit For
                End If
            Next sld
        End If
    Wend

    ' テキストファイルをクローズする
    Close textFile

    Exit Sub ' 正常終了

ErrorHandler:
    ' テキストファイルをクローズする
    If textFile <> 0 Then Close textFile
    MsgBox "エラーが発生しました。エラー番号: " & Err.Number & " エラーの内容: " & Err.Description & vbCrLf & "ファイル名: " & textFileName, vbCritical
End Sub


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



