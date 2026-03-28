Attribute VB_Name = "UtilFilePath"
Option Explicit

' API関数の宣言
Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As LongPtr) As LongPtr
Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As LongPtr

Const GHND = &H42
Const CF_UNICODETEXT = 13

Sub CopyToClipboard(text As String)
    Dim hGlobalMemory As LongPtr
    Dim lpGlobalMemory As LongPtr
    Dim hwnd As LongPtr
    Dim hClipboardData As LongPtr
    Dim textLength As Long

    ' クリップボードを開く
    If OpenClipboard(0&) Then
        ' クリップボードを空にする
        EmptyClipboard

        ' テキストの長さを取得（ワイド文字として）
        textLength = (Len(text) + 1) * 2 ' +1 はNULL文字のため

        ' グローバルメモリを確保する
        hGlobalMemory = GlobalAlloc(GHND, textLength)
        
        ' グローバルメモリをロックする
        lpGlobalMemory = GlobalLock(hGlobalMemory)

        ' メモリにテキストをコピーする
        lstrcpyW lpGlobalMemory, StrPtr(text)
        
        ' グローバルメモリのロックを解除する
        GlobalUnlock hGlobalMemory

        ' クリップボードにデータを設定する
        SetClipboardData CF_UNICODETEXT, hGlobalMemory

        ' クリップボードを閉じる
        CloseClipboard
    End If
End Sub

Function GetTextFullpath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim filePath As String
    Dim textFileName As String

    filePath = OneDriveUrlToLocalPath(ActivePresentation.Path)
    textFileName = fso.GetBaseName(ActivePresentation.FullName) & ".txt"
    
    GetTextFullpath = filePath & "\" & textFileName
End Function

Function GetTextBasepath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim filePath As String

    GetTextBasepath = fso.GetBaseName(ActivePresentation.FullName)
End Function

Function GetTextFldrpath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim filePath As String

    GetTextFldrpath = OneDriveUrlToLocalPath(ActivePresentation.Path)
End Function

Function TrimWhitespace(inputText As String) As String
    ' テキストの前後の空白や改行を除去する関数
    
    ' 前後の空白を除去
    inputText = LTrim(RTrim(inputText))
    
    ' 前後の改行を除去
    Do While Len(inputText) > 0 And (Left(inputText, 1) = vbCr Or Left(inputText, 1) = vbLf)
        inputText = Mid(inputText, 2)
    Loop
    
    Do While Len(inputText) > 0 And (Right(inputText, 1) = vbCr Or Right(inputText, 1) = vbLf)
        inputText = Left(inputText, Len(inputText) - 1)
    Loop
    
    TrimWhitespace = inputText
End Function

