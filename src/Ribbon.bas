Attribute VB_Name = "Ribbon"
Option Explicit

' 変数の定義
Public startDelay As Double
Public endDelay As Double
Public audioXPosition As Integer
Public transitTime As Double
Public doAllSlides As Boolean
Public doOverride As Boolean
Public useAudioFolder As Boolean
Public processDiff As Boolean
Public Ribbon As IRibbonUI
Public circleXPosition As Integer
Public showAudioIcon As Boolean

' ★追加：除外設定用の変数
Public excludeOutside As Boolean
Public excludeBottom As Boolean
Public bottomThreshold As Double

Private Const InitialShowAudioIcon As Boolean = False
Private Const InitialStartDelay As Double = 2#
Private Const InitialEndDelay As Double = 3#
Private Const InitialAudioXPosition As Integer = -50
Private Const InitialTransitTime As Double = 10#
Private Const InitialDoAllSlides As Boolean = False
Private Const InitialDoOverride As Boolean = True
Private Const InitialUseAudioFolder As Boolean = False
Private Const InitialProcessDiff As Boolean = True
Private Const InitialCircleXPosition As Integer = -50

' ★追加：除外設定用の初期値
Private Const InitialExcludeOutside As Boolean = True
Private Const InitialExcludeBottom As Boolean = True
Private Const InitialBottomThreshold As Double = 10#

Private Const SettingsFileName As String = "settings.txt"

' 設定ファイルのパスを取得する関数
Private Function GetSettingsFilePath() As String
    Dim localAppDataPath As String
    localAppDataPath = Environ("LOCALAPPDATA")

    Dim appFolderPath As String
    appFolderPath = localAppDataPath & "\PPTNaration"
    If Dir(appFolderPath, vbDirectory) = "" Then
        MkDir appFolderPath
    End If

    GetSettingsFilePath = appFolderPath & "\" & SettingsFileName
End Function

' ★新規：内部変数のみを安全に初期化するサブルーチン（UI更新なし）
Private Sub SetDefaultValues()
    startDelay = InitialStartDelay
    endDelay = InitialEndDelay
    audioXPosition = InitialAudioXPosition
    transitTime = InitialTransitTime
    doAllSlides = InitialDoAllSlides
    doOverride = InitialDoOverride
    useAudioFolder = InitialUseAudioFolder
    processDiff = InitialProcessDiff
    circleXPosition = InitialCircleXPosition
    showAudioIcon = InitialShowAudioIcon
    
    ' ★追加
    excludeOutside = InitialExcludeOutside
    excludeBottom = InitialExcludeBottom
    bottomThreshold = InitialBottomThreshold
End Sub

' 初期化コード（破損対策済み）
Sub InitializeVariables()
    ' まずデフォルト値で埋め、破損ファイル読み込み時の0初期化バグを防ぐ
    SetDefaultValues
    
    Dim settingsFilePath As String
    settingsFilePath = GetSettingsFilePath()

    If FileExists(settingsFilePath) Then
        LoadSettings
    Else
        SaveSettings ' 初回起動時はデフォルト値を保存
    End If
End Sub

' ユーザー操作による設定リセット関数
Sub ResetSettings()
    SetDefaultValues
    SaveSettings

    ' リボンUIの更新
    If Not Ribbon Is Nothing Then
        Ribbon.InvalidateControl "circleXPositionDropdown"
        Ribbon.InvalidateControl "showAudioIconBox"
        Ribbon.InvalidateControl "startDelayBox"
        Ribbon.InvalidateControl "endDelayBox"
        Ribbon.InvalidateControl "transitTimeBox"
        Ribbon.InvalidateControl "doAllSlidesBox"
        Ribbon.InvalidateControl "doOverrideBox"
        Ribbon.InvalidateControl "useAudioFolderBox"
        Ribbon.InvalidateControl "processDiffBox"
        Ribbon.InvalidateControl "audioXPositionDropdown"
        ' ★追加
        Ribbon.InvalidateControl "excludeOutsideBox"
        Ribbon.InvalidateControl "excludeBottomBox"
        Ribbon.InvalidateControl "bottomThresholdBox"
    Else
        Call HandleRibbonLoss
    End If

    MsgBox "変数を初期値にリセットしました。", vbInformation
End Sub

Sub OnResetSettings(control As IRibbonControl)
    Dim response As VbMsgBoxResult
    response = MsgBox("変数を初期化しますか?", vbYesNo + vbQuestion, "変数初期化")
    If response = vbYes Then ResetSettings
End Sub

Sub RibbonOnLoad(ribbonUI As IRibbonUI)
    Set Ribbon = ribbonUI
    InitializeVariables
End Sub

' ==========================================
' XML UI コールバック (UI消失時も変数は更新するよう改善)
' ==========================================
Sub OnStartDelayChange(control As IRibbonControl, text As String)
    If IsNumeric(text) Then
        startDelay = CDbl(text)
        SaveSettings
    Else
        MsgBox "有効な数値を入力してください。", vbExclamation
        If Not Ribbon Is Nothing Then Ribbon.InvalidateControl control.id
    End If
End Sub

Sub OnEndDelayChange(control As IRibbonControl, text As String)
    If IsNumeric(text) Then
        endDelay = CDbl(text)
        SaveSettings
    Else
        MsgBox "有効な数値を入力してください。", vbExclamation
        If Not Ribbon Is Nothing Then Ribbon.InvalidateControl control.id
    End If
End Sub

Sub OnTransitTimeChange(control As IRibbonControl, text As String)
    If IsNumeric(text) Then
        transitTime = CDbl(text)
        SaveSettings
    Else
        MsgBox "有効な数値を入力してください。", vbExclamation
        If Not Ribbon Is Nothing Then Ribbon.InvalidateControl control.id
    End If
End Sub

Sub OnDoAllSlidesChange(control As IRibbonControl, pressed As Boolean)
    doAllSlides = pressed
    SaveSettings
End Sub

Sub OnDoOverrideChange(control As IRibbonControl, pressed As Boolean)
    doOverride = pressed
    SaveSettings
End Sub

Sub OnUseAudioFolderChange(control As IRibbonControl, pressed As Boolean)
    useAudioFolder = pressed
    SaveSettings
End Sub

Sub OnProcessDiffChange(control As IRibbonControl, pressed As Boolean)
    processDiff = pressed
    SaveSettings
End Sub

Sub OnAudioXPositionChange(control As IRibbonControl, id As String, index As Integer)
    Select Case id
        Case "pos50": audioXPosition = 50
        Case "pos-50": audioXPosition = -50
        Case "pos-100": audioXPosition = -100
        Case "pos-150": audioXPosition = -150
        Case "pos-200": audioXPosition = -200
        Case "pos-250": audioXPosition = -250
    End Select
    SaveSettings
End Sub

Sub OnCircleXPositionChange(control As IRibbonControl, id As String, index As Integer)
    Select Case id
        Case "circle50":  circleXPosition = 50
        Case "circle-50": circleXPosition = -50
        Case "circle-100": circleXPosition = -100
        Case "circle-150": circleXPosition = -150
        Case "circle-200": circleXPosition = -200
        Case "circle-250": circleXPosition = -250
    End Select
    SaveSettings
End Sub

Sub OnShowAudioIconChange(control As IRibbonControl, pressed As Boolean)
    showAudioIcon = pressed
    SaveSettings
End Sub

' ★追加：テキスト抽出の除外設定コールバック
Sub OnExcludeOutsideChange(control As IRibbonControl, pressed As Boolean)
    excludeOutside = pressed
    SaveSettings
End Sub

Sub OnExcludeBottomChange(control As IRibbonControl, pressed As Boolean)
    excludeBottom = pressed
    SaveSettings
End Sub

Sub OnBottomThresholdChange(control As IRibbonControl, text As String)
    Dim cleanText As String
    cleanText = Replace(text, "%", "")
    
    If IsNumeric(cleanText) Then
        bottomThreshold = CDbl(cleanText)
        SaveSettings
    Else
        MsgBox "有効な数値を入力してください。", vbExclamation
        If Not Ribbon Is Nothing Then Ribbon.InvalidateControl control.id
    End If
End Sub

' ==========================================
' 初期値を取得するコールバック
' ==========================================
Sub GetStartDelay(control As IRibbonControl, ByRef returnedVal)
    returnedVal = startDelay
End Sub

Sub GetEndDelay(control As IRibbonControl, ByRef returnedVal)
    returnedVal = endDelay
End Sub

Sub GetTransitTime(control As IRibbonControl, ByRef returnedVal)
    returnedVal = transitTime
End Sub

Sub GetDoAllSlides(control As IRibbonControl, ByRef returnedVal)
    returnedVal = doAllSlides
End Sub

Sub GetDoOverride(control As IRibbonControl, ByRef returnedVal)
    returnedVal = doOverride
End Sub

Sub GetUseAudioFolder(control As IRibbonControl, ByRef returnedVal)
    returnedVal = useAudioFolder
End Sub

Sub GetProcessDiff(control As IRibbonControl, ByRef returnedVal)
    returnedVal = processDiff
End Sub

Sub GetShowAudioIcon(control As IRibbonControl, ByRef returnedVal)
    returnedVal = showAudioIcon
End Sub

' ★追加
Sub GetExcludeOutside(control As IRibbonControl, ByRef returnedVal)
    returnedVal = excludeOutside
End Sub

Sub GetExcludeBottom(control As IRibbonControl, ByRef returnedVal)
    returnedVal = excludeBottom
End Sub

Sub GetBottomThreshold(control As IRibbonControl, ByRef returnedVal)
    returnedVal = bottomThreshold
End Sub

Sub GetAudioXPositionIndex(control As IRibbonControl, ByRef returnedVal)
    Select Case audioXPosition
        Case 50: returnedVal = 0
        Case -50: returnedVal = 1
        Case -100: returnedVal = 2
        Case -150: returnedVal = 3
        Case -200: returnedVal = 4
        Case -250: returnedVal = 5
    End Select
End Sub

Sub GetCircleXPositionIndex(control As IRibbonControl, ByRef returnedVal)
    Select Case circleXPosition
        Case 50: returnedVal = 0
        Case -50: returnedVal = 1
        Case -100: returnedVal = 2
        Case -150: returnedVal = 3
        Case -200: returnedVal = 4
        Case -250: returnedVal = 5
    End Select
End Sub

' ==========================================
' 設定の保存と読み込み（エラーハンドリング強化）
' ==========================================
Sub SaveSettings()
    On Error GoTo ErrorHandler
    Dim fileNum As Integer
    fileNum = FreeFile
    Dim settingsFilePath As String
    settingsFilePath = GetSettingsFilePath()
    Open settingsFilePath For Output As #fileNum
    Print #fileNum, "StartDelay=" & startDelay
    Print #fileNum, "EndDelay=" & endDelay
    Print #fileNum, "AudioXPosition=" & audioXPosition
    Print #fileNum, "CircleXPosition=" & circleXPosition
    Print #fileNum, "TransitTime=" & transitTime
    Print #fileNum, "DoAllSlides=" & doAllSlides
    Print #fileNum, "DoOverride=" & doOverride
    Print #fileNum, "UseAudioFolder=" & useAudioFolder
    Print #fileNum, "ProcessDiff=" & processDiff
    Print #fileNum, "ShowAudioIcon=" & showAudioIcon
    
    ' ★追加
    Print #fileNum, "ExcludeOutside=" & excludeOutside
    Print #fileNum, "ExcludeBottom=" & excludeBottom
    Print #fileNum, "BottomThreshold=" & bottomThreshold
    
    Close #fileNum
    Exit Sub
ErrorHandler:
    MsgBox "設定の保存中にエラーが発生しました: " & Err.Description, vbExclamation
    On Error Resume Next ' 二重エラー防止
    If fileNum > 0 Then Close #fileNum
End Sub

Sub LoadSettings()
    On Error GoTo ErrorHandler
    Dim fileNum As Integer
    Dim line As String
    Dim parts() As String
    Dim settingsFilePath As String
    settingsFilePath = GetSettingsFilePath()
    fileNum = FreeFile
    Open settingsFilePath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        parts = Split(line, "=")
        If UBound(parts) = 1 Then
            ' 破損データが混ざっていても無視して処理を続けるための保護
            On Error Resume Next
            Select Case parts(0)
                Case "StartDelay": startDelay = CDbl(parts(1))
                Case "EndDelay": endDelay = CDbl(parts(1))
                Case "AudioXPosition": audioXPosition = CInt(parts(1))
                Case "CircleXPosition": circleXPosition = CInt(parts(1))
                Case "TransitTime": transitTime = CDbl(parts(1))
                Case "DoAllSlides": doAllSlides = CBool(parts(1))
                Case "DoOverride": doOverride = CBool(parts(1))
                Case "UseAudioFolder": useAudioFolder = CBool(parts(1))
                Case "ProcessDiff": processDiff = CBool(parts(1))
                Case "ShowAudioIcon": showAudioIcon = CBool(parts(1))
                
                ' ★追加
                Case "ExcludeOutside": excludeOutside = CBool(parts(1))
                Case "ExcludeBottom": excludeBottom = CBool(parts(1))
                Case "BottomThreshold": bottomThreshold = CDbl(parts(1))
            End Select
            On Error GoTo ErrorHandler ' エラーハンドラを戻す
        End If
    Loop
    Close #fileNum
    Exit Sub
ErrorHandler:
    On Error Resume Next
    If fileNum > 0 Then Close #fileNum
End Sub

Function FileExists(FilePath As String) As Boolean
    FileExists = Dir(FilePath) <> ""
End Function

Sub Auto_Exit(ByVal Pres As Presentation)
    SaveSettings
End Sub

Sub Auto_Open()
    InitializeVariables
End Sub

' ==========================================
' スライド操作とその他の機能
' ==========================================
Sub TestPreview(control As IRibbonControl)
    On Error Resume Next
    Application.CommandBars.ExecuteMso "AnimationPreview"
    On Error GoTo 0
End Sub

Sub HandleRibbonLoss()
    Debug.Print "Ribbon オブジェクトが Nothing です。"
    MsgBox "VBAの内部状態によりリボンメニューの表示更新が一時停止しました。" & vbCrLf & _
           "（内部の設定値は正常に変更・保存されています）" & vbCrLf & _
           "表示を元に戻すにはPowerPointを再起動してください。", vbInformation, "リボンの状態"
End Sub

Function GetTargetSlide() As Slide
    On Error Resume Next
    Dim sld As Slide
    If ActiveWindow.Selection.Type = ppSelectionSlides Then
        Set sld = ActiveWindow.Selection.SlideRange(1)
    Else
        Set sld = ActivePresentation.Slides(ActiveWindow.View.Slide.SlideIndex)
    End If
    On Error GoTo 0
    Set GetTargetSlide = sld
End Function

Sub MoveSlideToFirst(control As IRibbonControl)
    ActiveWindow.View.GotoSlide index:=1
End Sub

Sub MoveSlideUp(control As IRibbonControl)
    Dim currentIndex As Integer
    currentIndex = ActiveWindow.View.Slide.SlideIndex
    If currentIndex > 1 Then
        ActiveWindow.View.GotoSlide index:=currentIndex - 1
    End If
End Sub

Sub MoveSlideDown(control As IRibbonControl)
    Dim currentIndex As Integer
    Dim totalSlides As Integer
    currentIndex = ActiveWindow.View.Slide.SlideIndex
    totalSlides = ActivePresentation.Slides.Count
    If currentIndex < totalSlides Then
        ActiveWindow.View.GotoSlide index:=currentIndex + 1
    End If
End Sub

Sub MoveSlideToLast(control As IRibbonControl)
    ActiveWindow.View.GotoSlide index:=ActivePresentation.Slides.Count
End Sub

Sub MoveNextAndPreview(control As IRibbonControl)
    Dim currentIndex As Integer
    Dim totalSlides As Integer
    
    currentIndex = ActiveWindow.View.Slide.SlideIndex
    totalSlides = ActivePresentation.Slides.Count
    
    If currentIndex < totalSlides Then
        ActiveWindow.View.GotoSlide index:=currentIndex + 1
        DoEvents
        On Error Resume Next
        Application.CommandBars.ExecuteMso "AnimationPreview"
        On Error GoTo 0
    Else
        MsgBox "最後のスライドです。これ以上進めません。", vbInformation, "情報"
    End If
End Sub
