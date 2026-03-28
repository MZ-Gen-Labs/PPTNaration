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

Public hideAudioIcon As Boolean
Private Const InitialHideAudioIcon As Boolean = False

Private sysPrompt1Text As String
Private sysPrompt2Text As String
Private sysPrompt3Text As String

Private Const InitialStartDelay As Double = 2#
Private Const InitialEndDelay As Double = 3#
Private Const InitialAudioXPosition As Integer = -50
Private Const InitialTransitTime As Double = 10#
Private Const InitialDoAllSlides As Boolean = False
Private Const InitialDoOverride As Boolean = True
Private Const InitialUseAudioFolder As Boolean = False
Private Const InitialProcessDiff As Boolean = True
Private Const SettingsFileName As String = "xmal_settings.txt"
Private Const InitialCircleXPosition As Integer = -50


' 設定ファイルのパスを取得する関数
Private Function GetSettingsFilePath() As String
    Dim localAppDataPath As String
    localAppDataPath = Environ("LOCALAPPDATA")

    ' xmalアドイン用のフォルダを作成（存在しない場合）
    Dim xmalFolderPath As String
    xmalFolderPath = localAppDataPath & "\xmal_addin"
    If Dir(xmalFolderPath, vbDirectory) = "" Then
        MkDir xmalFolderPath
    End If

    GetSettingsFilePath = xmalFolderPath & "\" & SettingsFileName
End Function

' 初期化コード
Sub InitializeVariables()
    Dim settingsFilePath As String
    settingsFilePath = GetSettingsFilePath()

    If FileExists(settingsFilePath) Then
        LoadSettings
    Else
        ResetSettings
    End If
End Sub

' 設定をリセットする関数
Sub ResetSettings()
    startDelay = InitialStartDelay
    endDelay = InitialEndDelay
    audioXPosition = InitialAudioXPosition
    transitTime = InitialTransitTime
    doAllSlides = InitialDoAllSlides
    doOverride = InitialDoOverride
    useAudioFolder = InitialUseAudioFolder
    processDiff = InitialProcessDiff
    circleXPosition = InitialCircleXPosition
    Ribbon.InvalidateControl "circleXPositionDropdown"
    hideAudioIcon = InitialHideAudioIcon
    Ribbon.InvalidateControl "hideAudioIconBox"

    SaveSettings

    ' リボンUIの更新
    If Not Ribbon Is Nothing Then
        Ribbon.InvalidateControl "startDelayBox"
        Ribbon.InvalidateControl "endDelayBox"
        Ribbon.InvalidateControl "transitTimeBox"
        Ribbon.InvalidateControl "doAllSlidesBox"
        Ribbon.InvalidateControl "doOverrideBox"
        Ribbon.InvalidateControl "useAudioFolderBox"
        Ribbon.InvalidateControl "processDiffBox"
        Ribbon.InvalidateControl "audioXPositionDropdown"
    Else
        Debug.Print "Ribbon オブジェクトが Nothing のため、リボンの更新をスキップします。"
        Call GetRibbonObject ' 再初期化を試みる
        If Not Ribbon Is Nothing Then
            Ribbon.InvalidateControl "startDelayBox"
            Ribbon.InvalidateControl "endDelayBox"
            Ribbon.InvalidateControl "transitTimeBox"
            Ribbon.InvalidateControl "doAllSlidesBox"
            Ribbon.InvalidateControl "doOverrideBox"
            Ribbon.InvalidateControl "useAudioFolderBox"
            Ribbon.InvalidateControl "processDiffBox"
            Ribbon.InvalidateControl "audioXPositionDropdown"
        Else
            MsgBox "リボンオブジェクトの再取得に失敗しました。", vbCritical
        End If
    End If

    MsgBox "Settings have been reset to default values.", vbInformation
    Debug.Print "Settings reset to default values"
End Sub

' リボンのリセットボタンのコールバック関数
Sub OnResetSettings(control As IRibbonControl)
    Dim response As VbMsgBoxResult
    response = MsgBox("変数を初期化しますか?", vbYesNo + vbQuestion, "変数初期化")

    If response = vbYes Then
        ResetSettings
    End If
End Sub

' リボンが読み込まれたときのイベント
Sub RibbonOnLoad(ribbonUI As IRibbonUI)
    Set Ribbon = ribbonUI
    InitializeVariables
End Sub

' XML UI コールバック

Sub OnStartDelayChange(control As IRibbonControl, text As String)
    If Ribbon Is Nothing Then Call HandleRibbonLoss: Exit Sub
    If IsNumeric(text) Then
        startDelay = CDbl(text)
        SaveSettings  ' 設定変更時に即座に保存
    Else
        MsgBox "Please enter a valid number.", vbExclamation
        Ribbon.InvalidateControl control.id
    End If
End Sub

Sub OnEndDelayChange(control As IRibbonControl, text As String)
    If Ribbon Is Nothing Then Call HandleRibbonLoss: Exit Sub
    If IsNumeric(text) Then
        endDelay = CDbl(text)
        SaveSettings  ' 設定変更時に即座に保存
    Else
        MsgBox "Please enter a valid number.", vbExclamation
        Ribbon.InvalidateControl control.id
    End If
End Sub

Sub OnTransitTimeChange(control As IRibbonControl, text As String)
    If Ribbon Is Nothing Then Call HandleRibbonLoss: Exit Sub
    If IsNumeric(text) Then
        transitTime = CDbl(text)
        SaveSettings  ' 設定変更時に即座に保存
    Else
        MsgBox "Please enter a valid number.", vbExclamation
        Ribbon.InvalidateControl control.id
    End If
End Sub

Sub OnDoAllSlidesChange(control As IRibbonControl, pressed As Boolean)
    doAllSlides = pressed
    SaveSettings  ' 設定変更時に即座に保存
End Sub

Sub OnDoOverrideChange(control As IRibbonControl, pressed As Boolean)
    doOverride = pressed
    SaveSettings  ' 設定変更時に即座に保存
End Sub

Sub OnUseAudioFolderChange(control As IRibbonControl, pressed As Boolean)
    useAudioFolder = pressed
    SaveSettings  ' 設定変更時に即座に保存
End Sub

Sub OnProcessDiffChange(control As IRibbonControl, pressed As Boolean)
    processDiff = pressed
    SaveSettings  ' 設定変更時に即座に保存
End Sub

Sub OnAudioXPositionChange(control As IRibbonControl, id As String, index As Integer)
    If Ribbon Is Nothing Then Call HandleRibbonLoss: Exit Sub
    Select Case id
        Case "pos50"
            audioXPosition = 50
        Case "pos-50"
            audioXPosition = -50
        Case "pos-100"
            audioXPosition = -100
        Case "pos-150"
            audioXPosition = -150
        Case "pos-200"
            audioXPosition = -200
        Case "pos-250"
            audioXPosition = -250
    End Select
    SaveSettings  ' 設定変更時に即座に保存
End Sub


' 初期値を取得するコールバック
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

Sub GetAudioXPositionIndex(control As IRibbonControl, ByRef returnedVal)
    Select Case audioXPosition
        Case 50
            returnedVal = 0
        Case -50
            returnedVal = 1
        Case -100
            returnedVal = 2
        Case -150
            returnedVal = 3
        Case -200
            returnedVal = 4
        Case -250
            returnedVal = 5
    End Select
End Sub

' 変数を表示するコマンド
Sub ShowVariables(control As IRibbonControl)
    MsgBox "Start Delay: " & startDelay & vbCrLf & _
           "End Delay: " & endDelay & vbCrLf & _
           "Audio X Position: " & audioXPosition & vbCrLf & _
           "Transit Time: " & transitTime & vbCrLf & _
           "Do All Slides: " & doAllSlides & vbCrLf & _
           "Do Override: " & doOverride & vbCrLf & _
           "Use Audio Folder: " & useAudioFolder & vbCrLf & _
           "Process Diff Text: " & processDiff
End Sub

' 設定を保存する関数
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
    Print #fileNum, "HideAudioIcon=" & hideAudioIcon
    Close #fileNum
'    MsgBox "Settings saved successfully to " & settingsFilePath, vbInformation
    Debug.Print "Settings saved successfully to " & settingsFilePath
    Exit Sub
ErrorHandler:
    MsgBox "Error saving settings: " & Err.Description, vbCritical
    Debug.Print "Error saving settings: " & Err.Description
    If fileNum > 0 Then Close #fileNum
End Sub

' 設定を読み込む関数
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
            Select Case parts(0)
                Case "StartDelay"
                    startDelay = CDbl(parts(1))
                Case "EndDelay"
                    endDelay = CDbl(parts(1))
                Case "AudioXPosition"
                    audioXPosition = CInt(parts(1))
                Case "CircleXPosition"
                    circleXPosition = CInt(parts(1))
                Case "TransitTime"
                    transitTime = CDbl(parts(1))
                Case "DoAllSlides"
                    doAllSlides = CBool(parts(1))
                Case "DoOverride"
                    doOverride = CBool(parts(1))
                Case "UseAudioFolder"
                    useAudioFolder = CBool(parts(1))
                Case "ProcessDiff"
                    processDiff = CBool(parts(1))
                Case "HideAudioIcon"
                    hideAudioIcon = CBool(parts(1))
            End Select
        End If
    Loop
    Close #fileNum
    Debug.Print "Settings loaded successfully"
    Exit Sub
ErrorHandler:
    MsgBox "Error loading settings: " & Err.Description, vbCritical
    Debug.Print "Error loading settings: " & Err.Description
    If fileNum > 0 Then Close #fileNum
End Sub

' ファイルが存在するかチェックする関数
Function FileExists(filePath As String) As Boolean
    FileExists = Dir(filePath) <> ""
End Function

' PowerPointが終了する際に呼び出される関数
Sub Auto_Exit(ByVal Pres As Presentation)
    SaveSettings
    Debug.Print "Auto_Exit called, settings saved"
End Sub

' アドインが読み込まれたときに呼び出される関数
Sub Auto_Open()
    InitializeVariables
    Debug.Print "Auto_Open called, variables initialized"
End Sub

' アニメーション-プレビューの実行
Sub TestPreview()
    On Error Resume Next
    Application.CommandBars.ExecuteMso "AnimationPreview"
    On Error GoTo 0
End Sub

' リボンオブジェクトを取得する処理
Sub GetRibbonObject()
    Dim objAddIn As COMAddIn

    On Error Resume Next
    Set objAddIn = Application.COMAddIns("PPTGenVoice2.ThisAddIn") ' あなたのアドインの ProgID
    On Error GoTo 0

    If Not objAddIn Is Nothing Then
        ' Ribbon オブジェクトを再取得
        On Error Resume Next
        Set Ribbon = objAddIn.Object.GetRibbonUI() ' IRibbonExtensibility インターフェースの GetRibbonUI メソッドを使用
        On Error GoTo 0
        If Ribbon Is Nothing Then
            Debug.Print "IRibbonUI オブジェクトの再取得に失敗しました。"
        End If
    Else
        Debug.Print "COMAddIn オブジェクトの取得に失敗しました。"
    End If
End Sub

' リボンオブジェクトが Nothing だった場合の共通処理
Sub HandleRibbonLoss()
    Debug.Print "Ribbon オブジェクトが Nothing です。再初期化を試みます。"
    Call GetRibbonObject ' Ribbon オブジェクトを再取得
    If Not Ribbon Is Nothing Then
        InitializeVariables ' 初期化処理を再度実行
    Else
        MsgBox "リボンオブジェクトの再取得に失敗しました。", vbCritical
    End If
End Sub

Sub OnCircleXPositionChange(control As IRibbonControl, id As String, index As Integer)
    If Ribbon Is Nothing Then Call HandleRibbonLoss: Exit Sub
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

' 現在表示・選択されているスライドを安全に取得する共通関数
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

' スライドをトップ（先頭）に移動
Sub MoveSlideToFirst(control As IRibbonControl)
    ActiveWindow.View.GotoSlide index:=1
End Sub

' スライドを一つ前に移動
Sub MoveSlideUp(control As IRibbonControl)
    Dim currentIndex As Integer
    currentIndex = ActiveWindow.View.Slide.SlideIndex
    If currentIndex > 1 Then
        ActiveWindow.View.GotoSlide index:=currentIndex - 1
    End If
End Sub

' スライドを一つ後に移動
Sub MoveSlideDown(control As IRibbonControl)
    Dim currentIndex As Integer
    Dim totalSlides As Integer
    currentIndex = ActiveWindow.View.Slide.SlideIndex
    totalSlides = ActivePresentation.Slides.Count
    If currentIndex < totalSlides Then
        ActiveWindow.View.GotoSlide index:=currentIndex + 1
    End If
End Sub

' スライドを最後に移動
Sub MoveSlideToLast(control As IRibbonControl)
    ActiveWindow.View.GotoSlide index:=ActivePresentation.Slides.Count
End Sub

Sub OnHideAudioIconChange(control As IRibbonControl, pressed As Boolean)
    hideAudioIcon = pressed
    SaveSettings  ' 設定変更時に即座に保存
End Sub

Sub GetHideAudioIcon(control As IRibbonControl, ByRef returnedVal)
    returnedVal = hideAudioIcon
End Sub

' 次のスライドへ移動し、直後にプレビューを実行する
Sub MoveNextAndPreview(control As IRibbonControl)
    Dim currentIndex As Integer
    Dim totalSlides As Integer
    
    ' 現在のスライド番号と全スライド数を取得
    currentIndex = ActiveWindow.View.Slide.SlideIndex
    totalSlides = ActivePresentation.Slides.Count
    
    ' 最後のスライドでなければ処理を実行
    If currentIndex < totalSlides Then
        ' 1. 次のスライドへ移動
        ActiveWindow.View.GotoSlide index:=currentIndex + 1
        
        ' 2. PPTの画面描画（スライド切り替え）が完了するのを一瞬待つ
        DoEvents
        
        ' 3. プレビューコマンドを実行
        On Error Resume Next
        Application.CommandBars.ExecuteMso "AnimationPreview"
        On Error GoTo 0
    Else
        MsgBox "最後のスライドです。これ以上進めません。", vbInformation, "情報"
    End If
End Sub