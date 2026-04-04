Attribute VB_Name = "MdlAudioTreat"
Option Explicit

Enum operationType
    AddOperation
    ChangeOperation
    RemoveOperation
    DeleteOperation
End Enum

Sub AddAudioToSlides()
    Dim sld As Slide
    Dim slds As SlideRange
    
    If ActivePresentation.Path = "" Then
        MsgBox "プレゼンテーションが保存されていません。一度ファイルを保存してから実行してください。", vbExclamation, "保存確認"
        Exit Sub
    End If

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
        AddAudioToSlide sld
        AddAutoTransitToSlide sld
        TreattransitOnSlide sld, AddOperation
    Next sld
End Sub

Sub RemoveAudioFromSlides()
    Dim sld As Slide
    For Each sld In ActivePresentation.Slides
        RemoveAudioFromSlide sld
    Next sld
End Sub

Sub MoveAudioInSlides()
    Dim sld As Slide
    Dim slds As SlideRange
    
    If ActivePresentation.Path = "" Then
        MsgBox "プレゼンテーションが保存されていません。一度ファイルを保存してから実行してください。", vbExclamation, "保存確認"
        Exit Sub
    End If

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
        MoveAudioInSlide sld
        AddAutoTransitToSlide sld
        TreattransitOnSlide sld, AddOperation
    Next sld
End Sub

' スライドへの音声配置処理
Sub AddAudioToSlide(sld As Slide)
    Dim shp As Shape
    Dim FilePath As String
    Dim audioFile As String
    Dim slideNumber As Long
    Dim effect As effect
    Dim presentationName As String

    On Error GoTo ErrorHandler ' メディア挿入時の不測のクラッシュを防止

    slideNumber = sld.slideNumber

    If useAudioFolder Then
        FilePath = OneDriveUrlToLocalPath(ActivePresentation.Path) & "\audio\"
    Else
        presentationName = Left(ActivePresentation.Name, InStrRev(ActivePresentation.Name, ".") - 1)
        FilePath = OneDriveUrlToLocalPath(ActivePresentation.Path) & "\" & presentationName & "\"
    End If

    If Dir(FilePath, vbDirectory) = "" Then
        Exit Sub
    End If

    ' ★拡張：よく使われる音声フォーマットに幅広く対応
    audioFile = ""
    If Dir(FilePath & slideNumber & ".wav") <> "" Then
        audioFile = FilePath & slideNumber & ".wav"
    ElseIf Dir(FilePath & slideNumber & ".mp3") <> "" Then
        audioFile = FilePath & slideNumber & ".mp3"
    ElseIf Dir(FilePath & slideNumber & ".m4a") <> "" Then
        audioFile = FilePath & slideNumber & ".m4a"
    ElseIf Dir(FilePath & slideNumber & ".wma") <> "" Then
        audioFile = FilePath & slideNumber & ".wma"
    End If

    If audioFile <> "" Then
        If doOverride Then
            RemoveAudioFromSlide sld
        End If
        
        ' スライドに音声ファイルを追加する
        Set shp = sld.Shapes.AddMediaObject2(audioFile, msoFalse, msoTrue, sld.Master.Width + audioXPosition, sld.Master.Height - 50)

        ' 音声オブジェクトに確実にタグを設定する
        shp.Tags.Add Name:="AudioObject", Value:="True"

        ' アニメーションを「非表示にする前」に確実に登録する
        Set effect = sld.TimeLine.MainSequence.AddEffect(shp, msoAnimEffectMediaPlay, Trigger:=msoAnimTriggerWithPrevious)
        effect.Timing.TriggerDelayTime = startDelay
        
        ' 最後に表示/非表示を切り替える
        If showAudioIcon Then
            shp.AnimationSettings.PlaySettings.HideWhileNotPlaying = msoFalse
        Else
            shp.AnimationSettings.PlaySettings.HideWhileNotPlaying = msoTrue
        End If
    End If
    Exit Sub

ErrorHandler:
    MsgBox "スライド " & slideNumber & " への音声追加中にエラーが発生しました。ファイル形式が未対応か壊れている可能性があります。" & vbCrLf & Err.Description, vbExclamation, "メディア挿入エラー"
End Sub

Sub AddAutoTransitToSlide(sld As Slide)
    sld.SlideShowTransition.AdvanceOnTime = msoTrue
    sld.SlideShowTransition.AdvanceTime = transitTime
End Sub

Sub RemoveAudioFromSlide(sld As Slide)
    Dim shp As Shape
    Dim i As Long
    For i = sld.Shapes.Count To 1 Step -1
        Set shp = sld.Shapes(i)
        
        If shp.Type = msoMedia Then
            If shp.Tags.Item("AudioObject") = "True" Then
                shp.Delete
                GoTo NextShape
            End If
        End If

        If shp.AutoShapeType = msoShapeOval Then
            If shp.Tags.Item("AudioControl") = "True" Then
                shp.Delete
                GoTo NextShape
            End If
        End If
NextShape:
    Next i
End Sub

Sub MoveAudioInSlide(sld As Slide)
    Dim shp As Shape
    Dim eff As effect
    
    For Each shp In sld.Shapes
        If shp.Type = msoMedia Then
            If shp.Tags.Item("AudioObject") = "True" Then
                ' 1. 位置の再設定
                shp.Left = sld.Master.Width + audioXPosition
                shp.Top = sld.Master.Height - 50

                ' 2. 音声アイコンの表示/非表示の再設定
                If showAudioIcon Then
                    shp.AnimationSettings.PlaySettings.HideWhileNotPlaying = msoFalse
                Else
                    shp.AnimationSettings.PlaySettings.HideWhileNotPlaying = msoTrue
                End If
                
                ' 3. 開始遅延の再設定
                For Each eff In sld.TimeLine.MainSequence
                    If eff.Shape.Name = shp.Name And eff.EffectType = msoAnimEffectMediaPlay Then
                        eff.Timing.TriggerDelayTime = startDelay
                    End If
                Next eff
            End If
        End If
    Next shp
End Sub

Sub MoveAudioPosition(x As Integer, y As Integer)
    Dim sld As Slide
    Dim shp As Shape
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Type = msoMedia Then
                If shp.Tags.Item("AudioObject") = "True" Then
                    shp.Left = sld.Master.Width + x
                    shp.Top = sld.Master.Height + y
                End If
            End If
        Next shp
    Next sld
End Sub

Sub MakeVideoTransparent(shp As Shape)
    On Error Resume Next
    shp.Fill.Transparency = 1
    shp.line.Transparency = 1
    On Error GoTo 0
End Sub

Sub MakeAllVideosTransparent()
    Dim sld As Slide
    Dim shp As Shape
    ' ★修正：エラーで処理全体が止まるのを防ぐため、ループ内でResume Nextを使用
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Type = msoMedia Then
                If shp.Tags.Item("AudioObject") = "True" Then
                    On Error Resume Next
                    shp.Fill.Transparency = 1
                    On Error GoTo 0
                End If
            End If
        Next shp
    Next sld
End Sub

Sub MakeAudioTransparent(shp As Shape, Optional transparencyLevel As Single = 1)
    On Error Resume Next
    If shp.Type = msoMedia Then
        If transparencyLevel = 1 Then
            shp.Fill.Visible = msoFalse
        Else
            shp.Fill.Transparency = transparencyLevel
            shp.line.Transparency = transparencyLevel
        End If
    End If
    On Error GoTo 0
End Sub

Sub TreattransitOnSlide(sld As Slide, optype As operationType)
    Dim shpcnt As Integer
    Dim shp As Shape
    Dim eff As effect
    shpcnt = 0
    
    For Each shp In sld.Shapes
        If shp.AutoShapeType = msoShapeOval Then
            If shp.Tags.Item("AudioControl") = "True" Then
                GoTo AnimationProcess
            End If
        End If
        GoTo NextShape
AnimationProcess:
        Dim i As Integer
        Dim effect As effect
        Select Case optype
            Case AddOperation, ChangeOperation
                If shp.AnimationSettings.Animate = msoTrue Then
                    For i = sld.TimeLine.MainSequence.Count To 1 Step -1
                        Set effect = sld.TimeLine.MainSequence(i)
                        If effect.Shape.Name = shp.Name Then
                            sld.TimeLine.MainSequence(i).Delete
                        End If
                    Next i
                End If
                
                Set eff = sld.TimeLine.MainSequence.AddEffect(Shape:=shp, effectId:=msoAnimEffectSplit)
                With eff.Timing
                    .Duration = endDelay
                    .TriggerType = msoAnimTriggerAfterPrevious
                End With
                
                shpcnt = shpcnt + 1
            Case RemoveOperation
                If shp.AnimationSettings.Animate = msoTrue Then
                    For i = sld.TimeLine.MainSequence.Count To 1 Step -1
                        Set effect = sld.TimeLine.MainSequence(i)
                        If effect.Shape.Name = shp.Name Then
                            sld.TimeLine.MainSequence(i).Delete
                        End If
                    Next i
                End If
            Case DeleteOperation
                shp.Delete
        End Select
NextShape:
    Next shp
    
    If (optype = AddOperation) And (shpcnt = 0) Then
        Dim posX As Single
        Dim posY As Single
        posX = sld.Master.Width + circleXPosition
        posY = sld.Master.Height - 50
    
        Set shp = sld.Shapes.AddShape(msoShapeOval, posX, posY, 50, 50)
        shp.Tags.Add Name:="AudioControl", Value:="True"
        shp.Fill.Transparency = 1
        shp.line.Transparency = 1

        If shp.AnimationSettings.Animate = msoTrue Then
            shp.AnimationSettings.Animate = msoFalse
        End If
        
        Set eff = sld.TimeLine.MainSequence.AddEffect(Shape:=shp, effectId:=msoAnimEffectSplit)
        With eff.Timing
            .Duration = endDelay
            .TriggerType = msoAnimTriggerAfterPrevious
        End With
    End If
    
    Select Case optype
        Case AddOperation
            sld.SlideShowTransition.AdvanceOnTime = msoTrue
            sld.SlideShowTransition.AdvanceTime = transitTime
        Case ChangeOperation
            sld.SlideShowTransition.AdvanceOnTime = msoTrue
            sld.SlideShowTransition.AdvanceTime = transitTime
        Case RemoveOperation
            sld.SlideShowTransition.AdvanceOnTime = msoFalse
        Case DeleteOperation
            sld.SlideShowTransition.AdvanceOnTime = msoFalse
    End Select
End Sub
