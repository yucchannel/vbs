
Option Explicit

Dim folderPath, mp3Files, file, player, randomPlayDuration, randomStopDuration
Dim continueGame, randomNum

' 再生するフォルダーのパスを指定
folderPath = "C:\Path\To\Your\MP3s" ' フォルダーのパスを指定

' 再生時間と停止時間の範囲（秒単位）
Dim minPlayDuration, maxPlayDuration
Dim minStopDuration, maxStopDuration

minPlayDuration = 5  ' 最短再生時間
maxPlayDuration = 15 ' 最長再生時間
minStopDuration = 3  ' 最短停止時間
maxStopDuration = 8  ' 最長停止時間

Set player = CreateObject("WMPlayer.OCX")
Set mp3Files = CreateObject("Scripting.FileSystemObject").GetFolder(folderPath).Files
continueGame = True

If mp3Files.Count = 0 Then
    MsgBox "指定したフォルダーに MP3 ファイルがありません。"
    WScript.Quit
End If

Do While continueGame
    ' ランダムに MP3 ファイルを選択
    For Each file In mp3Files
        If LCase(Right(file.Name, 4)) = ".mp3" Then
            player.URL = file.Path
            player.controls.play
            Exit For
        End If
    Next

    ' 再生時間をランダムに設定
    Randomize
    randomPlayDuration = Int((maxPlayDuration - minPlayDuration + 1) * Rnd + minPlayDuration)
    WScript.Sleep randomPlayDuration * 1000

    ' 再生を停止
    player.controls.stop

    ' 停止時間をランダムに設定
    randomStopDuration = Int((maxStopDuration - minStopDuration + 1) * Rnd + minStopDuration)
    WScript.Sleep randomStopDuration * 1000

    ' ランダムでゲームを続けるか終了するかを判定
    randomNum = Int((2 - 1 + 1) * Rnd + 1) ' 1 または 2 を生成
    If randomNum = 1 Then
        continueGame = False
        MsgBox "爆弾が爆発しました！ゲームオーバー。"
    Else
        MsgBox "セーフ！再生を続けます。"
    End If
Loop

' 後片付け
player.controls.stop
Set player = Nothing

