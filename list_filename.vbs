Option Explicit

' プロンプトでフォルダパスと拡張子を入力
Dim folderPath, fileExtension, outputFilePath
folderPath = InputBox("フォルダのパスを入力してください")

If folderPath = "" Then
    WScript.Quit
End If

fileExtension = InputBox("拡張子を入力してください 例 txt")
If fileExtension = "" Then
    WScript.Quit
End If

' 拡張子にドットが含まれていない場合は追加
If Left(fileExtension, 1) <> "." Then
    fileExtension = "." & fileExtension
End If

' 出力ファイルのパスを設定
outputFilePath = folderPath & "\file_list.txt"

Dim fso, folder, files, file, fileList, outputFile
Set fso = CreateObject("Scripting.FileSystemObject")

' フォルダの存在を確認
If Not fso.FolderExists(folderPath) Then
    MsgBox "指定したフォルダが存在しません", vbExclamation
    WScript.Quit
End If

Set folder = fso.GetFolder(folderPath)
Set files = folder.Files

' ファイルリストを作成
fileList = ""
For Each file In files
    If LCase(fso.GetExtensionName(file)) = LCase(Mid(fileExtension, 2)) Then
        fileList = fileList & file.Name & vbCrLf
    End If
Next

' ファイル名を昇順に並べ替え
Dim fileArray
fileArray = Split(Trim(fileList), vbCrLf)
If UBound(fileArray) >= 0 Then
    Call QuickSort(fileArray, LBound(fileArray), UBound(fileArray))
End If

' 出力ファイルに書き込み
Set outputFile = fso.CreateTextFile(outputFilePath, True)
For Each file In fileArray
    outputFile.WriteLine file
Next
outputFile.Close

' テキストファイルを表示
Dim shell
Set shell = CreateObject("WScript.Shell")
shell.Run "notepad.exe " & outputFilePath

' クイックソートの実装
Sub QuickSort(arr, low, high)
    Dim i, j, pivot, temp
    If low < high Then
        pivot = arr((low + high) \ 2)
        i = low
        j = high
        Do While i <= j
            Do While arr(i) < pivot
                i = i + 1
            Loop
            Do While arr(j) > pivot
                j = j - 1
            Loop
            If i <= j Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop
        Call QuickSort(arr, low, j)
        Call QuickSort(arr, i, high)
    End If
End Sub