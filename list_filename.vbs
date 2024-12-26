Option Explicit

' �v�����v�g�Ńt�H���_�p�X�Ɗg���q�����
Dim folderPath, fileExtension, outputFilePath
folderPath = InputBox("�t�H���_�̃p�X����͂��Ă�������")

If folderPath = "" Then
    WScript.Quit
End If

fileExtension = InputBox("�g���q����͂��Ă������� �� txt")
If fileExtension = "" Then
    WScript.Quit
End If

' �g���q�Ƀh�b�g���܂܂�Ă��Ȃ��ꍇ�͒ǉ�
If Left(fileExtension, 1) <> "." Then
    fileExtension = "." & fileExtension
End If

' �o�̓t�@�C���̃p�X��ݒ�
outputFilePath = folderPath & "\file_list.txt"

Dim fso, folder, files, file, fileList, outputFile
Set fso = CreateObject("Scripting.FileSystemObject")

' �t�H���_�̑��݂��m�F
If Not fso.FolderExists(folderPath) Then
    MsgBox "�w�肵���t�H���_�����݂��܂���", vbExclamation
    WScript.Quit
End If

Set folder = fso.GetFolder(folderPath)
Set files = folder.Files

' �t�@�C�����X�g���쐬
fileList = ""
For Each file In files
    If LCase(fso.GetExtensionName(file)) = LCase(Mid(fileExtension, 2)) Then
        fileList = fileList & file.Name & vbCrLf
    End If
Next

' �t�@�C�����������ɕ��בւ�
Dim fileArray
fileArray = Split(Trim(fileList), vbCrLf)
If UBound(fileArray) >= 0 Then
    Call QuickSort(fileArray, LBound(fileArray), UBound(fileArray))
End If

' �o�̓t�@�C���ɏ�������
Set outputFile = fso.CreateTextFile(outputFilePath, True)
For Each file In fileArray
    outputFile.WriteLine file
Next
outputFile.Close

' �e�L�X�g�t�@�C����\��
Dim shell
Set shell = CreateObject("WScript.Shell")
shell.Run "notepad.exe " & outputFilePath

' �N�C�b�N�\�[�g�̎���
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