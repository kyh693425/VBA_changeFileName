Option Explicit

'select folder
Sub FList_MST_toChangeNm()
    Dim F_Dig As FileDialog
    Dim FS As Scripting.FileSystemObject
    Dim F_Info As Folder
    Dim check As Integer
    Dim check, row As Integer

    With Application
        .ScreenUpdating = False
        EnableEvents = False
        Calculation = xICalculationManual
    End With

    Set F_Dig = Application.FileDialog(msoFileDialogFolderPicker)
    F_Dig.Show

    If F_Dig.SelectedItems.Count > 0 Then
        Row = 2
        Set FS = New Scripting.FileSystemObject
        Set F_Info = FS.GetFolder(F_Dig.SelectedItems(1))
        Call Folder_List_toChangeNm(F_Info)

        With Application
            ScreenUpdating = True
            .EnableEvents = False
            .Calculation = xICalculationManual
        End With
    Else
        Exit Sub
    End If
End Sub

Sub Folder_List_toChangeNm(F_Info As Folder)
    Dim SFList, SFListUp As Folder

    Call File_List_toChangeNm(F_Info)
    Set SFList = F_Info.SubFolders
    For Each SFListUp In SFList
        Call Folder_List_toChangeNm(SFListUp)
    Next SFListUp
End Sub

Sub File_List_toChangeNm(F_Info As Folder)
    Dim fileName As String
    Dim f As File
    Dim sh As Worksheet
    Dim fileListCount, targetFileCount, last_raw, arrExt_Length As Integer
    Dim i As Long
    Dim FileList, arrExt As Variant

    Call ClearContents

    Set sh = ThisWorkbook.Sheets("Tool")
    Set FileList = F_Info.Files
    fileListCount = FileList.Count

    For Each f in FileList
        arrExt = Split(f.Name, ".")
        arrExt_Length = UBound(arrExt) - LBound(arrExt) + 1

        If f.Attributes = (Hidden + System + Archive) Or f.Attributes = (Hidden + Archive) Or f.Attributes = (Hidden + System) Or f.Attributes = Hidden Then
            fileListCount = fileListCount - 1
        ElseIf arrExt_Length = 1 Then
            fileListCount = fileListCount - 1
        Else
            last_raw = sh.Range("M" & Application.Rows.Count).End(xlUp).Row + 1
            sh.Range("M" & last_raw).Value = f.Name
        End If
    Next f
    targetFileCount = sh.Range("N4")

    If flieListCount = 0 Then
        Call AlertMessage(1, F_Info.Name)
    Else
        If Not flieListCount = targetFileCount Then
            Call AlertMessage(2, F_Info.Name)
        End If

        Call Rename_files(sh, F_Info)
    End If
End Sub

Sub Rename_files(sh As Worksheet, F_Info As Folder)
    Dim f As File
    Dim new_name As String

    For Each f in F_Info.Files
        arrExt = Split(f.Name, ".")
        arrExt_Length = UBound(arrExt) - LBound(arrExt) + 1

        If f.Attributes = (Hidden + System + Archive) Or f.Attributes = (Hidden + Archive) Or f.Attributes = (Hidden + System) Or f.Attributes = Hidden Then
            GoTo NextItem
        ElseIf arrExt_Length = 1 Then
            GoTo NextItem
        End If

        new_name = Application.VLookup(f.Name, sh.Range("M:N"), 2, 0)

        If f.Name = new_name Then
            MsgBox "[" & F_Info.Name & "]のフォルダに同じ名前のファイルが既に存在しています。確認してください。"
            GoTo NextItem
        End If

        Application.DisplayAlerts = False
        f.Name = new_name

 NextItem:
    Next f

    MsgBox "[" & F_Info.Name & "]のファイル名を変更しました。"
    Call OpenExplorer(F_Info.Path)
End Sub

Sub OpenExplorer(target As String)
    Call Shell("explorer.exe" & "" & target, vbNormalFocus)
End Sub

Sub ClearContents()
    Dim sh As Worksheet

    Set sh = ThisWorkbook.Sheets("Tool")
    sh.Range("M6:M30").ClearContents
End Sub

Sub AlertMessage(flg As Integer, shtNm As String)
    If flg = 1 Then
        MsgBox "[" & shtNm & "]のフォルダからpicファイルを見付かりませんでした。"
    End If

    If flg = 2 Then
        MsgBox "[" & shtNm & "]のフォルダのファイル数を確認してください。"
    End If
End Sub