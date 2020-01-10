Attribute VB_Name = "folder_scan_file_rename"
Sub ScanFiles()
    p = ActiveWorkbook.Path
    fn = Dir("*.*")
    Do While fn <> ""
        Debug.Print fn
        fn = Dir()
    Loop
End Sub
