Sub CusDownload()

    Dim bates As Long
    Dim fileTerdownload As Long
    Dim wildcardLink As String
    Dim wildcardLinkLoop As String
    Dim nameFileAja As String
    Dim changeWildcardWith As String
    Dim destinasinya As String

    fileTerdownload = 0
    'wildcardLink = Cells(1, 6).Value
    destinasinya = Cells(2, 6).Value
    bates = Cells(3, 6).Value
    bates = bates + 4

    For i = 5 To bates
        wildcardLinkLoop = Cells(i, 2).Value
        nameFileAja = Cells(i, 3).Value
        DownloadFile wildcardLinkLoop, destinasinya & nameFileAja, True
        fileTerdownload = fileTerdownload + 1
    Next

    MsgBox ("Done! File terdownload: " & fileTerdownload)

End Sub
