Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Private Declare Function InternetReadBinaryFile Lib "wininet.dll" Alias "InternetReadFile" (ByVal hfile As Long, ByRef bytearray_firstelement As Byte, ByVal lNumBytesToRead As Long, ByRef lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sUrl As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
 
'original source >> https://analystcave.com/excel-downloading-files-using-vba/

Sub DownloadFile(sUrl As String, filePath As String, Optional overWriteFile As Boolean)
  Dim hInternet, hSession, lngDataReturned As Long, sBuffer() As Byte, totalRead As Long
  Const bufSize = 256
  ReDim sBuffer(bufSize)
  hSession = InternetOpen("", 0, vbNullString, vbNullString, 0)
  If hSession Then hInternet = InternetOpenUrl(hSession, sUrl, vbNullString, 0, INTERNET_FLAG_NO_CACHE_WRITE, 0)
  Set oStream = CreateObject("ADODB.Stream")
  oStream.Open
  oStream.Type = 1
 
  If hInternet Then
    iReadFileResult = InternetReadBinaryFile(hInternet, sBuffer(0), UBound(sBuffer) - LBound(sBuffer), lngDataReturned)
    ReDim Preserve sBuffer(lngDataReturned - 1)
    oStream.Write sBuffer
    ReDim sBuffer(bufSize)
    totalRead = totalRead + lngDataReturned
    Application.StatusBar = "Downloading file. " & CLng(totalRead / 1024) & " KB downloaded"
    DoEvents
 
    Do While lngDataReturned <> 0
      iReadFileResult = InternetReadBinaryFile(hInternet, sBuffer(0), UBound(sBuffer) - LBound(sBuffer), lngDataReturned)
      If lngDataReturned = 0 Then Exit Do
 
      ReDim Preserve sBuffer(lngDataReturned - 1)
      oStream.Write sBuffer
      ReDim sBuffer(bufSize)
      totalRead = totalRead + lngDataReturned
      Application.StatusBar = "Downloading file. " & CLng(totalRead / 1024) & " KB downloaded"
      DoEvents
    Loop
 
    Application.StatusBar = "Download complete"
    oStream.SaveToFile filePath, IIf(overWriteFile, 2, 1)
    oStream.Close
  End If
  Call InternetCloseHandle(hInternet)
End Sub
