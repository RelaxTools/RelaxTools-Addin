VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FindFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const INVALID_HANDLE_VALUE      As Long = (-1)              '無効なFile Handle値

'ファイル属性
Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10               'フォルダ

Const MAX_PATH                  As Long = 260&              'パスの最大長
Const conUnicodeMaxPath         As Long = MAX_PATH * 2 - 1  'Unicodeでのパス最大長

Private Type FILETIME
    dwHighDateTime As Long
    dwLowDateTime As Long
End Type

'WIN32_FIND_DATA構造体(ディレクトリエントリ(ファイル情報))
Private Type WIN32_FIND_DATA
    dwFileAttributes            As Long             'ファイル属性
    ftCreationTime              As FILETIME         '作成日時
    ftLastAccessTime            As FILETIME         '最終アクセス日時
    ftLastWriteTime             As FILETIME         '最終更新日時
    nFileSizeHigh               As Long             'ファイルサイズの上位32bit値
    nFileSizeLow                As Long             'ファイルサイズの下位32bit値
    dwReserved0                 As Long             '予約(現状なし)
    dwReserved1                 As Long             '予約(現状なし)
    cFileName(MAX_PATH * 2 - 1) As Byte             'ロングファイル名
    cAlternate(14 * 2 - 1)      As Byte             'ショートファイル名(8+3文字)
End Type


#If VBA7 And Win64 Then
    Private Declare PtrSafe Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As LongPtr, lpFindFileData As WIN32_FIND_DATA) As LongPtr
    Private Declare PtrSafe Function FindNextFile Lib "kernel32" Alias "FindNextFileW" (ByVal hFindFile As LongPtr, lpFindFileData As WIN32_FIND_DATA) As LongPtr
    Private Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As LongPtr) As LongPtr
    Private lngFindFileHandle       As LongPtr                 '[FindFirstFile]のハンドル
#Else
    Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
    Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileW" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
    Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
    Private lngFindFileHandle       As Long                 '[FindFirstFile]のハンドル
#End If
        
 Public Function Find(ByVal strSearchPath As String)
 
    Dim udtWin32FindData        As WIN32_FIND_DATA      '[WIN32_FIND_DATA]構造体(検索結果)
    Dim strSeachFullPath As String
    
    Find = ""
    
    '検索フルパス名を生成
    If strSearchPath Like "\\*" Then
        strSeachFullPath = "\\?\UNC\" & Mid$(strSearchPath, 3)
    Else
        strSeachFullPath = "\\?\" & strSearchPath
    End If
    
    '文字列に一致するファイルを検索し、WIN32_FIND_DATA構造体に値を代入
    lngFindFileHandle = FindFirstFile(StrPtr(strSeachFullPath), udtWin32FindData)
    
    '検索結果ファイルハンドルが無効な場合終了
    If lngFindFileHandle <> INVALID_HANDLE_VALUE Then
        Find = GetFile(udtWin32FindData)
    End If
    
    If Find = "" Then
         'ファイルハンドルをクローズ
        lngFindFileHandle = FindClose(lngFindFileHandle)
    End If
 
 End Function
 Public Function FindNext() As String
 
    Dim udtWin32FindData        As WIN32_FIND_DATA      '[WIN32_FIND_DATA]構造体(検索結果)
    
    FindNext = ""
    
    If FindNextFile(lngFindFileHandle, udtWin32FindData) Then
        FindNext = GetFile(udtWin32FindData)
    End If
 
    If FindNext = "" Then
         'ファイルハンドルをクローズ
        lngFindFileHandle = FindClose(lngFindFileHandle)
    End If
 
 End Function
        
Private Function GetFile(udtWin32FindData As WIN32_FIND_DATA) As String
    
    Dim strFindFileName As String               '検索結果ファイル名

    GetFile = ""

    strFindFileName = CStr(udtWin32FindData.cFileName)
    strFindFileName = Left$(strFindFileName, InStr(strFindFileName, vbNullChar) - 1)
    
    If strFindFileName <> "." And strFindFileName <> ".." Then
        'ファイルの属性がディレクトリの場合
        If udtWin32FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
        Else
            GetFile = strFindFileName
        End If
    End If

End Function

