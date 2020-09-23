Attribute VB_Name = "Module1"
Option Explicit

Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'ITEMIDLIST

   Public Const MAX_PATH = 260
   Public Const MAXDWORD = &HFFFF
   Public Const INVALID_HANDLE_VALUE = -1
   Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
   Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
   Public Const FILE_ATTRIBUTE_HIDDEN = &H2
   Public Const FILE_ATTRIBUTE_NORMAL = &H80
   Public Const FILE_ATTRIBUTE_READONLY = &H1
   Public Const FILE_ATTRIBUTE_SYSTEM = &H4
   Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

   Type FILETIME
     dwLowDateTime As Long
     dwHighDateTime As Long
   End Type

   Type WIN32_FIND_DATA
     dwFileAttributes As Long
     ftCreationTime As FILETIME
     ftLastAccessTime As FILETIME
     ftLastWriteTime As FILETIME
     nFileSizeHigh As Long
     nFileSizeLow As Long
     dwReserved0 As Long
     dwReserved1 As Long
     cFileName As String * MAX_PATH
     cAlternate As String * 14
   End Type

   Type SYSTEMTIME
     wYear As Integer
     wMonth As Integer
     wDayOfWeek As Integer
     wDay As Integer
     wHour As Integer
     wMinute As Integer
     wSecond As Integer
     wMilliseconds As Integer
   End Type

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000





Function BrowseFolder(hwnd As Long, szDialogTitle As String) As String
    Dim X As Long, BI As BROWSEINFO, dwIList As Long, szPath As String, wPos As Integer
    
    
    BI.hOwner = hwnd

    BI.lpszTitle = szDialogTitle
    BI.ulFlags = BIF_RETURNONLYFSDIRS
    dwIList = SHBrowseForFolder(BI)
    szPath = Space$(512)
    X = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    If X Then
        wPos = InStr(szPath, Chr(0))
        BrowseFolder = Left$(szPath, wPos - 1)
    Else
        BrowseFolder = ""
    End If

End Function


   
   Public Function StripNulls(OriginalStr As String) As String
      If (InStr(OriginalStr, Chr(0)) > 0) Then
         OriginalStr = Left(OriginalStr, _
          InStr(OriginalStr, Chr(0)) - 1)
      End If
      StripNulls = OriginalStr
   End Function
Public Sub QuickSortStringsAscending(sarray() As String, inLow As Long, inHi As Long)
  
   Dim pivot As String
   Dim tmpSwap As String
   Dim tmpLow As Long
   Dim tmpHi As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = sarray(1, (inLow + inHi) / 2)
  
   While (tmpLow <= tmpHi)
   
      While (sarray(1, tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot < sarray(1, tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = sarray(0, tmpLow): sarray(0, tmpLow) = sarray(0, tmpHi): sarray(0, tmpHi) = tmpSwap
         tmpSwap = sarray(1, tmpLow): sarray(1, tmpLow) = sarray(1, tmpHi): sarray(1, tmpHi) = tmpSwap
         tmpSwap = sarray(2, tmpLow): sarray(2, tmpLow) = sarray(2, tmpHi): sarray(2, tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   
   Wend
  
   If (inLow < tmpHi) Then QuickSortStringsAscending sarray(), inLow, tmpHi
   If (tmpLow < inHi) Then QuickSortStringsAscending sarray(), tmpLow, inHi
  
End Sub

