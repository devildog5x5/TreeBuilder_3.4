Attribute VB_Name = "DoesFileExist"
Option Explicit

  Public Const INVALID_HANDLE_VALUE = -1
  Public Const MAX_PATH = 260

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

  Declare Function FindFirstFile Lib "kernel32" _
      Alias "FindFirstFileA" _
      (ByVal lpFileName As String, _
      lpFindFileData As WIN32_FIND_DATA) As Long

  Declare Function FindClose Lib "kernel32" _
      (ByVal hFindFile As Long) As Long

 Public Function FileExists(sSource As String) As Boolean

   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   
   hFile = FindFirstFile(sSource, WFD)
   FileExists = hFile <> INVALID_HANDLE_VALUE
   
   Call FindClose(hFile)
   
 End Function

