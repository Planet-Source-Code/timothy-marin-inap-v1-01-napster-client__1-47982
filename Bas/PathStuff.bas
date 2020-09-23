Attribute VB_Name = "PathStuff"
      Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" _
  (ByVal hWndOwner As Long, ByVal nFolder As Long, _
  pidl As ITEMIDLIST) As Long

Declare Function SHGetPathFromIDList Lib "Shell32.dll" _
  Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
  ByVal pszPath As String) As Long
  
Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

      Private Const BIF_RETURNONLYFSDIRS = 1
      Private Const BIF_DONTGOBELOWDOMAIN = 2
      Private Const MAX_PATH = 260

      Private Declare Function SHBrowseForFolder Lib "shell32" _
                                        (lpbi As BrowseInfo) As Long
                                        
      Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                        (ByVal lpString1 As String, ByVal _
                                        lpString2 As String) As Long
Public Type SH_ITEMID
    cb As Long
    abID As Byte
End Type
Dim zux As Integer
Public Type ITEMIDLIST
    mkid As SH_ITEMID
End Type

      Private Type BrowseInfo
         hWndOwner      As Long
         pIDLRoot       As Long
         pszDisplayName As Long
         lpszTitle      As Long
         ulFlags        As Long
         lpfnCallback   As Long
         lParam         As Long
         iImage         As Long
      End Type
      
Public Function FileExist(strFile As String) As Boolean
    If PathFileExists(strFile) = 1 Then
        FileExist = True
    ElseIf PathFileExists(strFile) = 0 Then
        FileExist = False
    End If
End Function

Public Function fGetSpecialFolder(CSIDL As Long) As String
    Dim sPath As String
    Dim IDL As ITEMIDLIST
    '
    ' Retrieve info about system folders such as the
    ' "Desktop" folder. Info is stored in the IDL structure.
    '
    fGetSpecialFolder = ""
    If SHGetSpecialFolderLocation(frmMain.hwnd, CSIDL, IDL) = 0 Then
        '
        ' Get the path from the ID list, and return the folder.
        '
        sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
            fGetSpecialFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1) & "\"
        End If
    End If
End Function

Public Function SelectFolder(coosh As Long)
      'Opens a Treeview control that displays the directories in a computer

         Dim lpIDList As Long
         Dim sBuffer As String
         Dim szTitle As String
         Dim tBrowseInfo As BrowseInfo

         szTitle = "Select Folder"
         With tBrowseInfo
            .hWndOwner = coosh
            .lpszTitle = lstrcat(szTitle, "")
            .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
         End With

         lpIDList = SHBrowseForFolder(tBrowseInfo)

         If (lpIDList) Then
            sBuffer = Space(MAX_PATH)
            SHGetPathFromIDList lpIDList, sBuffer
            sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
            SelectFolder = sBuffer
         End If
         
      End Function



Public Sub DirSize(ByVal path As String)
On Error Resume Next
Dim MyName  As String
 Dim MyDirNr As Long
 Dim MyDir() As String
 Dim a       As Long
 path = IIf(Right$(path, 1) = "\", path, path + "\")
 MyName = Dir$(path + "*.*", vbDirectory + vbArchive + vbReadOnly)
chklen = Split(path, "\")
 Do While (MyName <> "")
   zux = zux + 1
   If zux > 50 Then
       DoEvents
       zux = 0
   End If
   pA = path
   If MyName <> "." And MyName <> ".." Then
     If (GetAttr(path & MyName) And vbDirectory) <> vbDirectory Then
       Dim sizeX As Long
       sizeX = FileLen(path & MyName)
       If InStr(path & MyName, "$") Then GoTo ToSmall
       If sizeX < 1 Then GoTo ToSmall
       If Right(LCase(MyName), 3) = "mp3" Then
            frmShare.lblShare.Caption = MyName
            MP3FileName = path & MyName
            ReadHeader CStr(MP3FileName)
            frmMain.LVS.ListItems.Add , path & MyName, MyName, , 1
            frmMain.LVS.ListItems(path & MyName).SubItems(1) = Mid(Left(path, Len(path) - 1), InStrRev(Left(path, Len(path) - 1), "\") + 1, Len(Left(path, Len(path) - 1)))
            frmMain.LVS.ListItems(path & MyName).SubItems(2) = MP3HeaderInfo.Duration
            frmMain.LVS.ListItems(path & MyName).SubItems(3) = MP3HeaderInfo.Bitrate
            frmMain.LVS.ListItems(path & MyName).Tag = """" & path & MyName & """ 00000000000000000000000000000000 " & (FileLen(MP3FileName) - 1) & " " & MP3HeaderInfo.Bitrate & " " & MP3HeaderInfo.Frequency & " " & MP3HeaderInfo.PlayTime
       End If
ToSmall:
     Else
       ReDim Preserve MyDir(MyDirNr + 1)
       MyDirNr = MyDirNr + 1
       MyDir(MyDirNr) = MyName
     End If
   End If
   MyName = Dir
 Loop
    For a = 1 To MyDirNr
     DirSize path + MyDir(a) + "\"
   Next
 End Sub


