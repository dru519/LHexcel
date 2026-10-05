VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFolderCreateJournal"
Option Explicit
' Core is authoritative for mutation and rollback; this journal holds the planned paths.
Private mDirectoryPaths As Collection
Public Sub Bind(ByVal paths As Collection): Set mDirectoryPaths = paths: End Sub
Public Property Get DirectoryPaths() As Collection: Set DirectoryPaths = mDirectoryPaths: End Property
Public Property Get IsCoreOwned() As Boolean: IsCoreOwned = True: End Property
